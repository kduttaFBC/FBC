"""Final Audited Portfolio Summary — verified_portfolio_summary_record
(client-confirmed 2026-08-10).

Three layers:

    Original Portfolio Summary        what the lender reported      (PORT-01)
    Audited Portfolio Summary         what the audit calculated
    Final Audited Portfolio Summary   what the reviewer verified    (here)

THE HEADLINE TEST is equivalence with the historical workflow. Before this
feature the client did:

    Manual Review -> corrected workbook -> re-upload -> PORT-01 Portfolio Summary

and that result is the one the business has trusted. The new workflow must
reproduce it from the original workbook plus the reviewer's decisions, with no
round trip. MISSION VALLEY BANK is the primary ground truth, asserted field by
field to the cent.

Mission Valley qualifies uniquely among the eight banks, measured across every
historical vF corrected workbook:
  * 0 financial edits between the original and vF (so nothing the new workflow
    is not allowed to change),
  * 0 rows removed,
  * 10 of 10 client changes correspond to an engine recommendation, so every
    one is expressible as a review decision,
  * 1 engine recommendation the client did NOT take, exercising the reject path,
  * and a LIQUIDATION -> ACTIVE Loan Status change.

Community Bank and Trust is the secondary case: highest correction volume and
the only bank exercising closed-at-zero suppression. It has two historical
changes (GP 6536029107, both fields) that the engine never proposed, so there
is no review card for them. That gap is DOCUMENTED and asserted as a known
divergence rather than papered over with an artificial exact match.

Lendistry 2025 and Millennium are null controls: no closed-at-zero loans and
no removals, so suppression must move nothing at 1,234-loan scale.

The real-file tests skip unless the fixtures are present. Run with:
    FALCONBRIDGE_SOUP_XLSX=... FALCONBRIDGE_MEMO_XLSX=... \\
    FALCONBRIDGE_SRF_DIR=... FALCONBRIDGE_VF_DIR=... \\
    python -m pytest test_verified_portfolio_summary.py
"""

import os
import unittest

import pandas as pd

import checks
import corrections
import file_readers


SOUP_PATH = os.environ.get("FALCONBRIDGE_SOUP_XLSX")
MEMO_PATH = os.environ.get("FALCONBRIDGE_MEMO_XLSX")
SRF_DIR = os.environ.get("FALCONBRIDGE_SRF_DIR")      # original June submissions
VF_DIR = os.environ.get("FALCONBRIDGE_VF_DIR")        # historical corrected (vF) workbooks
REAL = all([SOUP_PATH, MEMO_PATH, SRF_DIR, VF_DIR])

FIELDS = [f"{c} ($)" for c in checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER] + ["Outstanding ($)"]
PCT_FIELDS = [f"{c} (%)" for c in checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER] + ["Outstanding (%)"]


# --------------------------------------------------------------------------
# Unit tests -- no fixtures required
# --------------------------------------------------------------------------

def _srf(rows):
    return pd.DataFrame([{
        "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}", "Loan Status": ls,
        "Delinquency Status": dq, "Beginning Balance": bb, "Ending Balance": eb,
        "Payment Received": rec, "Scheduled Payment": 500.0, "Servicing Fee": 0.0,
    } for gp, ls, dq, bb, eb, rec in rows])


SAMPLE = _srf([
    ("1", "ACTIVE", "CURRENT", 100000.0, 99000.0, 500.0),
    ("2", "ACTIVE", "CURRENT", 50000.0, 50000.0, 0.0),
    ("3", "LIQUIDATION", "120 DAYS", 25000.0, 25000.0, 0.0),
])


class NoOverridesIsTheIdentityCaseTests(unittest.TestCase):
    """With nothing verified, the Final summary IS the Original summary."""

    def test_empty_overrides_reproduces_port01_exactly(self):
        original = checks.portfolio_summary_check(SAMPLE, "T")["records"][0]
        final = checks.verified_portfolio_summary_record(SAMPLE, {}, "T")
        self.assertEqual(final, original)

    def test_none_overrides_is_treated_as_empty(self):
        self.assertEqual(checks.verified_portfolio_summary_record(SAMPLE, None, "T"),
                         checks.portfolio_summary_check(SAMPLE, "T")["records"][0])

    def test_the_aging_lambda_matches_port01s(self):
        # Not audited_status_by_loan: that would reinstate the audit's answer
        # on loans the reviewer rejected.
        for dq in ("CURRENT", "30 DAYS", "90 DAYS"):
            with self.subTest(dq=dq):
                df = _srf([("1", "ACTIVE", dq, 100.0, 100.0, 0.0)])
                self.assertEqual(checks.verified_portfolio_summary_record(df, {}, "T"),
                                 checks.portfolio_summary_check(df, "T")["records"][0])


class OverridesMoveTheRightBucketsTests(unittest.TestCase):

    def test_delinquency_override_moves_the_aging_bucket(self):
        out = checks.verified_portfolio_summary_record(
            SAMPLE, {("1", "Delinquency Status"): "60 DAYS"}, "T")
        self.assertEqual(out["Current ($)"], 50000.0)     # loan 2 only
        self.assertEqual(out["60 Days ($)"], 99000.0)     # loan 1 moved

    def test_loan_status_override_moves_a_loan_out_of_outstanding(self):
        base = checks.verified_portfolio_summary_record(SAMPLE, {}, "T")
        out = checks.verified_portfolio_summary_record(
            SAMPLE, {("1", "Loan Status"): "LIQUIDATION"}, "T")
        self.assertEqual(base["Outstanding ($)"], 149000.0)
        self.assertEqual(out["Outstanding ($)"], 50000.0)          # loan 1 excluded
        self.assertEqual(out["Liquidation ($)"], 125000.0)         # + its BEGINNING balance

    def test_loan_status_override_into_active_returns_a_loan_to_outstanding(self):
        # The LIQUIDATION -> ACTIVE shape Mission Valley exercises.
        out = checks.verified_portfolio_summary_record(
            SAMPLE, {("3", "Loan Status"): "ACTIVE"}, "T")
        self.assertEqual(out["Liquidation ($)"], 0.0)
        self.assertEqual(out["Outstanding ($)"], 174000.0)

    def test_loan_status_override_into_deferred(self):
        out = checks.verified_portfolio_summary_record(
            SAMPLE, {("2", "Loan Status"): "DEFERRED"}, "T")
        self.assertEqual(out["Deferred ($)"], 50000.0)

    def test_not_making_payment_is_unaffected_by_status_overrides(self):
        # It keys on Payment Received alone, which review cannot change.
        base = checks.verified_portfolio_summary_record(SAMPLE, {}, "T")
        out = checks.verified_portfolio_summary_record(
            SAMPLE, {("2", "Loan Status"): "DEFERRED",
                     ("2", "Delinquency Status"): "90 DAYS"}, "T")
        self.assertEqual(out["Not Making Payment ($)"], base["Not Making Payment ($)"])


class OverrideMapRobustnessTests(unittest.TestCase):

    def test_unknown_gp_is_ignored(self):
        self.assertEqual(
            checks.verified_portfolio_summary_record(SAMPLE, {("9999", "Loan Status"): "DEFERRED"}, "T"),
            checks.portfolio_summary_check(SAMPLE, "T")["records"][0])

    def test_the_input_frame_is_never_mutated(self):
        grouped = checks.consolidate_cycle_postings(SAMPLE)
        before = grouped["loan_status"].tolist()
        checks._apply_verified_overrides(grouped, {("1", "Loan Status"): "DEFERRED"})
        self.assertEqual(grouped["loan_status"].tolist(), before)


class FinancialFieldsAreUnreachableTests(unittest.TestCase):
    """CLIENT-CONFIRMED (2026-08-10, AMENDED 2026-08-15): Manual Review may not
    edit financial values GENERALLY.

    SUPERSEDED IN ONE NARROW CASE. The override domain widened from two fields to
    five on 2026-08-15, for a reviewer-approved reversal of a lender-reported
    LIQUIDATION back to ACTIVE: Ending Balance, Liquidation Date and Liquidation
    Loss became writable, because approving the reversal without them left the
    loan ACTIVE at $0.00 with a live Liquidation Date and understated the
    Portfolio Summary's Outstanding Balance.

    THE BARRIER DID NOT MOVE, IT MOVED LAYERS. Being in the mapping is no longer
    what makes a field unreachable -- corrections.py's gate is. These tests now
    assert BOTH halves: the mapping is exactly five fields, and the three new ones
    are reachable ONLY through that gate. Everything genuinely financial and
    unreviewable -- Beginning Balance, Payment Received, Scheduled Payment,
    Servicing Fee -- is still absent from the mapping entirely."""

    def test_genuinely_unreviewable_financial_fields_still_have_no_mapping(self):
        """Unchanged from 2026-08-10. Ending Balance deliberately absent from this
        list now; the four below remain unreachable by any path."""
        for field in ("Beginning Balance", "Payment Received",
                      "Scheduled Payment", "Servicing Fee"):
            with self.subTest(field=field):
                self.assertNotIn(field, checks._VERIFIED_OVERRIDE_COLUMNS)

    def test_the_override_domain_is_exactly_these_five_fields(self):
        self.assertEqual(set(checks._VERIFIED_OVERRIDE_COLUMNS),
                         {"Loan Status", "Delinquency Status", "Ending Balance",
                          "Liquidation Date", "Liquidation Loss"})

    def test_an_unreviewable_financial_override_is_still_silently_ignored(self):
        out = checks.verified_portfolio_summary_record(
            SAMPLE, {("1", "Payment Received"): 0.0,
                     ("1", "Scheduled Payment"): 1.0,
                     ("1", "Beginning Balance"): 1.0}, "T")
        self.assertEqual(out, checks.portfolio_summary_check(SAMPLE, "T")["records"][0])

    def test_a_none_ending_balance_override_never_zeroes_a_real_balance(self):
        """The defensive guard. None means NO DECISION WAS MADE, never "set to
        nothing" -- so an override map assembled from undecided candidates cannot
        destroy a balance. Liquidation Loss is deliberately the opposite: there,
        None IS the cleared value."""
        out = checks.verified_portfolio_summary_record(
            SAMPLE, {("1", "Ending Balance"): None}, "T")
        self.assertEqual(out, checks.portfolio_summary_check(SAMPLE, "T")["records"][0])

    def test_none_still_clears_liquidation_loss(self):
        grouped = checks.consolidate_cycle_postings(SAMPLE)
        applied = checks._apply_verified_overrides(
            grouped, {("1", "Liquidation Loss"): None})
        self.assertIsNone(applied.loc["1"]["liquidation_loss"])

    def test_the_three_new_fields_are_not_generally_reviewer_editable(self):
        """Reachability, not mapping: on an ordinary loan the review page offers
        no balance decision at all, so there is no id a reviewer could decide and
        no path to these three fields."""
        import corrections
        results = checks.run_servicer_reconciliation(
            "T", pd.DataFrame(), SAMPLE, pd.DataFrame(columns=["SBA GP Number"]),
            memo_df=None, memo_lender_key="T")
        stat01 = next(r for r in results if r["id"].split("::")[0] == "STAT-01")
        recs, _cands, _skipped = corrections.build_recommendations(stat01, "T")
        offered = [r for r in recs if r["field"] in
                   ("Ending Balance", "Liquidation Date", "Liquidation Loss")]
        self.assertEqual(offered, [])

    def test_the_balance_field_is_only_writable_with_a_server_side_restore_value(self):
        """corrections.py's gate, asserted here too so this file records where the
        barrier actually lives now."""
        import corrections
        ok, _reason = corrections.validate_approved_value(
            "Ending Balance", 1.0, balance_restore_value=None)
        self.assertFalse(ok)


class SuppressionTests(unittest.TestCase):

    def test_excluded_gps_are_removed_from_every_category(self):
        out = checks.verified_portfolio_summary_record(SAMPLE, {}, "T", excluded_gps=["3"])
        self.assertEqual(out["Liquidation ($)"], 0.0)
        self.assertEqual(out["120 Days ($)"], 0.0)
        self.assertEqual(out["Not Making Payment ($)"], 50000.0)   # loan 2 only

    def test_suppression_and_overrides_compose(self):
        out = checks.verified_portfolio_summary_record(
            SAMPLE, {("1", "Delinquency Status"): "30 DAYS"}, "T", excluded_gps=["3"])
        self.assertEqual(out["30 Days ($)"], 99000.0)
        self.assertEqual(out["Liquidation ($)"], 0.0)


# --------------------------------------------------------------------------
# Real-file equivalence
# --------------------------------------------------------------------------

BANKS = {
    "Mission Valley Bank": ("Mission Valley Bank", "Mission Valley Bank",
                            "2026.06_Mission Valley Bank_SOUP 2024-1 SRF Template_June 2026.xlsx",
                            "2026.06_Mission Valley Bank_SOUP 2024-1 SRF Template_June 2026 - vF_2026.07.15.xlsx"),
    "Community Bank and Trust": ("Community Bank & Trust", "Community Bank and Trust",
                                 "2026.06_Community Bank and Trust_SOUP 2025-1 SRF Template_June 2026.xlsx",
                                 "2026.06_Community Bank and Trust_SOUP 2025-1 SRF Template_June 2026 vF_2026.07.21.xlsx"),
    "Lendistry 2025": ("Lendistry", "Lendistry 2025",
                       "2026.06_Lendistry_SOUP 2025-1 SRF Template_June 2026.xlsx",
                       "2026.06_Lendistry_SOUP 2025-1 SRF Template_June 2026 vF_2026.07.21.xlsx"),
    "Millennium 2025": ("Millennium Bank", "Millennium 2025",
                        "2026.06_Millennium_SOUP 2025-1 SFR Template_June 2026.xlsx",
                        "2026.06_Millennium_SOUP 2025-1 SFR Template_June 2026 vF_2026.07.16.xlsx"),
}


def _load(label):
    soup_key, display, orig, vf = BANKS[label]
    soup = file_readers.read_soup_file(open(SOUP_PATH, "rb").read())
    memo = file_readers.read_service_memo(open(MEMO_PATH, "rb").read())
    fname = f"{soup_key}||{display}__x.xlsx"
    sk, dl, lt, si, deal = file_readers.read_servicer_file(
        open(os.path.join(SRF_DIR, orig), "rb").read(), filename=fname)
    _, _, _, si_vf, _ = file_readers.read_servicer_file(
        open(os.path.join(VF_DIR, vf), "rb").read(), filename=fname)
    return soup, memo, sk, dl, lt, si, si_vf, deal, display


def _reconstruct_decisions(si_original, si_vf, candidates):
    """Turns the historical original -> vF diff into review decisions.

    A change the engine recommended becomes an approval (a SUBSTITUTION when
    the client's value differs from the engine's, which is how "the client
    modified the recommendation" is expressed). A recommendation the client did
    NOT take becomes a rejection. A client change with no recommendation cannot
    be expressed at all -- returned separately as the documented gap."""
    a = checks.consolidate_cycle_postings(si_original)
    b = checks.consolidate_cycle_postings(si_vf)
    a.index = a.index.astype(str).str.strip()
    b.index = b.index.astype(str).str.strip()

    # AMENDED 2026-08-15. _VERIFIED_OVERRIDE_COLUMNS gained three fields, two of
    # them MONEY (Ending Balance, Liquidation Loss). The original comparison below
    # was a string compare, which is correct for a status label and WRONG for an
    # amount: the initial submissions carry unrounded floats where the vF files
    # carry 2dp values, so "276702.0999999" != "276702.1" was read as a client
    # change on dozens of loans that the client never touched. Amounts are now
    # compared to the CENT, which is what "did the client change this figure"
    # actually means.
    MONEY_FIELDS = {"Ending Balance", "Liquidation Loss"}

    def _same(field, old_value, new_value):
        if field in MONEY_FIELDS:
            def cents(v):
                if v is None:
                    return None
                try:
                    amount = float(v)
                except (TypeError, ValueError):
                    return None
                if amount != amount:            # NaN
                    return None
                return int(round(amount * 100))
            return cents(old_value) == cents(new_value)
        return (str(old_value or "").strip().upper()
                == str(new_value or "").strip().upper())

    client = {}
    for gp in sorted(set(a.index) & set(b.index)):
        for field, column in checks._VERIFIED_OVERRIDE_COLUMNS.items():
            old_value = a.loc[gp].get(column)
            new_value = b.loc[gp].get(column)
            if not _same(field, old_value, new_value):
                client[(gp, field)] = (new_value if field in MONEY_FIELDS
                                       else str(new_value or "").strip())

    decisions, unexpressible = [], dict(client)
    for cand in candidates:
        key = (str(cand["gp"]).strip(), cand["field"])
        if key in client:
            decision = {"id": cand["id"], "decision": "approve"}
            if str(client[key]).strip().upper() != str(cand["corrected_value"]).strip().upper():
                decision["approved_value"] = client[key]
            decisions.append(decision)
            unexpressible.pop(key, None)
        else:
            decisions.append({"id": cand["id"], "decision": "reject"})
    return decisions, unexpressible


def _run_new_workflow(label):
    soup, memo, sk, dl, lt, si, si_vf, deal, display = _load(label)
    sdf = file_readers.filter_soup_by_servicer(soup, sk)
    results = checks.run_servicer_reconciliation(dl, lt, si, sdf, memo_df=memo,
                                                 memo_lender_key=sk, deal=deal)
    stat01 = next(r for r in results if r["id"].split("::")[0] == "STAT-01")
    excluded = [t["SBA GP Number"] for t in stat01.get("suppressed_closed_loans") or []]
    _, candidates, _ = corrections.build_recommendations(stat01, display)
    decisions, unexpressible = _reconstruct_decisions(si, si_vf, candidates)

    outcome = corrections.generate_reviewed_workbook(
        open(os.path.join(SRF_DIR, BANKS[label][2]), "rb").read(),
        BANKS[label][2], display, candidates, decisions,
        "Regression Harness", "sha", stat01.get("_cycle_month"))

    new = checks.verified_portfolio_summary_record(
        si, outcome["verified_overrides"], display, excluded_gps=excluded)
    reference = checks.portfolio_summary_check(si_vf, display)["records"][0]
    # candidates returned too (2026-08-15) so a test can distinguish "no decision
    # was offered" from "a decision was offered and the client did not take it".
    outcome["_candidates"] = candidates
    return new, reference, unexpressible, excluded, outcome


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR / _VF_DIR")
class MissionValleyEquivalenceTests(unittest.TestCase):
    """PRIMARY GROUND TRUTH. The new workflow must reproduce the historical
    corrected-workbook PORT-01 result field by field, to the cent."""

    @classmethod
    def setUpClass(cls):
        cls.new, cls.ref, cls.gap, cls.excluded, cls.outcome = _run_new_workflow("Mission Valley Bank")

    # AMENDED 2026-08-15. One client change is NOT expressible, and that is the
    # intended consequence of keeping the liquidation-reversal rule narrow.
    #
    # THE LOAN. GP 4778059109 is reported LIQUIDATION with Beginning Balance
    # 279,150.10 and Ending Balance 276,702.10 -- a REAL, NON-ZERO closing
    # balance. The client's vF set Loan Status ACTIVE, cleared the Liquidation
    # Date, and LEFT THE ENDING BALANCE ALONE at 276,702.10.
    #
    # WHY IT CANNOT BE EXPRESSED. Liquidation Date is never a decision of its own;
    # it is cleared only as a CONSEQUENCE of a reviewer explicitly restoring a
    # ZEROED Ending Balance. This loan's balance is not zero, so it receives no
    # balance decision at all (condition 4 of the gate -- deliberately, because
    # neither "restore to 279,150.10" nor "keep 0.00" is correct for a loan whose
    # 276,702.10 is right as submitted). No balance decision means no cascade,
    # which means no way to clear the date.
    #
    # RECORDED, NOT FIXED. Whether a liquidation reversal that does NOT change the
    # balance should still be able to clear the Liquidation Date is a separate
    # business-rule question and a client decision. Reported as a future decision
    # item; deliberately out of scope for this patch.
    EXPECTED_UNEXPRESSIBLE = {("4778059109", "Liquidation Date")}

    def test_the_only_unexpressible_change_is_the_documented_liquidation_date(self):
        self.assertEqual(set(self.gap), self.EXPECTED_UNEXPRESSIBLE)

    def test_the_gap_loan_correctly_receives_no_balance_decision(self):
        """The narrow gate at work: a reversal whose Ending Balance is already
        positive gets its Loan Status decision alone. Pins that the gap is a
        deliberate scope boundary, not a decision the engine failed to offer."""
        offered = {(str(c["gp"]).strip(), c["field"])
                   for c in self.outcome["_candidates"]}
        self.assertNotIn(("4778059109", "Ending Balance"), offered)
        self.assertIn(("4778059109", "Loan Status"), offered)

    def test_no_balance_decision_is_raised_anywhere_in_this_bank(self):
        """Mission Valley has no zero-balance liquidation reversal, so the whole
        feature must be inert here -- no candidate, no override, no cascade."""
        fields = {c["field"] for c in self.outcome["_candidates"]}
        self.assertNotIn("Ending Balance", fields)
        written = {f for (_gp, f) in self.outcome["verified_overrides"]}
        self.assertNotIn("Ending Balance", written)
        self.assertNotIn("Liquidation Date", written)
        self.assertNotIn("Liquidation Loss", written)

    def test_that_loans_positive_ending_balance_is_never_touched(self):
        """The hazard condition 4 removes: 276,702.10 must survive intact."""
        grouped = checks._apply_verified_overrides(
            checks.consolidate_cycle_postings(_load("Mission Valley Bank")[5]),
            self.outcome["verified_overrides"])
        self.assertAlmostEqual(
            float(grouped.loc["4778059109"]["ending_balance"]), 276702.10, places=2)

    def test_no_closed_at_zero_loans_to_suppress(self):
        self.assertEqual(self.excluded, [])

    def test_the_reject_path_was_exercised(self):
        self.assertGreater(self.outcome["rejected_count"], 0)

    def test_a_loan_status_correction_was_applied(self):
        applied = {f for (_gp, f) in self.outcome["verified_overrides"]}
        self.assertIn("Loan Status", applied)

    def test_every_dollar_field_matches_to_the_cent(self):
        for field in FIELDS:
            with self.subTest(field=field):
                self.assertAlmostEqual(self.new[field], self.ref[field], places=2,
                                       msg=f"{field}: new {self.new[field]} vs historical {self.ref[field]}")

    def test_every_percentage_field_matches(self):
        for field in PCT_FIELDS:
            with self.subTest(field=field):
                self.assertAlmostEqual(self.new[field], self.ref[field], places=2)

    def test_the_whole_record_is_equal(self):
        self.assertEqual(self.new, self.ref)


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR / _VF_DIR")
class CommunityBankVerifiedCorrectionsTests(unittest.TestCase):
    """SECONDARY. Verifies suppression plus approved Loan Status and
    Delinquency Status corrections. Community Bank has two historical changes
    with no Manual Review card, so exact equality is NOT asserted -- the gap is
    documented and bounded instead."""

    KNOWN_GAP = {("6536029107", "Loan Status"), ("6536029107", "Delinquency Status")}

    @classmethod
    def setUpClass(cls):
        cls.new, cls.ref, cls.gap, cls.excluded, cls.outcome = _run_new_workflow("Community Bank and Trust")

    def test_the_documented_gap_is_exactly_the_known_two_changes(self):
        # If this moves, the engine's coverage changed and the divergence below
        # must be re-derived rather than assumed.
        self.assertEqual(set(self.gap), self.KNOWN_GAP)

    def test_closed_at_zero_suppression_is_exercised(self):
        self.assertIn("4338609104", self.excluded)

    def test_the_suppressed_loan_contributes_to_nothing(self):
        # Its historical treatment: the client deleted it from the corrected
        # submission. Suppression reproduces that.
        with_loan = checks.verified_portfolio_summary_record(
            _load("Community Bank and Trust")[5], self.outcome["verified_overrides"],
            "Community Bank and Trust", excluded_gps=None)
        self.assertGreater(with_loan["Liquidation ($)"], self.new["Liquidation ($)"])

    def test_approved_loan_status_corrections_were_applied(self):
        applied = {f for (_gp, f) in self.outcome["verified_overrides"]}
        self.assertIn("Loan Status", applied)

    def test_approved_delinquency_status_corrections_were_applied(self):
        applied = {f for (_gp, f) in self.outcome["verified_overrides"]}
        self.assertIn("Delinquency Status", applied)

    # The divergence the documented gap causes, measured rather than assumed.
    # GP 6536029107 carries an Ending Balance of 407,893.05 and the client
    # changed BOTH its fields in the historical vF:
    #     Loan Status        DEFERRED  -> ACTIVE
    #     Delinquency Status 60 DAYS   -> DELINQUENT
    # Neither had a Manual Review card, so the new workflow cannot express
    # them and the loan keeps its lender-reported values. That lands its whole
    # balance in exactly three buckets, in a known direction.
    EXPECTED_DIVERGENCE = {
        "Delinquent ($)": -407893.05,   # the client's move, not made
        "60 Days ($)": +407893.05,      # still bucketed on the lender's value
        "Deferred ($)": +407893.05,     # still DEFERRED rather than ACTIVE
    }

    def test_the_divergence_is_exactly_the_documented_gap(self):
        # Every field either matches the historical result to the cent, or
        # differs by EXACTLY the gap loan's balance in the expected direction.
        # Nothing else may move.
        for field in FIELDS:
            expected = self.EXPECTED_DIVERGENCE.get(field, 0.0)
            with self.subTest(field=field):
                self.assertAlmostEqual(self.new[field] - self.ref[field], expected, places=2,
                                       msg=f"{field}: new {self.new[field]} vs historical "
                                           f"{self.ref[field]}")

    def test_exactly_three_fields_diverge_and_all_others_match(self):
        moved = {f for f in FIELDS if abs(self.new[f] - self.ref[f]) > 0.005}
        self.assertEqual(moved, set(self.EXPECTED_DIVERGENCE))
        self.assertEqual(len(FIELDS) - len(moved), 11)

    def test_the_divergence_is_one_loans_balance(self):
        for field, delta in self.EXPECTED_DIVERGENCE.items():
            with self.subTest(field=field):
                self.assertAlmostEqual(abs(delta), 407893.05, places=2)

    def test_outstanding_and_the_activity_metrics_are_unaffected(self):
        # DEFERRED -> ACTIVE does not cross the Outstanding exclusion (only
        # PAID-IN-FULL and LIQUIDATION do), and review cannot touch payments,
        # so these must match exactly despite the gap.
        for field in ("Outstanding ($)", "Not Making Payment ($)",
                      "Liquidation ($)", "Paid In Full ($)"):
            with self.subTest(field=field):
                self.assertAlmostEqual(self.new[field], self.ref[field], places=2)


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR / _VF_DIR")
class NullControlTests(unittest.TestCase):
    """No closed-at-zero loans and no historical row removals, so suppression
    must move nothing -- including at 1,234-loan scale."""

    def test_no_suppression_and_no_movement(self):
        for label in ("Lendistry 2025", "Millennium 2025"):
            with self.subTest(bank=label):
                soup, memo, sk, dl, lt, si, si_vf, deal, display = _load(label)
                sdf = file_readers.filter_soup_by_servicer(soup, sk)
                results = checks.run_servicer_reconciliation(dl, lt, si, sdf, memo_df=memo,
                                                             memo_lender_key=sk, deal=deal)
                stat01 = next(r for r in results if r["id"].split("::")[0] == "STAT-01")
                excluded = [t["SBA GP Number"] for t in stat01.get("suppressed_closed_loans") or []]
                self.assertEqual(excluded, [], f"{label} should have nothing to suppress")
                self.assertEqual(
                    checks.verified_portfolio_summary_record(si, {}, display, excluded_gps=excluded),
                    checks.portfolio_summary_check(si, display)["records"][0])


if __name__ == "__main__":
    unittest.main()
