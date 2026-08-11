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
    """CLIENT-CONFIRMED (2026-08-10): Manual Review may not edit financial
    values. Two independent barriers; this asserts the second."""

    def test_financial_fields_have_no_override_mapping(self):
        for field in ("Ending Balance", "Beginning Balance", "Payment Received",
                      "Scheduled Payment", "Servicing Fee"):
            with self.subTest(field=field):
                self.assertNotIn(field, checks._VERIFIED_OVERRIDE_COLUMNS)

    def test_only_the_two_status_fields_are_mapped(self):
        self.assertEqual(set(checks._VERIFIED_OVERRIDE_COLUMNS),
                         {"Loan Status", "Delinquency Status"})

    def test_a_financial_override_is_silently_ignored_not_applied(self):
        out = checks.verified_portfolio_summary_record(
            SAMPLE, {("1", "Ending Balance"): 1.0, ("1", "Payment Received"): 0.0}, "T")
        self.assertEqual(out, checks.portfolio_summary_check(SAMPLE, "T")["records"][0])


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

    client = {}
    for gp in sorted(set(a.index) & set(b.index)):
        for field, column in checks._VERIFIED_OVERRIDE_COLUMNS.items():
            old = str(a.loc[gp].get(column, "")).strip().upper()
            new = str(b.loc[gp].get(column, "")).strip().upper()
            if old != new:
                client[(gp, field)] = str(b.loc[gp].get(column, "")).strip()

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
    return new, reference, unexpressible, excluded, outcome


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR / _VF_DIR")
class MissionValleyEquivalenceTests(unittest.TestCase):
    """PRIMARY GROUND TRUTH. The new workflow must reproduce the historical
    corrected-workbook PORT-01 result field by field, to the cent."""

    @classmethod
    def setUpClass(cls):
        cls.new, cls.ref, cls.gap, cls.excluded, cls.outcome = _run_new_workflow("Mission Valley Bank")

    def test_every_client_change_was_expressible_as_a_review_decision(self):
        self.assertEqual(self.gap, {},
                         f"Mission Valley must have no unexpressible changes; got {self.gap}")

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
