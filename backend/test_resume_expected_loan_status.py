"""Expected Loan Status on a confirmed deferral RESUME cycle
(client-confirmed 2026-08-10).

On a cycle the Servicer Memo confirms as the deferral resume cycle, the
audit's Expected Loan Status is ACTIVE, DERIVED FROM THE MEMO -- never copied
from the servicer's own current Loan Status.

The branch previously did `expected_loan_status = srf_status`. That was the
last surviving path where a servicer's assertion validated itself: the audit
adopted the servicer's status as the expected one, so the two could never
disagree and no LOAN STATUS MISMATCH was possible. Exactly the reasoning
behind the SUPERSEDED 2026-08-08 block in checks.py for terminal collection
statuses -- this branch predates it and was missed.

THE INDEPENDENT EVIDENCE is the Servicer Memo's Regular Payment Resume Date:
a separate document from the SRF submission under audit, and the very thing
that puts the cycle in the `resume` phase. A loan whose confirmed deferral
window has ended and whose payment schedule has resumed is back in normal
servicing -- ACTIVE. The branch already trusts that same date to reset the
delinquency ladder to CURRENT, which is the more consequential act.

Found on GP 6717929105 (July 2026, Community Bank and Trust): servicer
reported LIQUIDATION on a confirmed resume cycle, it matched itself, no Loan
Status recommendation was produced, and the reviewer had no way to correct it.

WHAT THIS DOES NOT CHANGE, each proved below: the delinquency calculation,
deferral start/mid logic, the PAID-IN-FULL valid-exit-transition carve-out,
the BANKRUPTCY override, the still-reports-DEFERRED contradiction path, the
permitted Loan Status domains, or the reviewer write path.

Run with:
    python -m pytest test_resume_expected_loan_status.py
"""

import unittest

import pandas as pd

import checks
import corrections
import file_readers


GP = "9990000001"
SERVICER = "Millennium Bank"

# Deferral window 2026-01-01 -> 2026-05-01, regular payments resume 2026-06-01.
# So 2026-06-01 is the resume cycle and 2026-03-01 sits mid-window.
RESUME_CYCLE = "2026-06-01"
MID_CYCLE = "2026-03-01"

FULL = "6-month principal and interest deferment"
PARTIAL = "6-month principal only deferment"
UNCLASSIFIED = "three months deferment"   # states neither full nor principal-only


def _memo(action=FULL, gp=GP, resume="2026-06-01"):
    df = pd.DataFrame([{
        "Lender": SERVICER, "SBA Loan ID": gp, "Lender Loan ID": gp,
        "Borrower Name": "T", "Securitization Action": "Deferred",
        "Servicer Action": action,
        "Deferral Start Date": pd.Timestamp("2026-01-01"),
        "Deferral End Date": pd.Timestamp("2026-05-01"),
        "Regular Payment Resume Date": pd.NaT if resume is None else pd.Timestamp(resume),
        "Report Date": pd.Timestamp("2026-05-01"),
    }])
    df["gp_sba"] = df["SBA Loan ID"]
    df["gp_lender"] = df["Lender Loan ID"]
    df["deferral_type"] = df["Servicer Action"].apply(file_readers.classify_deferral_type)
    return df


def _calc(srf_status, action=FULL, cycle=RESUME_CYCLE, soup_status="DEFERRED",
          prior_dq="90 DAYS", received=500.0, scheduled=500.0, memo=None):
    """_calculate_expected_status -> (calculated_dq, expected_loan_status, phase)."""
    calc, expected, deferral, *_ = checks._calculate_expected_status(
        GP, srf_status, soup_status, prior_dq, received, scheduled, True,
        _memo(action) if memo is None else memo, SERVICER, pd.Timestamp(cycle),
        srf_dq="180 DAYS", servicing_fee=0.0)
    return calc, expected, (deferral or {}).get("phase")


# --- end-to-end fixtures ----------------------------------------------------
# SOUP carries history THROUGH the prior cycle; checks._infer_cycle_month
# derives the audited cycle as one month after SOUP's latest Payment Month.
# Latest here is 2026-05-01, so the audited cycle is the 2026-06-01 resume.

def _soup(loan_status="DEFERRED"):
    return pd.DataFrame([{
        "SBA GP Number": GP, "Servicer": SERVICER, "Payment Month": month,
        "Loan Status": loan_status, "Delinquency Status": "90 DAYS",
        "Beginning Actual Balance": 100000.0, "Ending Actual Balance": 100000.0,
    } for month in ("2026-04-01", "2026-05-01")])


def _srf(loan_status, dq="90 DAYS", received=0.0, scheduled=500.0):
    return pd.DataFrame([{
        "SBA GP Number": GP, "Lender Loan Number": f"L-{GP}",
        "Loan Status": loan_status, "Delinquency Status": dq,
        "Beginning Balance": 100000.0, "Ending Balance": 100000.0,
        "Payment Received": received, "Scheduled Payment": scheduled,
        "Servicing Fee": 0.0,
        # consolidate_cycle_postings builds _all_excel_rows from the per-row
        # _excel_row values; without one there is nowhere to write and
        # build_recommendations routes the candidate to `skipped` instead.
        "_excel_row": 12, "_sheet_name": "Servicer Inputs",
        "_header_excel_row": 4,
    }])


def _end_to_end(srf_loan_status, action=FULL, soup_status="DEFERRED", dq="90 DAYS"):
    """Returns (stat01_record, correction_data_entry, recommendations)."""
    result = checks.status_match_check(
        _srf(srf_loan_status, dq=dq), _soup(soup_status),
        memo_df=_memo(action), servicer_name=SERVICER)
    record = next((r for r in result["records"]
                   if str(r.get("SBA GP Number")).strip() == GP), None)
    entry = next((c for c in result.get("correction_data") or []
                  if str(c.get("gp")).strip() == GP), None)
    recs, _, _ = corrections.build_recommendations(result, SERVICER)
    return record, entry, recs


class ResumeExpectedStatusIsDerivedNotCopiedTests(unittest.TestCase):
    """The core rule: on a confirmed resume the expected status is ACTIVE,
    whatever the servicer reported."""

    def test_resume_plus_servicer_liquidation_expects_active(self):
        # THE DEFECT. Pre-fix this returned "LIQUIDATION" -- the servicer's
        # own claim, which then matched itself.
        calc, expected, phase = _calc("LIQUIDATION")
        self.assertEqual(phase, "resume")
        self.assertEqual(expected, "ACTIVE")

    def test_resume_plus_servicer_active_still_expects_active(self):
        # The legitimate resume. Same answer as before the fix, reached by
        # derivation instead of by copying.
        _, expected, phase = _calc("ACTIVE")
        self.assertEqual(phase, "resume")
        self.assertEqual(expected, "ACTIVE")

    def test_expected_status_is_independent_of_what_the_servicer_reports(self):
        # The property that actually matters: one memo, one cycle, one answer,
        # regardless of the field being compared against it. If any of these
        # echoed its input, the servicer would be validating itself again.
        for reported in ("ACTIVE", "LIQUIDATION", "FORECLOSURE", "REO",
                         "SUBSTITUTION-OUT", "SUBSTITUTION-IN", "Active", "active"):
            with self.subTest(reported=reported):
                _, expected, _ = _calc(reported)
                self.assertEqual(expected, "ACTIVE")

    def test_other_terminal_collection_statuses_behave_like_liquidation(self):
        for reported in ("FORECLOSURE", "REO"):
            with self.subTest(reported=reported):
                _, expected, phase = _calc(reported)
                self.assertEqual(phase, "resume")
                self.assertEqual(expected, "ACTIVE")

    def test_expected_status_is_canonically_spelled(self):
        # Pre-fix the servicer's own casing was echoed back ("Active"), so the
        # displayed Expected Loan Status varied by however the servicer typed
        # it. Derivation gives one canonical spelling.
        _, expected, _ = _calc("Active")
        self.assertEqual(expected, "ACTIVE")


class ResumeIsTypeAgnosticTests(unittest.TestCase):
    """The resume boundary is a property of the deferral WINDOW, not its type
    -- matching the branch's existing phase-not-type keying."""

    def test_full_partial_and_unclassified_all_expect_active_at_resume(self):
        for action in (FULL, PARTIAL, UNCLASSIFIED):
            with self.subTest(action=action):
                _, expected, phase = _calc("LIQUIDATION", action=action)
                self.assertEqual(phase, "resume")
                self.assertEqual(expected, "ACTIVE")

    def test_unclassified_reports_the_mismatch_but_produces_no_candidate(self):
        # Weak evidence is REPORTED, never written: an unclassified deferral
        # carries DEFERRAL DATA QUALITY WARNING, which corrections.
        # WARNING_REASONS excludes from candidacy entirely.
        record, entry, recs = _end_to_end("LIQUIDATION", action=UNCLASSIFIED)
        self.assertIn("DEFERRAL DATA QUALITY WARNING", entry["reasons"])
        self.assertEqual([r for r in recs if r["field"] == "Loan Status"], [])
        self.assertIn("DEFERRAL DATA QUALITY WARNING", corrections.WARNING_REASONS)


class UntouchedCarveOutTests(unittest.TestCase):
    """Three arrivals are handled independently of expected_loan_status and
    must keep behaving exactly as before."""

    def test_paid_in_full_at_resume_still_matches(self):
        # is_valid_exit_transition short-circuits the match test BEFORE the
        # comparison (checks.py: _VALID_EXIT_TRANSITION_KEYWORDS), so the
        # expected value never gets a vote. Arriving at PAID-IN-FULL is
        # normal lifecycle, not a mismatch.
        for reported in ("PAID-IN-FULL", "PIF", "Paid-in-Full"):
            with self.subTest(reported=reported):
                record, entry, recs = _end_to_end(reported)
                self.assertEqual(record["Loan Status Match"], "YES")
                self.assertNotIn("LOAN STATUS MISMATCH", entry["reasons"])
                self.assertEqual([r for r in recs if r["field"] == "Loan Status"], [])

    def test_paid_in_full_never_reaches_the_resume_branch_at_all(self):
        # PAID-IN-FULL is protected TWICE over, and the outer guard fires
        # first: is_terminal_this_cycle short-circuits ahead of every deferral
        # branch, so the memo is never consulted, no phase is assigned, and
        # the changed line is not executed. Asserted explicitly so a future
        # change to either layer is not mistaken for the other.
        calc, expected, deferral, _, _, is_terminal, *_ = \
            checks._calculate_expected_status(
                GP, "PAID-IN-FULL", "DEFERRED", "90 DAYS", 500.0, 500.0, True,
                _memo(), SERVICER, pd.Timestamp(RESUME_CYCLE),
                srf_dq="180 DAYS", servicing_fee=0.0)
        self.assertTrue(is_terminal)
        self.assertIsNone((deferral or {}).get("phase"))
        self.assertIsNone(expected)

    def test_bankruptcy_at_resume_is_still_accepted_as_authoritative(self):
        # The `if "BANKRUPT" in srf_status_u` override runs AFTER this branch
        # and wins as the final word -- unchanged by this fix.
        for reported in ("BANKRUPTCY", "BANKRUPT"):
            with self.subTest(reported=reported):
                _, expected, phase = _calc(reported)
                self.assertEqual(phase, "resume")
                self.assertEqual(expected, reported)

    def test_bankruptcy_at_resume_matches_end_to_end(self):
        record, entry, recs = _end_to_end("BANKRUPTCY")
        self.assertEqual(record["Loan Status Match"], "YES")
        self.assertNotIn("LOAN STATUS MISMATCH", entry["reasons"])

class ServicerStillReportsDeferredTests(unittest.TestCase):
    """SUPERSEDED 2026-08-10 (checks.py site 4). This sub-branch used to set
    expected_loan_status = None and raise a DEFERRAL DATA QUALITY WARNING. Both
    were wrong: None fell through to SOUP's stale still-DEFERRED prior so the
    servicer's claim matched itself, and the warning claimed the memo dates
    could not place the cycle when they place it exactly. It is a Loan Status
    contradiction, not a deferral-evidence defect."""

    def test_expected_is_active_not_none(self):
        _, expected, phase = _calc("DEFERRED")
        self.assertEqual(phase, "resume")
        self.assertEqual(expected, "ACTIVE")

    def test_no_deferral_data_quality_warning_is_raised_from_this_path(self):
        _, _, _, memo_warning, warning_type, *_ = checks._calculate_expected_status(
            GP, "DEFERRED", "DEFERRED", "90 DAYS", 500.0, 500.0, True,
            _memo(), SERVICER, pd.Timestamp(RESUME_CYCLE),
            srf_dq="180 DAYS", servicing_fee=0.0)
        self.assertIsNone(warning_type)
        self.assertFalse(memo_warning)

    def test_it_becomes_a_loan_status_mismatch_end_to_end(self):
        record, entry, _ = _end_to_end("DEFERRED")
        self.assertEqual(record["Expected Loan Status"], "ACTIVE")
        self.assertEqual(record["Loan Status Match"], "NO")
        self.assertIn("LOAN STATUS MISMATCH", entry["reasons"])
        self.assertNotIn("DEFERRAL DATA QUALITY WARNING", entry["reasons"])
        self.assertEqual(record["Deferral Data Quality Warning"], "")

    def test_the_delinquency_correction_is_no_longer_suppressed(self):
        # The warning was in WARNING_REASONS, so it excluded the WHOLE loan
        # from candidacy -- silently killing an unrelated, deterministic DQ
        # correction on the grounds that the evidence could not be trusted.
        _, _, recs = _end_to_end("DEFERRED", dq="CURRENT")
        dq = [r for r in recs if r["field"] == "Delinquency Status"]
        self.assertEqual(len(dq), 1)
        self.assertEqual(dq[0]["reported"], "CURRENT")
        self.assertIsNone(dq[0]["reviewFlag"])   # ordinary calculation, nothing to weigh

    def test_a_review_flagged_deferred_to_active_recommendation_is_produced(self):
        _, _, recs = _end_to_end("DEFERRED")
        ls = [r for r in recs if r["field"] == "Loan Status"]
        self.assertEqual(len(ls), 1)
        self.assertEqual(ls[0]["reported"], "DEFERRED")
        self.assertEqual(ls[0]["calculated"], "ACTIVE")
        self.assertEqual(ls[0]["reviewFlag"],
                         corrections.REVIEW_REASON_RESUME_CONTRADICTION)

    def test_the_contradiction_text_is_not_lost_only_relocated(self):
        # It moved from a suppressing warning column onto the control where the
        # reviewer acts.
        self.assertIn("still reports Loan Status = DEFERRED",
                      corrections.REVIEW_REASON_RESUME_CONTRADICTION)

    def test_exiting_deferred_is_never_auto_applied(self):
        # The v1 "never auto-exit DEFERRED" scope rule is superseded, but only
        # for PROPOSING -- review_required is unconditionally True here.
        eligible = [e for e in corrections.build_eligible_corrections(
            checks.status_match_check(_srf("DEFERRED"), _soup(), memo_df=_memo(),
                                      servicer_name=SERVICER)["correction_data"])
            if e["field"] == "Loan Status"]
        self.assertEqual(len(eligible), 1)
        self.assertTrue(eligible[0]["review_required"])

    def test_only_one_loan_status_candidate_is_emitted(self):
        # The three Loan Status branches must be mutually exclusive: the
        # "-> DEFERRED" branch needs expected_raw == DEFERRED and phase
        # start/mid; the terminal branch needs a collection status.
        eligible = corrections.build_eligible_corrections(
            checks.status_match_check(_srf("DEFERRED"), _soup(), memo_df=_memo(),
                                      servicer_name=SERVICER)["correction_data"])
        self.assertEqual(len([e for e in eligible if e["field"] == "Loan Status"]), 1)

    def test_undocumented_evidence_does_not_get_a_candidate(self):
        # The branch is gated on documented evidence because it proposes
        # EXITING a deferral -- the one direction v1 refused.
        record, entry, recs = _end_to_end("DEFERRED", action=UNCLASSIFIED)
        self.assertEqual(record["Loan Status Match"], "NO")
        self.assertIn("LOAN STATUS MISMATCH", entry["reasons"])
        self.assertEqual([r for r in recs if r["field"] == "Loan Status"], [])

    def test_the_gate_is_evidence_status_not_merely_the_phase(self):
        for entry_evidence, expect in (("documented", 1), ("date-anomaly", 0),
                                       ("type-unclassified", 0), ("lookup-failed", 0)):
            with self.subTest(evidence=entry_evidence):
                cd = checks.status_match_check(_srf("DEFERRED"), _soup(), memo_df=_memo(),
                                               servicer_name=SERVICER)["correction_data"]
                cd[0]["evidence_status"] = entry_evidence
                got = [e for e in corrections.build_eligible_corrections(cd)
                       if e["field"] == "Loan Status"]
                self.assertEqual(len(got), expect)


class LegitimateWarningsArePreservedTests(unittest.TestCase):
    """Only checks.py's site 4 stopped raising DEFERRAL DATA QUALITY WARNING.
    The other producers -- missing, unresolved, incomplete and absent deferral
    evidence -- are untouched and still suppress corrections."""

    def _warning(self, srf_status, soup_status="DEFERRED", memo=None, cycle=RESUME_CYCLE):
        _, _, _, memo_warning, warning_type, *_ = checks._calculate_expected_status(
            GP, srf_status, soup_status, "90 DAYS", 500.0, 500.0, True,
            _memo() if memo is None else memo, SERVICER, pd.Timestamp(cycle),
            srf_dq="180 DAYS", servicing_fee=0.0)
        return warning_type, memo_warning

    def test_blank_resume_date_still_warns(self):
        # Site 1 (resume_unverified) -- the most important warning in the set.
        wt, msg = self._warning("ACTIVE", memo=_memo(resume=None), cycle="2026-07-01")
        self.assertEqual(wt, "deferral_dates")
        self.assertIn("Regular Payment Resume Date is blank", msg)

    def test_blank_resume_date_still_reads_documented(self):
        # WHY the fix could not be a predicate on evidence_status: a blank
        # resume date leaves orderly dates and a known type, so this loan is
        # ALSO `documented`. Suppressing the warning wherever evidence is
        # documented would have disarmed site 1 silently.
        *_, evidence = checks._calculate_expected_status(
            GP, "ACTIVE", "DEFERRED", "90 DAYS", 500.0, 500.0, True,
            _memo(resume=None), SERVICER, pd.Timestamp("2026-07-01"),
            srf_dq="180 DAYS", servicing_fee=0.0)
        self.assertEqual(evidence["evidence_status"], "documented")

    def test_unclassified_type_still_warns(self):
        # Site 3.
        wt, msg = self._warning("ACTIVE", memo=_memo(action=UNCLASSIFIED), cycle=MID_CYCLE)
        self.assertEqual(wt, "deferral_dates")
        self.assertTrue(msg)

    def test_unclassified_warning_survives_the_resume_cycle_too(self):
        # Site 4 used to OVERWRITE this message with its own. Now that it
        # assigns nothing, the upstream warning survives -- strictly more
        # coverage than before, not less.
        wt, msg = self._warning("DEFERRED", memo=_memo(action=UNCLASSIFIED))
        self.assertEqual(wt, "deferral_dates")
        self.assertNotIn("still reports Loan Status = DEFERRED", msg or "")

    def test_reported_deferred_with_no_memo_deferral_still_warns(self):
        # Sites 5-6: nothing in the memo supports the reported DEFERRED.
        for srf, soup_st in (("DEFERRED", "DEFERRED"), ("ACTIVE", "DEFERRED"),
                             ("DEFERRED", "ACTIVE")):
            with self.subTest(srf=srf, soup=soup_st):
                wt, msg = self._warning(srf, soup_status=soup_st,
                                        memo=_memo(gp="8880000002"))
                self.assertEqual(wt, "deferral_dates")
                self.assertIn("does not provide sufficient deferral", msg)

    def test_warning_reasons_still_suppresses_corrections(self):
        # The gate itself is unchanged.
        self.assertIn("DEFERRAL DATA QUALITY WARNING", corrections.WARNING_REASONS)
        self.assertIn("MEMO DATA QUALITY WARNING", corrections.WARNING_REASONS)
        self.assertIn("SCHEDULED PAYMENT AMBIGUOUS", corrections.WARNING_REASONS)


class InsufficientEvidenceTests(unittest.TestCase):
    """Evidence is sufficient by CONSTRUCTION inside the resume branch -- the
    phase exists only because a dated memo resume assertion was found for this
    cycle. Genuine insufficiency lives in sibling branches, which decline to
    assert a status and are untouched."""

    def test_no_resume_date_never_reaches_the_resume_branch(self):
        # Never infer a boundary the memo does not state.
        _, expected, phase = _calc("LIQUIDATION", memo=_memo(resume=None))
        self.assertNotEqual(phase, "resume")
        self.assertNotEqual(expected, "ACTIVE")

    def test_resume_unverified_still_asserts_nothing(self):
        # Past the Deferral End Date with no known resume date: do not advance,
        # do not reset, do not assert a Loan Status.
        m = _memo(resume=None)
        _, expected, phase = _calc("LIQUIDATION", cycle="2026-07-01", memo=m)
        self.assertEqual(phase, "resume_unverified")
        self.assertIsNone(expected)

    def test_no_memo_row_for_this_loan_asserts_nothing_from_this_branch(self):
        _, expected, phase = _calc("LIQUIDATION", memo=_memo(gp="8880000002"))
        self.assertIsNone(phase)
        self.assertNotEqual(expected, "ACTIVE")


class StartAndMidPhasesUnchangedTests(unittest.TestCase):
    """Deferral start/mid logic is explicitly out of scope."""

    def test_mid_window_still_expects_deferred(self):
        _, expected, _ = _calc("ACTIVE", cycle=MID_CYCLE)
        self.assertEqual(expected, "DEFERRED")

    def test_mid_window_partial_still_expects_deferred(self):
        _, expected, _ = _calc("ACTIVE", action=PARTIAL, cycle=MID_CYCLE)
        self.assertEqual(expected, "DEFERRED")

    def test_full_deferral_still_advances_one_rank_regardless_of_payment(self):
        for paid in (0.0, 500.0, 2000.0):
            with self.subTest(paid=paid):
                calc, expected, _ = _calc("ACTIVE", cycle=MID_CYCLE,
                                          prior_dq="30 DAYS", received=paid)
                self.assertEqual(calc, "60 DAYS")
                self.assertEqual(expected, "DEFERRED")


class DelinquencyCalculationUnchangedTests(unittest.TestCase):
    """This change touches Loan Status only. The DQ result at resume must be
    identical for every reported Loan Status."""

    def test_resume_dq_is_unaffected_by_the_reported_loan_status(self):
        baseline = _calc("ACTIVE", prior_dq="90 DAYS", received=500.0)[0]
        self.assertEqual(baseline, "CURRENT")   # reset, then one satisfied installment
        for reported in ("LIQUIDATION", "FORECLOSURE", "REO", "PAID-IN-FULL",
                         "BANKRUPTCY", "SUBSTITUTION-OUT"):
            with self.subTest(reported=reported):
                self.assertEqual(_calc(reported, prior_dq="90 DAYS", received=500.0)[0],
                                 baseline)

    def test_resume_still_resets_then_evaluates(self):
        self.assertEqual(_calc("LIQUIDATION", prior_dq="120 DAYS", received=0.0)[0],
                         "DELINQUENT")


class EndToEndRecommendationTests(unittest.TestCase):
    """The GP 6717929105 shape, through status_match_check and
    build_recommendations -- the path that was producing no Loan Status
    control on the Manual Review page."""

    def test_liquidation_at_resume_is_now_a_loan_status_mismatch(self):
        record, entry, _ = _end_to_end("LIQUIDATION")
        self.assertEqual(record["Servicer Loan Status"], "LIQUIDATION")
        self.assertEqual(record["Expected Loan Status"], "ACTIVE")
        self.assertEqual(record["Loan Status Match"], "NO")
        self.assertIn("LOAN STATUS MISMATCH", entry["reasons"])
        self.assertEqual(entry["expected_loan_status_effective"], "ACTIVE")

    def test_liquidation_at_resume_produces_a_review_flagged_active_recommendation(self):
        _, _, recs = _end_to_end("LIQUIDATION")
        loan_status = [r for r in recs if r["field"] == "Loan Status"]
        self.assertEqual(len(loan_status), 1)
        rec = loan_status[0]
        self.assertEqual(rec["reported"], "LIQUIDATION")
        self.assertEqual(rec["calculated"], "ACTIVE")
        self.assertEqual(rec["reviewFlag"], corrections.REVIEW_REASON_TERMINAL_STATUS)

    def test_the_recommendation_is_never_auto_applied(self):
        # review_required is what stops an automatic write; the reviewer_approved
        # gate in _apply_one_correction is the other half and is untouched here.
        eligible = [e for e in corrections.build_eligible_corrections(
            checks.status_match_check(_srf("LIQUIDATION"), _soup(),
                                      memo_df=_memo(), servicer_name=SERVICER)["correction_data"])
            if e["field"] == "Loan Status"]
        self.assertEqual(len(eligible), 1)
        self.assertTrue(eligible[0]["review_required"])

    def test_active_is_an_allowed_reviewer_value_for_loan_status(self):
        # Unchanged by this commit -- asserted so the end-to-end claim
        # ("the reviewer can select ACTIVE") rests on something checked.
        self.assertEqual(corrections.validate_approved_value("Loan Status", "ACTIVE",
                                                             substituted=True),
                         (True, "ACTIVE"))

    def test_the_application_still_cannot_propose_a_collection_status(self):
        # The proposed value can only ever move a loan OUT of a terminal
        # collection status, never into one.
        for value in ("LIQUIDATION", "FORECLOSURE", "REO"):
            with self.subTest(value=value):
                self.assertIsNone(corrections._canonical_proposable_loan_status(value))

    def test_other_collection_statuses_produce_the_same_recommendation(self):
        for reported in ("FORECLOSURE", "REO"):
            with self.subTest(reported=reported):
                record, entry, recs = _end_to_end(reported)
                self.assertEqual(record["Loan Status Match"], "NO")
                self.assertIn("LOAN STATUS MISMATCH", entry["reasons"])
                loan_status = [r for r in recs if r["field"] == "Loan Status"]
                self.assertEqual(len(loan_status), 1)
                self.assertEqual(loan_status[0]["calculated"], "ACTIVE")

    def test_active_at_resume_produces_no_loan_status_finding_at_all(self):
        # The GP 5658839107 / 5520789102 regression: a legitimate resume must
        # stay clean, not become a new false positive.
        record, entry, recs = _end_to_end("ACTIVE")
        if record is not None:
            self.assertEqual(record["Loan Status Match"], "YES")
        if entry is not None:
            self.assertNotIn("LOAN STATUS MISMATCH", entry["reasons"])
        self.assertEqual([r for r in recs if r["field"] == "Loan Status"], [])

    def test_mixed_case_active_at_resume_stays_clean(self):
        # GP 5658839107 reports "Active", not "ACTIVE". The match test is
        # case-insensitive, so derivation must not turn it into a mismatch.
        record, entry, recs = _end_to_end("Active")
        if record is not None:
            self.assertEqual(record["Loan Status Match"], "YES")
        if entry is not None:
            self.assertNotIn("LOAN STATUS MISMATCH", entry["reasons"])
        self.assertEqual([r for r in recs if r["field"] == "Loan Status"], [])

    def test_substitution_out_is_flagged_but_never_auto_proposed(self):
        # Option 1's only new behaviour, with zero occurrences in either
        # delivered drop: a non-collection, non-carve-out status at resume
        # becomes a visible mismatch, and no branch proposes a value for it.
        record, entry, recs = _end_to_end("SUBSTITUTION-OUT")
        self.assertEqual(record["Loan Status Match"], "NO")
        self.assertIn("LOAN STATUS MISMATCH", entry["reasons"])
        self.assertEqual([r for r in recs if r["field"] == "Loan Status"], [])


if __name__ == "__main__":
    unittest.main()
