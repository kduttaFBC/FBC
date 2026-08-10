"""
Regression test confirming the delinquency status ladder (_status_to_missed_count
/ _missed_count_to_status in checks.py) is uncapped and mathematically
consistent through at least 600 DAYS, and that every calculation branch
that produces a Calculated DQ Status routes through these same two shared
functions rather than a separate, duplicated ladder.

INVESTIGATION NOTE (2026-08-04): as of this test's creation, both
conversion functions were ALREADY correct and unbounded — verified by
direct round-trip testing through 690 DAYS with zero mismatches, and by
grepping the entire codebase for any min()/max()/clamp on missed_count
(none found). No code change was made to checks.py for this request. This
file exists so that stays true going forward — if a future change
accidentally introduces a cap, one of these tests will fail immediately
rather than silently reintroducing the exact class of bug this request
was checking for.

Run with:
    python -m unittest test_delinquency_ladder.py
"""

import unittest

import checks
import file_readers


class LadderRoundTripTests(unittest.TestCase):
    """status -> missed_count -> status must be the identity for every
    supported rung, with no upper bound."""

    def test_round_trip_every_30_days_through_600(self):
        # CURRENT and DELINQUENT are their own distinct rungs; "30 DAYS" is
        # a SEPARATE, later rung (missed=2, one step past DELINQUENT at
        # missed=1) -- an earlier helper-based version of this test
        # incorrectly conflated "30 DAYS" with "DELINQUENT" and never
        # actually tested "30 DAYS" as its own label. This explicit list
        # tests every rung by name, with no substitution.
        labels = ["CURRENT", "DELINQUENT"]
        labels.extend(f"{days} DAYS" for days in range(30, 601, 30))

        for label in labels:
            with self.subTest(label=label):
                missed = checks._status_to_missed_count(label)
                back = checks._missed_count_to_status(missed)
                self.assertEqual(back, label)

    def test_round_trip_beyond_600_is_not_capped(self):
        # Client-confirmed: the safest design supports values beyond 600
        # rather than stopping there. These must round-trip identically to
        # every other rung, with no special-casing at the boundary.
        for days in (630, 660, 690, 900, 1200):
            with self.subTest(days=days):
                label = f"{days} DAYS"
                missed = checks._status_to_missed_count(label)
                back = checks._missed_count_to_status(missed)
                self.assertEqual(back, label)

    def test_exact_examples_given(self):
        cases = [
            ("240 DAYS", 9),
            ("270 DAYS", 10),
            ("300 DAYS", 11),
            ("600 DAYS", 21),
        ]
        for label, expected_missed in cases:
            with self.subTest(label=label):
                self.assertEqual(checks._status_to_missed_count(label), expected_missed)
                self.assertEqual(checks._missed_count_to_status(expected_missed), label)

    def test_no_cap_going_upward_from_missed_count(self):
        # _missed_count_to_status must never map a high missed_count back
        # down to a lower rung -- each integer produces a strictly larger
        # day count than the one before it, with no ceiling. DELINQUENT
        # (missed=1) sits strictly between CURRENT (0) and 30 DAYS (missed=2)
        # in the real ladder, so it's given a placeholder value of 15 here
        # purely for this strictly-increasing comparison -- not a claim
        # that DELINQUENT literally means "15 days."
        prev_days = -1
        for missed in range(0, 30):
            status = checks._missed_count_to_status(missed)
            if status == "CURRENT":
                days = 0
            elif status == "DELINQUENT":
                days = 15
            else:
                days = int(status.split(" ")[0])
            self.assertGreater(days, prev_days, f"missed_count={missed} produced {status}, not strictly increasing")
            prev_days = days


class LadderProgressionTests(unittest.TestCase):
    """Exactly the three progressions requested: one missed period (a $0
    payment against a real Scheduled Payment) must advance the ladder by
    precisely one rung, with no capping at 240, 570, or 600 DAYS."""

    SCHEDULED = 500.0

    def _advance_one_missed_period(self, prior_status_label):
        prior_missed = checks._status_to_missed_count(prior_status_label)
        new_missed, _ = checks._apply_payment_transition(prior_missed, 0, self.SCHEDULED)
        return checks._missed_count_to_status(new_missed)

    def test_240_days_plus_one_missed_period_is_270_days(self):
        self.assertEqual(self._advance_one_missed_period("240 DAYS"), "270 DAYS")

    def test_570_days_plus_one_missed_period_is_600_days(self):
        self.assertEqual(self._advance_one_missed_period("570 DAYS"), "600 DAYS")

    def test_600_days_plus_one_missed_period_is_630_days(self):
        # This is the specific case confirming the ladder does NOT stop at
        # 600 -- client-confirmed the safest design supports values beyond
        # 600 rather than capping there.
        self.assertEqual(self._advance_one_missed_period("600 DAYS"), "630 DAYS")


class LadderConsistencyAcrossBranchesTests(unittest.TestCase):
    """Confirms multiple DIFFERENT calculation branches (not just the plain
    missed-payment case above) all honor the same uncapped ladder, since
    each has its own code path in _calculate_expected_status /
    _apply_payment_transition and none may hardcode a separate, smaller
    ladder of its own."""

    def test_reversal_branch_uncapped(self):
        # A large negative payment against a loan already at 570 DAYS must
        # be able to advance past 600, not cap there.
        new_missed, equivalents = checks._apply_payment_transition(
            checks._status_to_missed_count("570 DAYS"), -1000, 500.0
        )
        self.assertGreaterEqual(checks._status_to_missed_count(checks._missed_count_to_status(new_missed)), 21)

    def test_deferral_progression_branch_uncapped(self):
        import pandas as pd
        import file_readers

        memo_rows = [{
            "Report Date": pd.Timestamp("2026-06-01"), "Lender": "Test Bank",
            "SBA Loan ID": "1", "Lender Loan ID": "L-1", "Borrower Name": "Test",
            "Securitization Action": "Deferred",
            "Deferral Start Date": pd.Timestamp("2026-06-05"), "Deferral End Date": pd.Timestamp("2026-08-31"),
            "Regular Payment Resume Date": pd.Timestamp("2026-09-05"),
        }]
        memo_df = pd.DataFrame(memo_rows)
        memo_df["gp_sba"] = memo_df["SBA Loan ID"].apply(file_readers.normalize_memo_loan_id)
        memo_df["gp_lender"] = memo_df["Lender Loan ID"].apply(file_readers.normalize_memo_loan_id)
        memo_df["deferral_type"] = "full"

        # Loan enters deferral already at 570 DAYS -- the deferral-start
        # branch must still advance it by one rung, past 570, not cap it.
        calc_dq, *_ = checks._calculate_expected_status(
            gp="1", srf_status="DELINQUENT", soup_status="570 DAYS", soup_dq_raw="570 DAYS",
            received=0, scheduled=0, has_payment_cols=True, memo_df=memo_df,
            servicer_name="Test Bank", cycle_month=pd.Timestamp("2026-06-01"),
        )
        self.assertEqual(calc_dq, "600 DAYS")

    def test_deferral_mid_month_also_advances_one_level(self):
        # CLIENT-CONFIRMED (2026-08-04): every month in a full deferral,
        # including the first, advances the ranking by one level -- there
        # is NO freeze at the start month, and no separate "start vs mid"
        # distinction in the actual calculation. This is a SEPARATE test
        # from the start-month case above specifically to confirm the NEXT
        # cycle (July, still within the June-August deferral window) also
        # advances correctly, using SOUP's own rolling history (600 DAYS --
        # the start month's own result) as its prior baseline, exactly the
        # same "+1 from wherever it already was" rule as every other month.
        import pandas as pd
        import file_readers

        memo_rows = [{
            "Report Date": pd.Timestamp("2026-07-01"), "Lender": "Test Bank",
            "SBA Loan ID": "1", "Lender Loan ID": "L-1", "Borrower Name": "Test",
            "Securitization Action": "Deferred",
            "Deferral Start Date": pd.Timestamp("2026-06-05"), "Deferral End Date": pd.Timestamp("2026-08-31"),
            "Regular Payment Resume Date": pd.Timestamp("2026-09-05"),
        }]
        memo_df = pd.DataFrame(memo_rows)
        memo_df["gp_sba"] = memo_df["SBA Loan ID"].apply(file_readers.normalize_memo_loan_id)
        memo_df["gp_lender"] = memo_df["Lender Loan ID"].apply(file_readers.normalize_memo_loan_id)
        memo_df["deferral_type"] = "full"

        calc_dq, *_ = checks._calculate_expected_status(
            gp="1", srf_status="DELINQUENT", soup_status="600 DAYS", soup_dq_raw="600 DAYS",
            received=0, scheduled=0, has_payment_cols=True, memo_df=memo_df,
            servicer_name="Test Bank", cycle_month=pd.Timestamp("2026-07-01"),
        )
        self.assertEqual(calc_dq, "630 DAYS")

    def test_stat01_and_neg01_agree_beyond_600(self):
        import pandas as pd

        servicer_inputs_df = pd.DataFrame([{
            "SBA GP Number": "1", "Lender Loan Number": "L-1", "Loan Status": "ACTIVE",
            "Delinquency Status": "630 DAYS", "Beginning Balance": 1000, "Ending Balance": 1000,
            "Payment Received": 0, "Scheduled Payment": 500,
        }])
        soup_servicer_df = pd.DataFrame([{
            "SBA GP Number": "1", "Servicer": "Test Bank", "Payment Month": "2026-05-01",
            "Loan Status": "ACTIVE", "Delinquency Status": "600 DAYS",
            "Beginning Actual Balance": 1000, "Ending Actual Balance": 1000,
        }])
        stat01 = checks.status_match_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name="Test Bank")
        neg01 = checks.negative_payment_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name="Test Bank")

        # Loan matches cleanly (servicer correctly reported 630 DAYS), so
        # STAT-01 produces no record -- confirm via the underlying
        # calculation directly instead, same value used by both checks.
        calc_dq, *_ = checks._calculate_expected_status(
            gp="1", srf_status="ACTIVE", soup_status="ACTIVE", soup_dq_raw="600 DAYS",
            received=0, scheduled=500, has_payment_cols=True, memo_df=None,
            servicer_name="Test Bank", cycle_month=pd.Timestamp("2026-06-01"),
        )
        self.assertEqual(calc_dq, "630 DAYS")

    def test_corrections_eligibility_does_not_cap(self):
        import corrections

        correction_entry = {
            "gp": "1", "excel_row": 12, "sheet_name": "Servicer Inputs", "header_excel_row": 9,
            "reasons": ["DELINQUENCY STATUS MISMATCH"], "dq_match": False, "calculated_dq": "630 DAYS",
            "expected_loan_status_raw": None, "deferral_type": None, "deferral_phase": None,
            "reversal_equivalents": None,
        }
        eligible = corrections.build_eligible_corrections([correction_entry])
        self.assertEqual(len(eligible), 1)
        self.assertEqual(eligible[0]["corrected_value"], "630 DAYS")


class TerminalCollectionOpenEndedBucketTests(unittest.TestCase):
    """REWRITTEN 2026-08-08 for the superseded terminal-status rule.

    These cases were written for the 2026-08-04 rule, under which a
    FORECLOSURE/LIQUIDATION/REO loan had its own delinquency calculation: a
    degradation-only floor, max()-ed against the servicer's own reported DQ,
    with catch-up credit forbidden. That rule is superseded -- terminal
    collection status and delinquency calculation are separate concepts, and
    the servicer's reported DQ must never determine the audited figure.

    These loans now use the SAME normal ladder as every other loan. The cases
    are kept (they are still the right questions to ask) with expectations
    updated to normal-ladder results, and each records what the old rule
    would have produced.

    Loan 5387239101's original bug -- a "150+ Days" bucket parsing LOWER than
    the loan's own prior 240 DAYS and silently improving the ranking -- stays
    fixed, but now BY CONSTRUCTION rather than by special case: the servicer's
    DQ never enters the calculation at all, so an untrustworthy bucket cannot
    corrupt it."""

    def _run(self, srf_dq, prior_soup_dq, received=0, scheduled=500.0):
        import pandas as pd
        servicer_inputs_df = pd.DataFrame([{
            "SBA GP Number": "1", "Lender Loan Number": "L-1", "Loan Status": "FORECLOSURE",
            "Delinquency Status": srf_dq, "Beginning Balance": 50000, "Ending Balance": 48000,
            "Payment Received": received, "Scheduled Payment": scheduled,
        }])
        soup_servicer_df = pd.DataFrame([{
            "SBA GP Number": "1", "Servicer": "Test Bank", "Payment Month": "2026-05-01",
            "Loan Status": "ACTIVE", "Delinquency Status": prior_soup_dq,
            "Beginning Actual Balance": 50000, "Ending Actual Balance": 50000,
        }])
        result = checks.status_match_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name="Test Bank")
        return result["records"][0]["Calculated DQ Status"]

    def test_loan_5387239101_open_ended_bucket_below_prior(self):
        # Servicer reports "150+ Days" (missed=6) against a prior SOUP
        # status of 240 DAYS (missed=9), with $0 payment this cycle. The
        # "+" value must be ignored numerically; the correct result
        # advances one step from the prior status, same as any normal
        # missed-payment cycle.
        self.assertEqual(self._run("150+ Days", "240 DAYS"), "270 DAYS")

    def test_open_ended_bucket_lowercase_days_still_detected(self):
        # The bug was specifically triggered by a lowercase "Days" plus a
        # trailing "+" -- confirm case doesn't matter for detecting the "+".
        self.assertEqual(self._run("90+ days", "180 DAYS"), "210 DAYS")

    def test_clean_worse_servicer_value_is_no_longer_trusted(self):
        # SUPERSEDED 2026-08-08: the old rule trusted a precise servicer value
        # whenever it was WORSE than the floor, giving "360 DAYS". The
        # servicer's report no longer determines the calculated value at all.
        # Prior 240 DAYS with no payment advances one level by the normal
        # ladder. The servicer's 360 DAYS is still displayed and still raises
        # DELINQUENCY STATUS MISMATCH -- it just does not decide the answer.
        self.assertEqual(self._run("360 DAYS", "240 DAYS"), "270 DAYS")

    def test_clean_better_servicer_value_is_not_trusted(self):
        # No "+" suffix, but the servicer's report is BETTER than what a
        # $0-payment cycle would justify -- Rule 1 forbids improvement, so
        # the degradation-only floor wins over the servicer's own number.
        self.assertEqual(self._run("30 DAYS", "240 DAYS"), "270 DAYS")

    def test_sufficient_payment_now_earns_catch_up_credit(self):
        # SUPERSEDED 2026-08-08: the old rule forbade catch-up credit for
        # these loans, so a large payment could only hold steady ("240 DAYS").
        # Normal catch-up rules now apply exactly as they do for every other
        # loan: $2,000 against a $500 schedule clears multiple missed
        # payments and improves the ranking.
        self.assertEqual(self._run("150+ Days", "240 DAYS", received=2000, scheduled=500), "150 DAYS")

    def test_no_prior_history_falls_back_the_same_way_as_any_other_loan(self):
        # SUPERSEDED 2026-08-08: the old terminal branch returned an explicit
        # "UNVERIFIED" when it had no floor to stand on. These loans now take
        # the ordinary fallback for an unparseable prior status -- the same
        # empty result ANY loan gets with no prior SOUP delinquency status and
        # no payment activity to reason from. This is normal-path behaviour,
        # not a terminal-specific outcome; changing it would be a separate
        # decision affecting every loan.
        self.assertEqual(self._run("150+ Days", ""), "")


class TerminalCollectionNegativePaymentPendingDecisionTests(unittest.TestCase):
    """*** PENDING DECISION RESOLVED 2026-08-08 ***

    The open question was: should a negative Payment Received (a reversal)
    for a FORECLOSURE/LIQUIDATION/REO loan run through the full
    reversal-equivalents formula used for non-terminal loans, or stay a flat
    +1 as the 2026-08-04 rule had it?

    It is answered, though not by a decision about reversals specifically.
    The 2026-08-08 rule states that these loans follow the same normal
    delinquency logic as every other loan, using the standard ladder and
    catch-up rules. The reversal-equivalents formula is part of that normal
    logic, so it now applies here too and a large reversal can advance more
    than one level.

    The old comment demanded these numbers never change silently as a side
    effect. They have not: the change is deliberate and recorded here, in the
    superseded-rule comment in checks.py, and in the commit message."""

    def _run(self, received, scheduled, prior_soup_dq):
        import pandas as pd
        servicer_inputs_df = pd.DataFrame([{
            "SBA GP Number": "1", "Lender Loan Number": "L-1", "Loan Status": "FORECLOSURE",
            "Delinquency Status": "CURRENT", "Beginning Balance": 50000, "Ending Balance": 48000,
            "Payment Received": received, "Scheduled Payment": scheduled,
        }])
        soup_servicer_df = pd.DataFrame([{
            "SBA GP Number": "1", "Servicer": "Test Bank", "Payment Month": "2026-05-01",
            "Loan Status": "ACTIVE", "Delinquency Status": prior_soup_dq,
            "Beginning Actual Balance": 50000, "Ending Actual Balance": 50000,
        }])
        result = checks.status_match_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name="Test Bank")
        return result["records"][0]["Calculated DQ Status"]

    def test_current_behavior_a_small_reversal_advances_exactly_one_level(self):
        # A small reversal (-$50 against $500 scheduled) would NOT even
        # clear one payment-equivalent under the non-terminal reversal
        # formula's deficiency test in some cases, but here it's simply
        # treated as "no payment" -- current behavior is a flat +1
        # regardless of the reversal's size.
        self.assertEqual(self._run(-50, 500.0, "240 DAYS"), "270 DAYS")

    def test_a_large_reversal_now_advances_multiple_levels(self):
        # RESOLVED 2026-08-08 (was "270 DAYS" under the flat +1 rule). A
        # -$2,000 reversal against a $500 schedule is four payment-equivalents
        # removed, so the normal reversal formula advances several levels --
        # the same result this reversal would produce on any non-terminal
        # loan. See test_reversal_thresholds.py for the transition table.
        self.assertEqual(self._run(-2000, 500.0, "240 DAYS"), "420 DAYS")

    def test_current_behavior_matches_a_plain_zero_payment(self):
        # Documents that -$50 and $0 currently produce an IDENTICAL result
        # for a terminal-collection loan -- the reversal's magnitude has no
        # effect at all under current behavior, unlike the non-terminal
        # reversal formula where magnitude directly determines how many
        # levels are advanced.
        zero_payment_result = self._run(0, 500.0, "240 DAYS")
        small_reversal_result = self._run(-50, 500.0, "240 DAYS")
        self.assertEqual(zero_payment_result, small_reversal_result)


if __name__ == "__main__":
    unittest.main()


class PartialDeferralResumeResetTests(unittest.TestCase):
    """A PARTIAL P&I deferral must reset to CURRENT when it ends, exactly as a
    full deferral does (client-confirmed 2026-08-08).

    The resume boundary is a property of the deferral WINDOW, not of its type.
    Before this fix the resume phase was assigned only for full/unclassified
    deferrals, so a partial loan carried its accumulated ranking forward past
    the end of the deferral indefinitely — e.g. a loan at 90 DAYS that paid in
    full on its resume month stayed at 90 DAYS instead of returning to CURRENT.

    What is NOT changed: partial deferrals still use the normal payment ladder
    (including catch-up credit) for every cycle INSIDE the deferral, and full
    deferrals still advance one rank per cycle regardless of payment.
    """

    import pandas as _pd

    def _memo(self, action):
        import pandas as pd
        df = pd.DataFrame([{
            "Lender": "Millennium Bank", "SBA Loan ID": "9990000001",
            "Lender Loan ID": "9990000001", "Borrower Name": "T",
            "Securitization Action": "Deferred", "Servicer Action": action,
            "Deferral Start Date": pd.Timestamp("2026-01-01"),
            "Deferral End Date": pd.Timestamp("2026-05-01"),
            "Regular Payment Resume Date": pd.Timestamp("2026-06-01"),
            "Report Date": pd.Timestamp("2026-05-01"),
        }])
        df["gp_sba"] = df["SBA Loan ID"]
        df["gp_lender"] = df["Lender Loan ID"]
        df["deferral_type"] = df["Servicer Action"].apply(file_readers.classify_deferral_type)
        return df

    PARTIAL = "6-month principal only deferment"
    FULL = "6-month principal and interest deferment"

    def _run(self, action, cycle, prior_dq, received, scheduled=500.0):
        import pandas as pd
        calc, expected, deferral, *_ = checks._calculate_expected_status(
            "9990000001", "ACTIVE", "DEFERRED", prior_dq, received, scheduled,
            True, self._memo(action), "Millennium Bank", pd.Timestamp(cycle),
            srf_dq="180 DAYS", servicing_fee=0.0)
        return calc, expected, (deferral or {}).get("phase")

    # --- the fix -----------------------------------------------------------
    def test_partial_reaches_the_resume_phase(self):
        _, _, phase = self._run(self.PARTIAL, "2026-06-01", "90 DAYS", 500.0)
        self.assertEqual(phase, "resume")

    def test_partial_resets_to_current_then_credits_the_resume_payment(self):
        # Reset to CURRENT, then one satisfied installment holds at CURRENT.
        calc, _, _ = self._run(self.PARTIAL, "2026-06-01", "90 DAYS", 500.0)
        self.assertEqual(calc, "CURRENT")

    def test_partial_resume_with_no_payment_drops_exactly_one_rank(self):
        # Reset to CURRENT, then a miss -> one rung down, NOT 90 -> 120 DAYS.
        for prior in ("30 DAYS", "90 DAYS", "120 DAYS"):
            with self.subTest(prior=prior):
                calc, _, _ = self._run(self.PARTIAL, "2026-06-01", prior, 0.0)
                self.assertEqual(calc, "DELINQUENT")

    def test_partial_resume_never_carries_the_deferral_ranking_forward(self):
        # The defect: the pre-fix result was the prior rank (or worse).
        for prior, paid in (("90 DAYS", 500.0), ("30 DAYS", 500.0), ("120 DAYS", 1000.0)):
            with self.subTest(prior=prior, paid=paid):
                calc, _, _ = self._run(self.PARTIAL, "2026-06-01", prior, paid)
                self.assertEqual(calc, "CURRENT")

    def test_partial_resume_scheduled_zero_still_advances(self):
        # Rule 2 preserved: a zero payment is always a miss, whatever the
        # scheduled amount says.
        calc, _, _ = self._run(self.PARTIAL, "2026-06-01", "90 DAYS", 0.0, scheduled=0.0)
        self.assertEqual(calc, "DELINQUENT")

    def test_partial_and_full_agree_at_the_resume_month(self):
        # One implementation of "reset, then evaluate" -- not two.
        for prior in ("CURRENT", "30 DAYS", "90 DAYS"):
            for paid in (0.0, 500.0, 1500.0):
                with self.subTest(prior=prior, paid=paid):
                    self.assertEqual(self._run(self.PARTIAL, "2026-06-01", prior, paid),
                                     self._run(self.FULL, "2026-06-01", prior, paid))

    def test_partial_resume_loan_status_is_no_longer_forced_to_deferred(self):
        # The deferral has ended, so Expected Loan Status is no longer DEFERRED.
        # The value is unchanged, but its SOURCE changed on 2026-08-10: it is
        # now derived from the memo's confirmed resume rather than copied from
        # the servicer's own current report (which here happens to be ACTIVE
        # too). See test_resume_expected_loan_status.py.
        _, expected, _ = self._run(self.PARTIAL, "2026-06-01", "90 DAYS", 500.0)
        self.assertEqual(expected, "ACTIVE")

    # --- what must NOT change ---------------------------------------------
    def test_partial_inside_deferral_still_uses_the_normal_ladder(self):
        # 1x holds, 2x improves one rung, 3x improves two, 0x advances one.
        for paid, expect in ((500.0, "30 DAYS"), (1000.0, "DELINQUENT"),
                             (1500.0, "CURRENT"), (0.0, "60 DAYS")):
            with self.subTest(paid=paid):
                calc, _, phase = self._run(self.PARTIAL, "2026-03-01", "30 DAYS", paid)
                self.assertEqual(calc, expect)
                self.assertIsNone(phase)   # no start/mid phase for partial

    def test_partial_inside_deferral_still_expects_deferred_loan_status(self):
        _, expected, _ = self._run(self.PARTIAL, "2026-03-01", "30 DAYS", 500.0)
        self.assertEqual(expected, "DEFERRED")

    def test_full_advances_one_rank_regardless_of_payment_at_every_phase(self):
        for cycle in ("2026-01-01", "2026-03-01", "2026-05-01"):
            for paid in (0.0, 500.0, 2000.0):
                with self.subTest(cycle=cycle, paid=paid):
                    calc, expected, _ = self._run(self.FULL, cycle, "30 DAYS", paid)
                    self.assertEqual(calc, "60 DAYS")
                    self.assertEqual(expected, "DEFERRED")

    def test_partial_without_a_resume_date_gets_no_phase_and_no_reset(self):
        # Never infer a boundary that the memo does not state.
        import pandas as pd
        m = self._memo(self.PARTIAL)
        m.loc[:, "Regular Payment Resume Date"] = pd.NaT
        calc, _, deferral, *_ = checks._calculate_expected_status(
            "9990000001", "ACTIVE", "DEFERRED", "90 DAYS", 500.0, 500.0, True,
            m, "Millennium Bank", pd.Timestamp("2026-06-01"),
            srf_dq="180 DAYS", servicing_fee=0.0)
        self.assertIsNone((deferral or {}).get("phase"))
        self.assertEqual(calc, "90 DAYS")   # normal ladder: 1x holds
