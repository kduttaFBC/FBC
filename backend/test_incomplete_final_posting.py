"""
Regression tests for the incomplete-final-posting fallback in
consolidate_cycle_postings(), and the new DQ-01 (Final Posting Incomplete)
check that surfaces it.

Discovered on real Peoples Bank data (GP 9031018509, June 2026): a loan's
chronologically last posting -- explicitly flagged "Last Payment Instance?
= Y" -- was completely blank (no Loan Status, Delinquency Status, or
Ending Balance), while four earlier, complete postings existed for the
same cycle. The fallback must be NARROW: it should only activate when the
final posting has no authoritative signal at all, never for a posting
that legitimately reports a real $0 balance or a real terminal status.

Run with: python -m unittest test_incomplete_final_posting.py
"""

import unittest

import pandas as pd

import checks


def _row(gp, instance, loan_status, dq_status, beginning, ending, received, last_flag="N"):
    return {
        "SBA GP Number": gp, "Loan Payment Instance": instance,
        "Loan Status": loan_status, "Delinquency Status": dq_status,
        "Beginning Balance": beginning, "Ending Balance": ending,
        "Payment Received": received, "Scheduled Payment": 500,
        "Last Payment Instance?": last_flag,
    }


class IncompleteFinalPostingFallbackTests(unittest.TestCase):
    """Core fallback behavior, using the exact real-world shape that
    surfaced this: a blank final posting with empty-string (not NaN)
    Loan Status/Delinquency Status cells -- this distinction matters, see
    test_blank_detection_handles_empty_string_not_just_nan below."""

    def test_blank_final_posting_falls_back_to_last_valid(self):
        rows = [
            _row("1", 1, "ACTIVE", "CURRENT", 656666.21, 655127.99, 2292.28),
            _row("1", 2, "ACTIVE", "CURRENT", 655127.99, 652931.05, 3074.61),
            _row("1", 3, "ACTIVE", "CURRENT", 652931.05, 651530.73, 2275.05),
            _row("1", 4, "ACTIVE", "CURRENT", 651530.73, 650128.27, 2275.30),
            _row("1", 5, "", "", None, None, None, last_flag="Y"),
        ]
        si_grouped = checks.consolidate_cycle_postings(pd.DataFrame(rows))
        row = si_grouped.loc["1"]
        self.assertEqual(row["loan_status"], "ACTIVE")
        self.assertEqual(row["dq_status"], "CURRENT")
        self.assertEqual(row["ending_balance"], 650128.27)
        self.assertTrue(row["incomplete_final_posting"])

    def test_beginning_balance_unaffected_still_from_first_posting(self):
        # Beginning Balance is sourced from the FIRST posting and was never
        # touched by this bug -- confirms the fallback doesn't change it.
        rows = [
            _row("1", 1, "ACTIVE", "CURRENT", 656666.21, 655127.99, 2292.28),
            _row("1", 2, "", "", None, None, None, last_flag="Y"),
        ]
        si_grouped = checks.consolidate_cycle_postings(pd.DataFrame(rows))
        self.assertEqual(si_grouped.loc["1"]["beginning_balance"], 656666.21)

    def test_disclosure_fields_point_at_the_correct_rows(self):
        rows = [
            _row("1", 1, "ACTIVE", "CURRENT", 100000, 99000, 500),
            _row("1", 2, "ACTIVE", "CURRENT", 99000, 98000, 500),
            _row("1", 3, "", "", None, None, None, last_flag="Y"),
        ]
        df = pd.DataFrame(rows)
        df["_excel_row"] = [10, 11, 12]
        si_grouped = checks.consolidate_cycle_postings(df)
        row = si_grouped.loc["1"]
        self.assertEqual(row["original_last_excel_row"], 12)
        self.assertEqual(row["effective_last_excel_row"], 11)


class IncompleteFinalPostingNarrownessTests(unittest.TestCase):
    """The fallback must NEVER trigger for a legitimate value -- these are
    the negative-control tests confirming it stays narrow."""

    def test_legitimate_zero_ending_balance_is_not_treated_as_blank(self):
        # A loan legitimately paid down to exactly $0 -- Loan Status and
        # Delinquency Status ARE present (e.g. PAID-IN-FULL/CURRENT), so
        # this must NOT trigger the fallback even though Ending Balance is
        # numerically zero.
        rows = [
            _row("1", 1, "ACTIVE", "CURRENT", 5000, 2500, 2500),
            _row("1", 2, "PAID-IN-FULL", "CURRENT", 2500, 0, 2500, last_flag="Y"),
        ]
        si_grouped = checks.consolidate_cycle_postings(pd.DataFrame(rows))
        row = si_grouped.loc["1"]
        self.assertFalse(row["incomplete_final_posting"])
        self.assertEqual(row["ending_balance"], 0)
        self.assertEqual(row["loan_status"], "PAID-IN-FULL")

    def test_legitimate_terminal_status_is_not_treated_as_blank(self):
        # A real, valid terminal status with a real Ending Balance -- must
        # not trigger the fallback just because it's an unusual value.
        rows = [
            _row("1", 1, "ACTIVE", "CURRENT", 50000, 48000, 2000),
            _row("1", 2, "FORECLOSURE", "240 DAYS", 48000, 47500, 0, last_flag="Y"),
        ]
        si_grouped = checks.consolidate_cycle_postings(pd.DataFrame(rows))
        row = si_grouped.loc["1"]
        self.assertFalse(row["incomplete_final_posting"])
        self.assertEqual(row["loan_status"], "FORECLOSURE")
        self.assertEqual(row["ending_balance"], 47500)

    def test_partial_blank_final_posting_does_not_trigger(self):
        # Only Ending Balance is blank, but Loan Status/DQ Status ARE
        # present on the final posting -- the rule requires ALL THREE
        # blank, so this must NOT trigger (a genuinely different, not-yet
        # -handled data-quality situation, not this one).
        rows = [
            _row("1", 1, "ACTIVE", "CURRENT", 50000, 48000, 2000),
            _row("1", 2, "ACTIVE", "CURRENT", 48000, None, 2000, last_flag="Y"),
        ]
        si_grouped = checks.consolidate_cycle_postings(pd.DataFrame(rows))
        row = si_grouped.loc["1"]
        self.assertFalse(row["incomplete_final_posting"])

    def test_single_posting_loan_with_real_data_is_unaffected(self):
        rows = [_row("1", 1, "ACTIVE", "CURRENT", 10000, 9500, 500, last_flag="Y")]
        si_grouped = checks.consolidate_cycle_postings(pd.DataFrame(rows))
        row = si_grouped.loc["1"]
        self.assertFalse(row["incomplete_final_posting"])
        self.assertEqual(row["ending_balance"], 9500)

    def test_blank_detection_handles_empty_string_not_just_nan(self):
        # Regression guard for a real bug caught during development: the
        # originally-proposed valid_rows filter used raw .notna(), which
        # does NOT catch an empty string -- pandas treats '' as "not NaN."
        # The real data that surfaced this whole issue had empty-string
        # blanks, not NaN, so this must be handled correctly or the
        # fallback silently does nothing.
        rows = [
            _row("1", 1, "ACTIVE", "CURRENT", 100000, 99000, 500),
            _row("1", 2, "", "", None, None, None, last_flag="Y"),  # empty string, not NaN/None
        ]
        si_grouped = checks.consolidate_cycle_postings(pd.DataFrame(rows))
        row = si_grouped.loc["1"]
        self.assertTrue(row["incomplete_final_posting"])
        self.assertEqual(row["loan_status"], "ACTIVE")
        self.assertEqual(row["ending_balance"], 99000)

    def test_all_postings_blank_falls_back_to_last_anyway(self):
        # Edge case: if EVERY posting for a loan is blank (no valid row
        # exists at all), effective_last must fall back to `last` rather
        # than crash -- there's nothing better to use, but it must not error.
        rows = [
            _row("1", 1, "", "", None, None, None),
            _row("1", 2, "", "", None, None, None, last_flag="Y"),
        ]
        si_grouped = checks.consolidate_cycle_postings(pd.DataFrame(rows))
        row = si_grouped.loc["1"]
        self.assertTrue(row["incomplete_final_posting"])
        self.assertEqual(row["loan_status"], "")


class FinalPostingIncompleteCheckTests(unittest.TestCase):
    """DQ-01: the check that surfaces the fallback for manual review."""

    def test_check_fires_for_an_affected_loan(self):
        rows = [
            _row("1", 1, "ACTIVE", "CURRENT", 656666.21, 655127.99, 2292.28),
            _row("1", 2, "ACTIVE", "CURRENT", 655127.99, 652931.05, 3074.61),
            _row("1", 3, "ACTIVE", "CURRENT", 652931.05, 651530.73, 2275.05),
            _row("1", 4, "ACTIVE", "CURRENT", 651530.73, 650128.27, 2275.30),
            _row("1", 5, "", "", None, None, None, last_flag="Y"),
        ]
        result = checks.final_posting_incomplete_check(pd.DataFrame(rows), servicer_name="Test Bank")
        self.assertEqual(result["status"], "WARN")
        self.assertEqual(result["issues_found"], 1)
        r = result["records"][0]
        self.assertEqual(r["SBA GP Number"], "1")
        self.assertEqual(r["Ending Balance Used"], 650128.27)
        self.assertIn("blank Loan Status, Delinquency Status, and Ending Balance", r["Detail"])

    def test_check_does_not_fire_for_normal_loans(self):
        rows = [
            _row("1", 1, "ACTIVE", "CURRENT", 10000, 9500, 500, last_flag="Y"),
            _row("2", 1, "DEFERRED", "30 DAYS", 20000, 20000, 0, last_flag="Y"),
        ]
        result = checks.final_posting_incomplete_check(pd.DataFrame(rows), servicer_name="Test Bank")
        self.assertEqual(result["status"], "PASS")
        self.assertEqual(result["issues_found"], 0)
        self.assertEqual(result["records"], [])

    def test_check_does_not_fire_for_legitimate_zero_balance_loan(self):
        rows = [
            _row("1", 1, "ACTIVE", "CURRENT", 5000, 2500, 2500),
            _row("1", 2, "PAID-IN-FULL", "CURRENT", 2500, 0, 2500, last_flag="Y"),
        ]
        result = checks.final_posting_incomplete_check(pd.DataFrame(rows), servicer_name="Test Bank")
        self.assertEqual(result["issues_found"], 0)


class RealPeoplesBankJuneRegressionTests(unittest.TestCase):
    """Locks in the exact real-world numbers this fix was built to
    produce, using the same synthetic reconstruction pattern established
    elsewhere in this test suite (the real file isn't bundled here, but
    these are the loan's actual real values, traced directly from it)."""

    def test_gp_9031018509_reconstruction(self):
        rows = [
            _row("9031018509", 1, "ACTIVE", "CURRENT", 656666.21, 655127.99, 2292.28),
            _row("9031018509", 2, "ACTIVE", "CURRENT", 655127.99, 652931.05, 3074.61),
            _row("9031018509", 3, "ACTIVE", "CURRENT", 652931.05, 651530.73, 2275.05),
            _row("9031018509", 4, "ACTIVE", "CURRENT", 651530.73, 650128.27, 2275.30),
            _row("9031018509", 5, "", "", None, None, None, last_flag="Y"),
        ]
        si_grouped = checks.consolidate_cycle_postings(pd.DataFrame(rows))
        row = si_grouped.loc["9031018509"]
        self.assertEqual(row["loan_status"], "ACTIVE")
        self.assertEqual(row["dq_status"], "CURRENT")
        self.assertEqual(row["ending_balance"], 650128.27)

        # Confirms this loan's real, correct Ending Balance is exactly what
        # closes the gap between the previously-computed Outstanding
        # ($9,035,136.64) and the client's validated total ($9,685,264.91).
        previously_computed_outstanding = 9035136.64
        client_validated_total = 9685264.91
        self.assertAlmostEqual(
            previously_computed_outstanding + row["ending_balance"],
            client_validated_total,
            places=2,
        )


if __name__ == "__main__":
    unittest.main()
