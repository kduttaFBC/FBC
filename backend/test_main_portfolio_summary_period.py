"""
Regression test for main._resolve_portfolio_summary_cycle_month, the pure
helper function that decides how to resolve a single reporting-month for
the Portfolio Summary workbook filename across multiple uploaded servicer
files, and detects a mixed-reporting-period conflict.

CLIENT-CONFIRMED (2026-08-04): combining different reporting periods into
one monthly portfolio summary could be misleading, so the preferred
behavior is to STOP and return a clear validation error rather than
silently picking the first month found or falling back to a neutral
filename.

Run with:
    python -m unittest test_main_portfolio_summary_period.py
"""

import unittest

import pandas as pd

import main
import report


class SingleReportingMonthTests(unittest.TestCase):
    """All uploaded files covering the same reporting month must resolve
    cleanly, with no error, and produce the correctly-named workbook."""

    def test_all_files_june_2026_resolves_without_error(self):
        june = pd.Timestamp("2026-06-01")
        cycle_month, error = main._resolve_portfolio_summary_cycle_month({june})
        self.assertIsNone(error)
        self.assertEqual(cycle_month, june)

    def test_all_files_june_2026_produces_correct_filename(self):
        june = pd.Timestamp("2026-06-01")
        cycle_month, error = main._resolve_portfolio_summary_cycle_month({june})
        self.assertIsNone(error)
        filename = report.portfolio_summary_filename(cycle_month, label="Audited")
        self.assertEqual(filename, "Portfolio_Summary_Audited_June_2026.xlsx")

    def test_single_month_added_multiple_times_is_still_one_month(self):
        # Multiple servicers all reporting the SAME month -- a set
        # naturally de-duplicates, but confirm this explicitly since it's
        # exactly the normal, expected case (most runs have every servicer
        # covering the same period).
        june = pd.Timestamp("2026-06-01")
        cycle_months = set()
        for _ in range(4):  # simulating 4 servicer files, all June
            cycle_months.add(june)
        cycle_month, error = main._resolve_portfolio_summary_cycle_month(cycle_months)
        self.assertIsNone(error)
        self.assertEqual(cycle_month, june)


class MixedReportingMonthTests(unittest.TestCase):
    """More than one distinct reporting month must produce a clear error
    and NOT a resolved cycle month — the whole request should stop rather
    than silently pick one or generate a neutral filename."""

    def test_june_and_july_produce_a_reporting_period_mismatch_error(self):
        june = pd.Timestamp("2026-06-01")
        july = pd.Timestamp("2026-07-01")
        cycle_month, error = main._resolve_portfolio_summary_cycle_month({june, july})
        self.assertIsNone(cycle_month)
        self.assertIsNotNone(error)
        self.assertIn("different reporting months", error)
        self.assertIn("June 2026", error)
        self.assertIn("July 2026", error)

    def test_mismatch_error_lists_months_in_chronological_order(self):
        # Regression guard for a real bug caught during development: an
        # earlier version sorted the FORMATTED strings alphabetically
        # ("July 2026, June 2026"), not the underlying dates
        # chronologically. Must always read oldest-to-newest.
        june = pd.Timestamp("2026-06-01")
        july = pd.Timestamp("2026-07-01")
        _, error = main._resolve_portfolio_summary_cycle_month({july, june})
        self.assertIn("(June 2026, July 2026)", error)

    def test_three_distinct_months_all_listed(self):
        may = pd.Timestamp("2026-05-01")
        june = pd.Timestamp("2026-06-01")
        july = pd.Timestamp("2026-07-01")
        cycle_month, error = main._resolve_portfolio_summary_cycle_month({july, may, june})
        self.assertIsNone(cycle_month)
        self.assertIn("(May 2026, June 2026, July 2026)", error)


class BlankCycleMonthFallbackTests(unittest.TestCase):
    """A blank/undeterminable cycle month must fall back safely, and must
    NOT be treated as a second distinct month causing a false conflict —
    only ACTUAL distinct known months can trigger the mismatch error."""

    def test_all_blank_falls_back_to_none_without_error(self):
        # Simulates every servicer file having an undeterminable cycle
        # month (e.g. empty/malformed SOUP history) -- nothing was ever
        # added to the set in main.py's loop in this case.
        cycle_month, error = main._resolve_portfolio_summary_cycle_month(set())
        self.assertIsNone(error)
        self.assertIsNone(cycle_month)

    def test_all_blank_falls_back_to_todays_date_filename(self):
        from datetime import date
        cycle_month, error = main._resolve_portfolio_summary_cycle_month(set())
        self.assertIsNone(error)
        filename = report.portfolio_summary_filename(cycle_month, label="Audited")
        self.assertEqual(filename, f"Portfolio_Summary_Audited_{date.today().isoformat()}.xlsx")

    def test_one_real_month_resolves_cleanly_with_no_conflicting_month(self):
        # This is the key case: a blank cycle month is represented by
        # simply NOT adding anything to the set (see main.py's loop) --
        # so a set containing only ONE real month, regardless of how many
        # OTHER servicers contributed nothing, must resolve to that one
        # month without any error. This test constructs the set the same
        # way main.py's loop actually would: only non-None cycle months
        # ever get added.
        cycle_months = set()
        found_months = [pd.Timestamp("2026-06-01"), None, None]  # 1 real, 2 blank servicers
        for cm in found_months:
            if cm is not None:
                cycle_months.add(cm)
        cycle_month, error = main._resolve_portfolio_summary_cycle_month(cycle_months)
        self.assertIsNone(error)
        self.assertEqual(cycle_month, pd.Timestamp("2026-06-01"))

    def test_one_real_month_plus_blanks_still_produces_correct_filename(self):
        cycle_months = set()
        for cm in [None, pd.Timestamp("2026-06-01"), None]:
            if cm is not None:
                cycle_months.add(cm)
        cycle_month, error = main._resolve_portfolio_summary_cycle_month(cycle_months)
        self.assertIsNone(error)
        self.assertEqual(report.portfolio_summary_filename(cycle_month, label="Audited"), "Portfolio_Summary_Audited_June_2026.xlsx")


if __name__ == "__main__":
    unittest.main()
