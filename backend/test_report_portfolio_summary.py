"""
Regression test for the Portfolio Summary workbook (report.generate_portfolio_summary_bytes),
a workbook entirely SEPARATE from the detailed Audit Intelligence Reconciliation
workbook (report.generate_report_bytes).

CLIENT-CONFIRMED (2026-08-04): built from the SAME PORT-01 records already
produced by checks.portfolio_summary_check for each servicer — this test
suite constructs those records directly (rather than re-running the full
audit pipeline) to keep the tests fast and focused purely on the workbook
generation logic itself, which is what changed.

Run with:
    python -m unittest test_report_portfolio_summary.py
"""

import io
import unittest

from openpyxl import load_workbook

import report


def _make_record(servicer, outstanding=100000.0, **overrides):
    """Builds a minimal, valid PORT-01-shaped record for testing, with
    every category defaulting to 0 unless overridden."""
    rec = {"Servicer": servicer, "Outstanding ($)": outstanding, "Outstanding (%)": 100.0 if outstanding else 0.0}
    for cat in report.PORTFOLIO_SUMMARY_CATEGORIES:
        if cat == "Outstanding":
            continue
        rec[f"{cat} ($)"] = 0.0
        rec[f"{cat} (%)"] = 0.0
    rec.update(overrides)
    return rec


class PortfolioSummaryWorkbookStructureTests(unittest.TestCase):
    """Basic structural requirements: exactly two sheets, correct names,
    correct fixed column order."""

    def test_workbook_has_exactly_two_sheets_with_correct_names(self):
        records = [_make_record("Test Bank")]
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes(records, buf)
        wb = load_workbook(buf)
        self.assertEqual(wb.sheetnames, ["Portfolio Summary Dollars", "Portfolio Summary Percentages"])

    def test_dollar_sheet_has_fixed_column_order(self):
        records = [_make_record("Test Bank")]
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes(records, buf)
        wb = load_workbook(buf)
        ws = wb["Portfolio Summary Dollars"]
        headers = [c.value for c in ws[1]]
        self.assertEqual(headers, ["Servicer"] + report.PORTFOLIO_SUMMARY_CATEGORIES)

    def test_percentage_sheet_has_same_fixed_column_order(self):
        records = [_make_record("Test Bank")]
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes(records, buf)
        wb = load_workbook(buf)
        ws = wb["Portfolio Summary Percentages"]
        headers = [c.value for c in ws[1]]
        self.assertEqual(headers, ["Servicer"] + report.PORTFOLIO_SUMMARY_CATEGORIES)


class PortfolioSummaryOneServicerTests(unittest.TestCase):
    """One uploaded servicer produces exactly one summary row."""

    def test_one_servicer_produces_one_row(self):
        records = [_make_record("Poppy Bank", outstanding=6660853.44)]
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes(records, buf)
        wb = load_workbook(buf)
        ws = wb["Portfolio Summary Dollars"]
        self.assertEqual(ws.max_row, 2)  # header + 1 data row
        self.assertEqual(ws.cell(row=2, column=1).value, "Poppy Bank")
        self.assertEqual(ws.cell(row=2, column=2).value, 6660853.44)


class PortfolioSummaryMultipleServicersTests(unittest.TestCase):
    """Multiple uploaded servicers produce one row each, in the same
    workbook, and servicer names/note series remain distinct (e.g.
    "Lendistry 2025" and "Lendistry 2024" as two separate rows)."""

    def test_multiple_servicers_produce_one_row_each(self):
        records = [
            _make_record("Community Bank and Trust", outstanding=31707250.83),
            _make_record("Lendistry 2025", outstanding=5000000.0),
            _make_record("Lendistry 2024", outstanding=4500000.0),
            _make_record("Peoples Bank", outstanding=9010133.52),
        ]
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes(records, buf)
        wb = load_workbook(buf)
        ws = wb["Portfolio Summary Dollars"]
        self.assertEqual(ws.max_row, 5)  # header + 4 servicers

        servicer_names = [ws.cell(row=r, column=1).value for r in range(2, 6)]
        self.assertEqual(servicer_names, ["Community Bank and Trust", "Lendistry 2025", "Lendistry 2024", "Peoples Bank"])

    def test_note_series_servicers_remain_distinct_rows(self):
        # Same underlying servicer ("Lendistry"), two different note-series
        # labels, must appear as two SEPARATE rows with their own figures,
        # not merged or overwritten.
        records = [
            _make_record("Lendistry 2025", outstanding=5000000.0),
            _make_record("Lendistry 2024", outstanding=4500000.0),
        ]
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes(records, buf)
        wb = load_workbook(buf)
        ws = wb["Portfolio Summary Dollars"]
        self.assertEqual(ws.cell(row=2, column=1).value, "Lendistry 2025")
        self.assertEqual(ws.cell(row=2, column=2).value, 5000000.0)
        self.assertEqual(ws.cell(row=3, column=1).value, "Lendistry 2024")
        self.assertEqual(ws.cell(row=3, column=2).value, 4500000.0)

    def test_dollar_and_percentage_sheets_have_same_servicer_order(self):
        records = [
            _make_record("Bank A", outstanding=1000.0),
            _make_record("Bank B", outstanding=2000.0),
            _make_record("Bank C", outstanding=3000.0),
        ]
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes(records, buf)
        wb = load_workbook(buf)
        dollars_ws = wb["Portfolio Summary Dollars"]
        pct_ws = wb["Portfolio Summary Percentages"]
        dollar_order = [dollars_ws.cell(row=r, column=1).value for r in range(2, 5)]
        pct_order = [pct_ws.cell(row=r, column=1).value for r in range(2, 5)]
        self.assertEqual(dollar_order, pct_order)
        self.assertEqual(dollar_order, ["Bank A", "Bank B", "Bank C"])


class PortfolioSummaryMissingCategoryTests(unittest.TestCase):
    """Categories with no balance must display zero rather than being
    omitted from the row."""

    def test_missing_categories_display_zero_not_blank(self):
        # A record with every category at 0 except Outstanding
        records = [_make_record("Test Bank", outstanding=50000.0)]
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes(records, buf)
        wb = load_workbook(buf)
        ws = wb["Portfolio Summary Dollars"]
        row_values = [ws.cell(row=2, column=c).value for c in range(2, len(report.PORTFOLIO_SUMMARY_CATEGORIES) + 2)]
        # First is Outstanding (50000.0), the rest must be 0 (not None, not "")
        self.assertEqual(row_values[0], 50000.0)
        for v in row_values[1:]:
            self.assertEqual(v, 0.0)
            self.assertIsNotNone(v)

    def test_record_missing_a_category_key_entirely_still_shows_zero(self):
        # A record that's missing a category key ENTIRELY (not just set to
        # 0) must still show 0 in the workbook, not raise a KeyError or
        # leave a blank cell.
        incomplete_record = {"Servicer": "Sparse Bank", "Outstanding ($)": 1000.0, "Outstanding (%)": 100.0}
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes([incomplete_record], buf)
        wb = load_workbook(buf)
        ws = wb["Portfolio Summary Dollars"]
        current_col = report.PORTFOLIO_SUMMARY_CATEGORIES.index("Current") + 2
        self.assertEqual(ws.cell(row=2, column=current_col).value, 0)


class PortfolioSummaryOutstandingPercentageTests(unittest.TestCase):
    """Outstanding percentage displays 100.00 percent whenever Outstanding
    is greater than zero."""

    def test_outstanding_percentage_is_100_when_nonzero(self):
        records = [_make_record("Test Bank", outstanding=12345.67)]
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes(records, buf)
        wb = load_workbook(buf)
        ws = wb["Portfolio Summary Percentages"]
        self.assertEqual(ws.cell(row=2, column=2).value, 100.0)
        self.assertEqual(ws.cell(row=2, column=2).number_format, '0.00"%"')

    def test_outstanding_percentage_is_zero_when_outstanding_is_zero(self):
        records = [_make_record("Empty Bank", outstanding=0.0)]
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes(records, buf)
        wb = load_workbook(buf)
        ws = wb["Portfolio Summary Percentages"]
        self.assertEqual(ws.cell(row=2, column=2).value, 0.0)


class DetailedReconciliationWorkbookUnaffectedTests(unittest.TestCase):
    """The detailed reconciliation workbook must remain completely
    unaffected by any of the Portfolio Summary workbook changes — no
    PORT-01 tab, and the existing currency auto-detection in _write_rows
    must behave identically to before."""

    def test_write_rows_currency_auto_detection_unchanged(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        headers = ["SBA GP Number", "Ending Balance", "Payment Received", "Status"]
        rows = [["123", 50000.0, 500.0, "ACTIVE"]]
        report._write_header(ws, headers)
        report._write_rows(ws, headers, rows)  # no explicit currency_cols/percentage_cols
        self.assertEqual(ws.cell(row=2, column=2).number_format, '"$"#,##0.00')
        self.assertEqual(ws.cell(row=2, column=3).number_format, '"$"#,##0.00')
        self.assertEqual(ws.cell(row=2, column=4).number_format, "General")

    def test_detailed_workbook_generation_still_works_with_no_percentage_args(self):
        # generate_report_bytes never passes currency_cols/percentage_cols
        # -- confirms the detailed workbook's own call path is untouched.
        import checks
        result = {
            "id": "BAL-01::Test Bank", "servicer": "Test Bank", "name": "Balance Continuity",
            "category": "SBA Reconciliation", "status": "PASS", "issues_found": 0,
            "details": "No issues found.", "records": [],
        }
        summary = {"servicers_processed": ["Test Bank"], "soup_records": 0, "servicer_records": 0, "total_issues": 0}
        buf = io.BytesIO()
        report.generate_report_bytes([result], summary, buf)
        wb = load_workbook(buf)
        self.assertNotIn("Portfolio Summary Dollars", wb.sheetnames)
        self.assertNotIn("Portfolio Summary Percentages", wb.sheetnames)
        # confirm no PORT-01-named tab exists either
        self.assertFalse(any("PORT" in s.upper() for s in wb.sheetnames))


class PortfolioSummaryFilenameTests(unittest.TestCase):
    """Filename reflects the actual reporting period when available,
    falling back to today's date otherwise."""

    def test_filename_uses_cycle_month_when_available(self):
        import pandas as pd
        cycle_month = pd.Timestamp("2026-06-01")
        self.assertEqual(report.portfolio_summary_filename(cycle_month), "Portfolio_Summary_June_2026.xlsx")

    def test_filename_falls_back_to_today_when_no_cycle_month(self):
        from datetime import date
        filename = report.portfolio_summary_filename(None)
        self.assertEqual(filename, f"Portfolio_Summary_{date.today().isoformat()}.xlsx")


if __name__ == "__main__":
    unittest.main()
