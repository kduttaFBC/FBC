"""
Integration test proving corrections_not_applied actually reaches the
/api/audit response, not just corrections.py's return value.

CLIENT-CONFIRMED (2026-08-05): a correction that matched every business
eligibility rule but had zero valid posting rows to write to must surface
to the user through the existing corrections_skipped structure — with the
servicer, SBA GP Number, field, and reason — and must never be counted in
correction_count.

This hits the real FastAPI /api/audit endpoint (not just corrections.py in
isolation) via TestClient, with file_readers.read_soup_file /
read_servicer_file / filter_soup_by_servicer monkeypatched to return
controlled DataFrames instead of actually parsing multipart bytes --
everything downstream (checks.py, corrections.py, report.py, main.py's own
response-building logic) runs for real, unmodified.

Run with:
    python -m unittest test_main_corrections_visibility.py
"""

import io
import unittest

import pandas as pd
from fastapi.testclient import TestClient
from openpyxl import Workbook

import file_readers
import main


HEADERS = ["SBA GP Number", "Loan Status", "Delinquency Status", "Beginning Balance",
           "Ending Balance", "Payment Received", "Scheduled Payment"]
HEADER_ROW = 3
SHEET_NAME = "Servicer Inputs"


def _build_servicer_workbook_bytes(loan1_rows):
    """Real .xlsx bytes for loan '1' only (the loan expected to succeed) —
    corrections.py actually opens and writes to this file, so it needs to
    genuinely match the row/column layout the crafted servicer_inputs_df
    below claims for loan '1'. Loan '2' (the zero-valid-row case) never
    gets written to at all, so it doesn't need a real row here."""
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.cell(row=1, column=1, value="Servicer Name: Test Bank")
    ws.cell(row=2, column=1, value="Reporting Period: June 2026")
    for col, h in enumerate(HEADERS, start=1):
        ws.cell(row=HEADER_ROW, column=col, value=h)
    excel_rows = []
    for i, row in enumerate(loan1_rows):
        r = HEADER_ROW + 1 + i
        excel_rows.append(r)
        for col, h in enumerate(HEADERS, start=1):
            ws.cell(row=r, column=col, value=row.get(h))
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), excel_rows


class CorrectionsNotAppliedReachesApiResponseTests(unittest.TestCase):
    def setUp(self):
        # loan "1": 3 postings, real row locations in the uploaded xlsx —
        # expected to succeed (all 3 rows corrected).
        loan1_postings = [
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000.0, "Ending Balance": 100000.0, "Payment Received": 0.0,
             "Scheduled Payment": 500.0},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000.0, "Ending Balance": 100000.0, "Payment Received": 0.0,
             "Scheduled Payment": 500.0},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000.0, "Ending Balance": 100000.0, "Payment Received": 0.0,
             "Scheduled Payment": 500.0},
        ]
        self.file_bytes, excel_rows = _build_servicer_workbook_bytes(loan1_postings)

        servicer_inputs_rows = []
        for i, row in enumerate(loan1_postings):
            servicer_inputs_rows.append({
                **row, "Loan Payment Instance": i + 1,
                "_excel_row": excel_rows[i], "_sheet_name": SHEET_NAME, "_header_excel_row": HEADER_ROW,
            })
        # loan "2": one posting, DELIBERATELY missing _excel_row/_sheet_name/
        # _header_excel_row -- simulates a row file_readers couldn't safely
        # tag with a physical location. Fully eligible for a Delinquency
        # Status correction by every other rule, but zero valid rows to
        # write to.
        servicer_inputs_rows.append({
            "SBA GP Number": "2", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
            "Beginning Balance": 50000.0, "Ending Balance": 50000.0, "Payment Received": 0.0,
            "Scheduled Payment": 500.0, "Loan Payment Instance": 1,
        })
        self.servicer_inputs_df = pd.DataFrame(servicer_inputs_rows)

        self.soup_df = pd.DataFrame([
            {"SBA GP Number": "1", "Servicer": "Test Bank", "Payment Month": "2026-05-01",
             "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Actual Balance": 100000.0, "Ending Actual Balance": 100000.0},
            {"SBA GP Number": "2", "Servicer": "Test Bank", "Payment Month": "2026-05-01",
             "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Actual Balance": 50000.0, "Ending Actual Balance": 50000.0},
        ])
        self.loan_terms_df = pd.DataFrame([
            {"SBA GP Number": "1"}, {"SBA GP Number": "2"},
        ])

        # Monkeypatch file parsing only -- everything downstream (checks.py,
        # corrections.py, report.py, main.py's own logic) runs unmodified.
        self._orig_read_soup_file = file_readers.read_soup_file
        self._orig_read_servicer_file = file_readers.read_servicer_file
        self._orig_filter_soup = file_readers.filter_soup_by_servicer

        file_readers.read_soup_file = lambda raw_bytes: self.soup_df
        file_readers.read_servicer_file = lambda raw_bytes, filename=None: (
            "Test Bank", "Test Bank", self.loan_terms_df, self.servicer_inputs_df, None
        )
        file_readers.filter_soup_by_servicer = lambda soup_df, soup_key: soup_df

        self.client = TestClient(main.app)

    def tearDown(self):
        file_readers.read_soup_file = self._orig_read_soup_file
        file_readers.read_servicer_file = self._orig_read_servicer_file
        file_readers.filter_soup_by_servicer = self._orig_filter_soup

    def test_zero_valid_row_correction_surfaces_in_response_and_is_not_counted(self):
        response = self.client.post(
            "/api/audit",
            files={
                "soup_file": ("soup.xlsx", b"dummy soup bytes", "application/octet-stream"),
                "servicer_files": ("servicer.xlsx", self.file_bytes, "application/octet-stream"),
            },
        )
        self.assertEqual(response.status_code, 200)
        body = response.json()

        # loan "1" succeeded: exactly one corrected workbook, correction_count == 1
        self.assertEqual(len(body["corrected_workbooks"]), 1)
        self.assertEqual(body["corrected_workbooks"][0]["correction_count"], 1)
        self.assertEqual(body["corrected_workbooks"][0]["servicer"], "Test Bank")

        # loan "2" (zero valid rows) must appear in corrections_skipped —
        # with servicer, SBA GP Number, field, and reason — not silently
        # dropped, and NOT folded into correction_count above.
        skipped = body["corrections_skipped"]
        gp2_entries = [s for s in skipped if s.get("SBA GP Number") == "2"]
        self.assertEqual(len(gp2_entries), 1, f"expected exactly one skipped entry for loan 2, got: {skipped}")
        entry = gp2_entries[0]
        self.assertEqual(entry["servicer"], "Test Bank")
        self.assertEqual(entry["field"], "Delinquency Status")
        self.assertTrue(entry["reason"])  # non-empty, human-readable reason present

        # loan "2" must never appear as a corrected workbook, and no
        # corrected workbook's correction_count may include it. With only
        # loan "1" eligible for a real write, correction_count must be
        # exactly 1 — loan 2 contributing would make this 2.
        self.assertEqual(sum(cw["correction_count"] for cw in body["corrected_workbooks"]), 1)
        # And loan "2" must not appear as its own corrected workbook entry.
        self.assertFalse(any(cw.get("gp") == "2" for cw in body["corrected_workbooks"]))


if __name__ == "__main__":
    unittest.main()
