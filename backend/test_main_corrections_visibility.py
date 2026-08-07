"""
Integration tests for the review-then-generate workflow across the real
FastAPI endpoints (via TestClient), with only file parsing monkeypatched —
everything downstream (checks.py, corrections.py, report.py, main.py's own
response-building and the SESSION_CACHE round-trip) runs unmodified.

REWRITTEN for the workflow restructure (BRIEF_correction_workflow_restructure):

  1. /api/audit returns RECOMMENDATIONS (not a generated workbook), and an
     eligible-but-unreviewable loan (no valid posting row) still surfaces in
     corrections_skipped — never as a recommendation, never applied.
  2. POST /api/corrections/review takes a reviewer's decisions for that same
     session and generates the corrected workbook from the ORIGINAL bytes held
     in the session, applying only approved values.

Run with:
    python -m pytest test_main_corrections_visibility.py
"""

import io
import unittest

import pandas as pd
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook
import base64

import file_readers
import main
import corrections


HEADERS = ["SBA GP Number", "Loan Status", "Delinquency Status", "Beginning Balance",
           "Ending Balance", "Payment Received", "Scheduled Payment"]
HEADER_ROW = 3
SHEET_NAME = "Servicer Inputs"


def _build_servicer_workbook_bytes(loan1_rows):
    """Real .xlsx bytes for loan '1' only (the loan expected to succeed) —
    corrections.py actually opens and writes to this file, so it needs to
    genuinely match the row/column layout the crafted servicer_inputs_df
    below claims for loan '1'. Loan '2' (the unreviewable case) never gets a
    real row here."""
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


class ReviewWorkflowEndToEndTests(unittest.TestCase):
    def setUp(self):
        # loan "1": 3 postings with real row locations — reviewable and, when
        # approved, all 3 rows corrected.
        loan1_postings = [
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000.0, "Ending Balance": 100000.0, "Payment Received": 0.0,
             "Scheduled Payment": 500.0},
        ] * 3
        self.file_bytes, excel_rows = _build_servicer_workbook_bytes(loan1_postings)

        servicer_inputs_rows = []
        for i, row in enumerate(loan1_postings):
            servicer_inputs_rows.append({
                **row, "Loan Payment Instance": i + 1,
                "_excel_row": excel_rows[i], "_sheet_name": SHEET_NAME, "_header_excel_row": HEADER_ROW,
            })
        # loan "2": one posting, DELIBERATELY missing _excel_row/_sheet_name/
        # _header_excel_row -- a row file_readers couldn't tag with a physical
        # location. Eligible by every other rule, but no valid row to write to,
        # so it is NOT reviewable.
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
        self.loan_terms_df = pd.DataFrame([{"SBA GP Number": "1"}, {"SBA GP Number": "2"}])

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

    def _run_audit(self):
        return self.client.post(
            "/api/audit",
            files={
                "soup_file": ("soup.xlsx", b"dummy soup bytes", "application/octet-stream"),
                "servicer_files": ("servicer.xlsx", self.file_bytes, "application/octet-stream"),
            },
        )

    def test_audit_returns_recommendations_not_a_generated_workbook(self):
        response = self._run_audit()
        self.assertEqual(response.status_code, 200)
        body = response.json()

        # No workbook is generated during the audit anymore.
        self.assertNotIn("corrected_workbooks", body)
        self.assertIn("recommendations", body)
        self.assertIn("session_id", body)

        # loan "1" is a reviewable recommendation; loan "2" is not.
        recs = body["recommendations"]
        self.assertEqual(len(recs), 1)
        rec = recs[0]
        self.assertEqual(rec["gp"], "1")
        self.assertEqual(rec["field"], "Delinquency Status")
        self.assertEqual(rec["servicer"], "Test Bank")
        self.assertEqual(rec["postings"], 3)
        self.assertEqual(sorted(rec["rows"]), sorted([HEADER_ROW + 1, HEADER_ROW + 2, HEADER_ROW + 3]))

        # loan "2" (no valid posting row) surfaces in corrections_skipped —
        # with servicer, SBA GP Number, field, reason — never as a recommendation.
        gp2 = [s for s in body["corrections_skipped"] if s.get("SBA GP Number") == "2"]
        self.assertEqual(len(gp2), 1, f"expected one skipped entry for loan 2, got {body['corrections_skipped']}")
        self.assertEqual(gp2[0]["servicer"], "Test Bank")
        self.assertEqual(gp2[0]["field"], "Delinquency Status")
        self.assertTrue(gp2[0]["reason"])
        self.assertFalse(any(r["gp"] == "2" for r in recs))

    def test_review_generates_corrected_workbook_from_original_bytes(self):
        audit = self._run_audit().json()
        session_id = audit["session_id"]
        rec = audit["recommendations"][0]

        review = self.client.post("/api/corrections/review", json={
            "session_id": session_id,
            "reviewed_by": "K. Dutta",
            "decisions": [{"id": rec["id"], "decision": "approve", "reviewer_note": "confirmed"}],
        })
        self.assertEqual(review.status_code, 200)
        body = review.json()
        self.assertEqual(body["summary"]["applied"], 1)
        self.assertEqual(len(body["corrected_workbooks"]), 1)
        cw = body["corrected_workbooks"][0]
        self.assertEqual(cw["servicer"], "Test Bank")
        self.assertEqual(cw["applied_count"], 1)

        # The generated workbook actually carries the correction on all 3 rows
        # and a Correction Log attributing it to the (self-declared) reviewer.
        # The written value is STAT-01's real calculated value (this test runs
        # the actual engine), which approving defaults to.
        calculated = rec["calculated"]
        self.assertTrue(calculated)
        wb = load_workbook(io.BytesIO(base64.b64decode(cw["file_b64"])), data_only=False)
        ws = wb[SHEET_NAME]
        for r in (HEADER_ROW + 1, HEADER_ROW + 2, HEADER_ROW + 3):
            self.assertEqual(ws.cell(row=r, column=HEADERS.index("Delinquency Status") + 1).value, calculated)
        log = wb[corrections.CORRECTION_LOG_SHEET_NAME]
        log_headers = [c.value for c in log[1]]
        self.assertEqual(log_headers, corrections.CORRECTION_LOG_HEADERS)
        reviewed_by = log.cell(row=2, column=log_headers.index("Reviewed By (self-declared)") + 1).value
        self.assertEqual(reviewed_by, "K. Dutta (unverified)")
        sha_col = log.cell(row=2, column=log_headers.index("Source File SHA-256") + 1).value
        self.assertTrue(sha_col)  # source hash recorded

    def test_review_requires_a_reviewer_name(self):
        audit = self._run_audit().json()
        rec = audit["recommendations"][0]
        resp = self.client.post("/api/corrections/review", json={
            "session_id": audit["session_id"],
            "reviewed_by": "",
            "decisions": [{"id": rec["id"], "decision": "approve"}],
        })
        self.assertEqual(resp.status_code, 400)
        self.assertIn("reviewer name", resp.json()["error"].lower())

    def test_review_on_unknown_session_is_rejected(self):
        resp = self.client.post("/api/corrections/review", json={
            "session_id": "does-not-exist",
            "reviewed_by": "K. Dutta",
            "decisions": [{"id": "x||y||z", "decision": "approve"}],
        })
        self.assertEqual(resp.status_code, 410)
        self.assertIn("expired", resp.json()["error"].lower())

    def test_rejected_decision_writes_no_cell_but_is_logged(self):
        audit = self._run_audit().json()
        rec = audit["recommendations"][0]
        review = self.client.post("/api/corrections/review", json={
            "session_id": audit["session_id"],
            "reviewed_by": "K. Dutta",
            "decisions": [{"id": rec["id"], "decision": "reject"}],
        }).json()
        self.assertEqual(review["summary"]["applied"], 0)
        self.assertEqual(review["summary"]["rejected"], 1)
        cw = review["corrected_workbooks"][0]
        wb = load_workbook(io.BytesIO(base64.b64decode(cw["file_b64"])), data_only=False)
        ws = wb[SHEET_NAME]
        # cells left exactly as reported
        for r in (HEADER_ROW + 1, HEADER_ROW + 2, HEADER_ROW + 3):
            self.assertEqual(ws.cell(row=r, column=HEADERS.index("Delinquency Status") + 1).value, "CURRENT")


if __name__ == "__main__":
    unittest.main()
