"""POST /api/corrections/review returns the Final Audited Portfolio Summary
(client-confirmed 2026-08-10).

Generated ONLY after Manual Review, from:

    original lender data + verified_overrides + closed-at-zero suppression

and only through checks.verified_portfolio_summary_record, which reuses
_build_portfolio_summary_record. No separate calculation engine; the Original
(PORT-01) and Audited summaries are untouched.

Two properties drive the design and are asserted throughout:

  * ONE ROW PER UPLOADED SERVICER, including servicers the reviewer made no
    decisions on. Their overrides are empty, so their row is PORT-01 minus any
    suppressed loans. Before this, session state was stashed only when a
    servicer had candidates, so those servicers vanished from the summary.
  * THE CORRECTED WORKBOOK AND THE SUMMARY SHARE ONE SOURCE. Both are built
    from the same verified_overrides map, populated at a single point in
    generate_reviewed_workbook, so they cannot disagree.

Run with:
    python -m pytest test_final_portfolio_summary_endpoint.py
"""

import asyncio
import base64
import io
import json
import os
import unittest

import pandas as pd
from openpyxl import Workbook, load_workbook

import checks
import main


HEADER_ROW = 9


def _srf_workbook_bytes(rows):
    """A workbook shaped like the real SRF: header on row 9, one row per loan."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Servicer Inputs"
    headers = ["SBA GP Number", "Lender Loan Number", "Loan Status", "Delinquency Status",
               "Payment Received", "Ending Balance", "Beginning Balance", "Scheduled Payment"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=HEADER_ROW, column=col, value=h)
    for i, (gp, ls, dq, rec, eb, bb) in enumerate(rows):
        r = HEADER_ROW + 1 + i
        for col, v in enumerate([gp, f"L-{gp}", ls, dq, rec, eb, bb, 500.0], start=1):
            ws.cell(row=r, column=col, value=v)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _srf_df(rows):
    return pd.DataFrame([{
        "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}", "Loan Status": ls,
        "Delinquency Status": dq, "Payment Received": rec, "Ending Balance": eb,
        "Beginning Balance": bb, "Scheduled Payment": 500.0, "Servicing Fee": 0.0,
        "_excel_row": HEADER_ROW + 1 + i, "_sheet_name": "Servicer Inputs",
        "_header_excel_row": HEADER_ROW,
    } for i, (gp, ls, dq, rec, eb, bb) in enumerate(rows)])


BANK_A_ROWS = [("1000000001", "ACTIVE", "CURRENT", 500.0, 100000.0, 100500.0),
               ("1000000002", "ACTIVE", "CURRENT", 0.0, 50000.0, 50000.0)]
BANK_B_ROWS = [("2000000001", "ACTIVE", "CURRENT", 500.0, 75000.0, 75500.0)]


def _candidate(gp, field, corrected_value, rec_id, excel_row):
    return {
        "id": rec_id, "gp": gp, "field": field, "corrected_value": corrected_value,
        "reported_value": "ACTIVE" if field == "Loan Status" else "CURRENT",
        "excel_rows": [excel_row], "skipped_row_refs": [],
        "sheet_name": "Servicer Inputs", "header_excel_row": HEADER_ROW,
        "eligibility_rule": "test", "review_required": False, "review_reason": None,
    }


def _session(candidates_a=None, excluded_a=None, include_bank_b=True):
    """Builds a SESSION_CACHE entry the way /api/audit now does: an entry for
    EVERY servicer, whether or not it has candidates."""
    state = {
        "Bank A": {
            "original_bytes": _srf_workbook_bytes(BANK_A_ROWS),
            "original_filename": "Bank A||Bank A__srf.xlsx",
            "source_sha256": "sha-a", "institution": "Bank A",
            "cycle_month": pd.Timestamp("2026-06-01"),
            "candidates": candidates_a if candidates_a is not None else [],
            "excluded_gps": excluded_a or [],
        },
    }
    if include_bank_b:
        state["Bank B"] = {
            "original_bytes": _srf_workbook_bytes(BANK_B_ROWS),
            "original_filename": "Bank B||Bank B__srf.xlsx",
            "source_sha256": "sha-b", "institution": "Bank B",
            "cycle_month": pd.Timestamp("2026-06-01"),
            "candidates": [],          # nothing reviewable for this servicer
            "excluded_gps": [],
        }
    sid = "test-session"
    main.SESSION_CACHE[sid] = {"soup_df": None, "created": 9e18, "corrections": state}
    return sid


def _post(payload):
    response = asyncio.new_event_loop().run_until_complete(main.corrections_review(payload))
    return response.status_code, json.loads(bytes(response.body).decode())


def _section(b64, title="BALANCE ($)"):
    """Reads one stacked section of the single-worksheet Final Audited
    Portfolio Summary. Returns (header, total_row, servicer_rows) with the
    Total row separated out, since it is written FIRST in every section."""
    ws = load_workbook(io.BytesIO(base64.b64decode(b64))).active
    col_a = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    start = col_a.index(title)                      # 0-based; sheet row is start + 1
    header = [ws.cell(row=start + 2, column=c).value
              for c in range(1, ws.max_column + 1)]
    rows = []
    r = start + 3
    while r <= ws.max_row and ws.cell(row=r, column=1).value:
        rows.append([ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)])
        r += 1
    return header, rows[0], rows[1:]


def _summary_rows(b64):
    """Back-compatible shim: the BALANCE ($) section's servicer rows, Total
    excluded, so the dollar assertions below read exactly as before."""
    header, _total, servicer_rows = _section(b64)
    return header, servicer_rows


def _row(body, servicer, section="BALANCE ($)"):
    header, _total, rows = _section(body["portfolio_summary_final_b64"], section)
    return dict(zip(header, next(r for r in rows if r[0] == servicer)))


def _total(body, section="BALANCE ($)"):
    header, total_row, _rows = _section(body["portfolio_summary_final_b64"], section)
    return dict(zip(header, total_row))


class ServicerCoverageTests(unittest.TestCase):
    """Every servicer appears exactly once, even with no decisions."""

    def setUp(self):
        main.SESSION_CACHE.clear()

    def test_a_servicer_with_no_corrections_is_still_included(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        status, body = _post({"session_id": sid, "reviewed_by": "QA",
                              "decisions": [{"id": "a1", "decision": "approve"}]})
        self.assertEqual(status, 200)
        _, data = _summary_rows(body["portfolio_summary_final_b64"])
        self.assertEqual([r[0] for r in data], ["Bank A", "Bank B"])

    def test_every_servicer_appears_exactly_once(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        _, data = _summary_rows(body["portfolio_summary_final_b64"])
        names = [r[0] for r in data]
        self.assertEqual(len(names), len(set(names)))

    def test_row_order_follows_upload_order(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        _, data = _summary_rows(body["portfolio_summary_final_b64"])
        self.assertEqual([r[0] for r in data],
                         list(main.SESSION_CACHE["test-session"]["corrections"].keys()))

    def test_the_unreviewed_servicers_row_equals_its_original_summary(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        bank_b = _row(body, "Bank B")
        original = checks.portfolio_summary_check(_srf_df(BANK_B_ROWS), "Bank B")["records"][0]
        self.assertEqual(bank_b["Outstanding"], original["Outstanding ($)"])
        self.assertEqual(bank_b["Current"], original["Current ($)"])


class VerifiedValuesDriveTheSummaryTests(unittest.TestCase):

    def setUp(self):
        main.SESSION_CACHE.clear()

    def test_an_approved_correction_moves_the_bucket(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        a = _row(body, "Bank A")
        self.assertEqual(a["60 Days"], 100000.0)
        self.assertEqual(a["Current"], 50000.0)          # the other loan only

    def test_approval_with_client_substitution_uses_the_CLIENTS_value(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve",
                                        "approved_value": "30 DAYS"}]})
        a = _row(body, "Bank A")
        self.assertEqual(a["30 Days"], 100000.0)         # the client's value
        self.assertEqual(a["60 Days"], 0.0)              # NOT the audit's

    def test_a_rejected_recommendation_leaves_the_lender_value(self):
        sid = _session(candidates_a=[
            _candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10),
            _candidate("1000000002", "Delinquency Status", "90 DAYS", "a2", 11)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"},
                                       {"id": "a2", "decision": "reject"}]})
        a = _row(body, "Bank A")
        self.assertEqual(a["60 Days"], 100000.0)         # approved
        self.assertEqual(a["90 Days"], 0.0)              # rejected -- not applied
        self.assertEqual(a["Current"], 50000.0)          # keeps the lender's CURRENT

    def test_an_approved_loan_status_correction_moves_the_status_buckets(self):
        sid = _session(candidates_a=[_candidate("1000000002", "Loan Status", "DEFERRED", "a1", 11)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        self.assertEqual(_row(body, "Bank A")["Deferred"], 50000.0)


class SuppressionTests(unittest.TestCase):

    def setUp(self):
        main.SESSION_CACHE.clear()

    def test_closed_at_zero_loans_are_excluded_from_the_final_summary(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)],
                       excluded_a=["1000000002"])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        a = _row(body, "Bank A")
        self.assertEqual(a["Outstanding"], 100000.0)     # the suppressed loan is gone
        self.assertEqual(a["Current"], 0.0)
        self.assertEqual(a["Not Making Payment"], 0.0)   # was the suppressed loan's


class ArtifactsShareOneSourceTests(unittest.TestCase):
    """The corrected workbook and the Final summary are built from the same
    verified_overrides. Asserted against the real writer."""

    def setUp(self):
        main.SESSION_CACHE.clear()

    def test_the_workbook_cells_and_the_summary_agree(self):
        sid = _session(candidates_a=[
            _candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10),
            _candidate("1000000002", "Delinquency Status", "90 DAYS", "a2", 11)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"},
                                       {"id": "a2", "decision": "reject"}]})
        cw = next(c for c in body["corrected_workbooks"] if c["servicer"] == "Bank A")
        ws = load_workbook(io.BytesIO(base64.b64decode(cw["file_b64"])))["Servicer Inputs"]
        self.assertEqual(ws.cell(row=10, column=4).value, "60 DAYS")   # approved, written
        self.assertEqual(ws.cell(row=11, column=4).value, "CURRENT")   # rejected, untouched
        a = _row(body, "Bank A")
        self.assertEqual(a["60 Days"], 100000.0)
        self.assertEqual(a["Current"], 50000.0)
        self.assertEqual(a["90 Days"], 0.0)

    def test_the_corrected_workbook_download_is_unchanged(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        cw = body["corrected_workbooks"][0]
        for key in ("servicer", "filename", "file_b64", "applied_count",
                    "rejected_count", "returned_count"):
            self.assertIn(key, cw)


class ResponseAndErrorContractTests(unittest.TestCase):

    def setUp(self):
        main.SESSION_CACHE.clear()

    def test_the_filename_names_the_layer_and_the_period(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        self.assertEqual(body["portfolio_summary_final_filename"],
                         "Portfolio_Summary_Final_Audited_June_2026.xlsx")

    def test_the_workbook_is_one_sheet_with_three_stacked_sections(self):
        # CHANGED 2026-08-10: the Final Audited workbook moved from the
        # two-sheet Audited layout to ONE worksheet carrying COUNT, BALANCE %
        # and BALANCE ($) stacked, Total first in each. The Audited workbook
        # keeps the two-sheet layout, asserted separately.
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        wb = load_workbook(io.BytesIO(base64.b64decode(body["portfolio_summary_final_b64"])))
        self.assertEqual(wb.sheetnames, ["Final Audited Portfolio Summary"])
        col_a = [wb.active.cell(row=r, column=1).value for r in range(1, wb.active.max_row + 1)]
        self.assertEqual([v for v in col_a if v in ("COUNT", "BALANCE %", "BALANCE ($)")],
                         ["COUNT", "BALANCE %", "BALANCE ($)"])

    def test_all_three_sections_carry_the_same_servicers(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        names = []
        for section in ("COUNT", "BALANCE %", "BALANCE ($)"):
            _h, _t, rows = _section(body["portfolio_summary_final_b64"], section)
            names.append([r[0] for r in rows])
        self.assertEqual(names[0], names[1])
        self.assertEqual(names[1], names[2])
        self.assertEqual(names[0], ["Bank A", "Bank B"])

    def test_the_count_section_agrees_with_the_dollar_section(self):
        # An approved correction must move Count and Dollars together.
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        counts = _row(body, "Bank A", "COUNT")
        dollars = _row(body, "Bank A", "BALANCE ($)")
        self.assertEqual(counts["60 Days"], 1)
        self.assertEqual(dollars["60 Days"], 100000.0)

    def test_the_total_row_sums_counts_and_dollars(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        for section, category in (("COUNT", "Outstanding"), ("BALANCE ($)", "Outstanding")):
            _h, total_row, rows = _section(body["portfolio_summary_final_b64"], section)
            header, _t, _r = _section(body["portfolio_summary_final_b64"], section)
            idx = header.index(category)
            with self.subTest(section=section):
                self.assertAlmostEqual(total_row[idx], sum(r[idx] for r in rows), places=2)

    def test_total_outstanding_percentage_is_100(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        _, body = _post({"session_id": sid, "reviewed_by": "QA",
                         "decisions": [{"id": "a1", "decision": "approve"}]})
        self.assertEqual(_total(body, "BALANCE %")["Outstanding"], 100.0)

    def test_zero_decisions_behaviour_is_unchanged(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        status, body = _post({"session_id": sid, "reviewed_by": "QA", "decisions": []})
        self.assertEqual(status, 400)
        self.assertIn("No decisions were submitted", body["error"])
        self.assertNotIn("portfolio_summary_final_b64", body)

    def test_expired_session_behaviour_is_unchanged(self):
        status, body = _post({"session_id": "does-not-exist", "reviewed_by": "QA",
                              "decisions": [{"id": "a1", "decision": "approve"}]})
        self.assertEqual(status, 410)
        self.assertIn("expired", body["error"])
        self.assertNotIn("portfolio_summary_final_b64", body)

    def test_missing_reviewer_name_behaviour_is_unchanged(self):
        sid = _session(candidates_a=[_candidate("1000000001", "Delinquency Status", "60 DAYS", "a1", 10)])
        status, body = _post({"session_id": sid, "reviewed_by": "  ",
                              "decisions": [{"id": "a1", "decision": "approve"}]})
        self.assertEqual(status, 400)
        self.assertIn("reviewer name", body["error"])

    def test_an_audit_with_no_candidates_anywhere_still_reports_nothing_reviewable(self):
        # The guard now checks for CANDIDATES, not merely a session entry,
        # since every servicer is stashed. Original meaning preserved.
        sid = _session(candidates_a=[])
        status, body = _post({"session_id": sid, "reviewed_by": "QA",
                              "decisions": [{"id": "x", "decision": "approve"}]})
        self.assertEqual(status, 400)
        self.assertIn("no reviewable corrections", body["error"])


class FrontendContractTests(unittest.TestCase):
    """The download button keys on these exact response fields."""

    JSX = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "..", "frontend", "src", "components", "CorrectionReview.jsx")

    def test_the_jsx_reads_the_fields_the_endpoint_returns(self):
        src = open(self.JSX).read()
        self.assertIn("portfolio_summary_final_b64", src)
        self.assertIn("portfolio_summary_final_filename", src)
        self.assertIn("Download Final Audited Portfolio Summary", src)

    def test_the_corrected_file_download_is_still_present(self):
        self.assertIn("Download Corrected File", open(self.JSX).read())


if __name__ == "__main__":
    unittest.main()
