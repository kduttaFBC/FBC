"""
Tests for the review-then-generate correction workflow (corrections.py).

REWRITTEN for the workflow restructure (BRIEF_correction_workflow_restructure):
the audit no longer writes a corrected workbook. It produces RECOMMENDATIONS
(build_recommendations); a reviewer then decides each one, and generate_
reviewed_workbook writes ONLY the approved values from a fresh copy of the
original bytes, recording every approve/reject/return in the Correction Log.

The multi-posting / dedupe / GP-safety / one-log-row-per-loan-correction
invariants from the previous auto-apply workflow all still hold — they are
retested here through the new API rather than dropped. New coverage: a
reviewer substituting a value different from the calculated one; rejected and
returned items leaving the cell untouched; all three appearing in the log with
the right Rows Updated; and the approved-value domain restriction refusing a
cross-domain substitution.

Run with:
    python -m pytest test_corrections.py
"""

import io
import os
import unittest

from openpyxl import Workbook, load_workbook
import pandas as pd

import corrections
import checks
import file_readers


HEADERS = [
    "SBA GP Number", "Loan Status", "Delinquency Status", "Beginning Balance",
    "Ending Balance", "Payment Received", "Scheduled Payment", "Net",
]
HEADER_ROW = 3  # simulates a template with title/blank rows above the header
SERVICER = "Test Bank"


def _build_workbook_bytes(rows, sheet_name="Servicer Inputs"):
    """rows: list of dicts keyed by HEADERS. 'Net' is written as a formula
    referencing Beginning/Ending Balance on its own row, so preservation
    tests have something real to check. Beginning Balance gets a currency
    number format for the same reason."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.cell(row=1, column=1, value="Servicer Name: Test Bank")  # title row filler
    ws.cell(row=2, column=1, value="Reporting Period: June 2026")
    for col, h in enumerate(HEADERS, start=1):
        ws.cell(row=HEADER_ROW, column=col, value=h)

    excel_rows = []
    for i, row in enumerate(rows):
        r = HEADER_ROW + 1 + i
        excel_rows.append(r)
        for col, h in enumerate(HEADERS, start=1):
            if h == "Net":
                ws.cell(row=r, column=col, value=f"=E{r}-D{r}")
            else:
                ws.cell(row=r, column=col, value=row.get(h))
        ws.cell(row=r, column=HEADERS.index("Beginning Balance") + 1).number_format = '"$"#,##0.00'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), excel_rows


def _read_cell(ws, row, header):
    col = HEADERS.index(header) + 1
    return ws.cell(row=row, column=col).value


def _reload(file_bytes, sheet_name="Servicer Inputs"):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
    return wb, wb[sheet_name]


def _correction_data_entry(gp, excel_rows, sheet_name="Servicer Inputs", header_excel_row=HEADER_ROW,
                           reasons=("DELINQUENCY STATUS MISMATCH",), dq_match=False,
                           calculated_dq="60 DAYS", **overrides):
    entry = {
        "gp": gp, "excel_row": excel_rows[-1], "excel_rows": excel_rows,
        "sheet_name": sheet_name, "header_excel_row": header_excel_row,
        "reasons": list(reasons), "dq_match": dq_match, "calculated_dq": calculated_dq,
        "srf_dq": "CURRENT", "status_match": True, "expected_loan_status_raw": None,
        "srf_status": "ACTIVE", "soup_status": "ACTIVE", "deferral_type": None,
        "deferral_phase": None, "reversal_equivalents": None,
        # read-only review evidence (see checks.py); irrelevant to write logic
        "prior_dq": "CURRENT", "payment_received": 0.0, "scheduled_payment": 500.0,
        "servicing_fee": 0.0, "ending_balance": 100000.0, "posting_count": len(excel_rows),
    }
    entry.update(overrides)
    return entry


# --- helpers for the two-step review workflow -----------------------------

def _recommend(correction_data, servicer=SERVICER, cycle_month=None):
    stat01 = {"correction_data": correction_data, "_cycle_month": cycle_month}
    return corrections.build_recommendations(stat01, servicer)


def _approve(rec, approved_value=None, note=None, notify=False):
    d = {"id": rec["id"], "decision": "approve"}
    if approved_value is not None:
        d["approved_value"] = approved_value
    if note is not None:
        d["reviewer_note"] = note
    if notify:
        d["notify_servicer"] = True
    return d


def _generate(file_bytes, candidates, decisions, servicer=SERVICER,
              reviewer="K. Dutta", sha="sha256-fixture", cycle_month=None,
              filename="test.xlsx", reviewed_on=None):
    return corrections.generate_reviewed_workbook(
        file_bytes, filename, servicer, candidates, decisions, reviewer, sha,
        cycle_month, reviewed_on=reviewed_on,
    )


def _log_row(file_bytes, row_index=2):
    wb, _ = _reload(file_bytes)
    log_ws = wb[corrections.CORRECTION_LOG_SHEET_NAME]
    headers = [c.value for c in log_ws[1]]
    return headers, {h: log_ws.cell(row=row_index, column=i + 1).value for i, h in enumerate(headers)}, log_ws


class SinglePostingApprovedUpdatesOneRowTests(unittest.TestCase):
    """A loan with exactly one posting, approved, updates exactly one row."""

    def test_single_posting_updates_exactly_one_row(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        recs, candidates, skipped = _recommend([_correction_data_entry("1", excel_rows)])
        self.assertEqual(len(recs), 1)
        self.assertEqual(skipped, [])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        self.assertEqual(outcome["applied_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[0], "Delinquency Status"), "60 DAYS")


class ThreePostingApprovedDelinquencyTests(unittest.TestCase):
    """An approved Delinquency Status correction for a 3-posting loan updates
    all three rows, and is ONE applied correction (not three)."""

    def test_all_three_rows_get_the_corrected_delinquency_status(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ] * 3)
        self.assertEqual(len(excel_rows), 3)
        recs, candidates, _ = _recommend([_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        self.assertEqual(outcome["applied_count"], 1)  # ONE loan-level correction
        _, ws = _reload(outcome["file_bytes"])
        for r in excel_rows:
            self.assertEqual(_read_cell(ws, r, "Delinquency Status"), "60 DAYS")


class LoanStatusApprovedUpdatesAllRowsTests(unittest.TestCase):
    """An approved Loan Status correction (-> DEFERRED) updates every posting
    row for that loan."""

    def test_all_rows_get_deferred(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ] * 2)
        entry = _correction_data_entry(
            "1", excel_rows, reasons=("LOAN STATUS MISMATCH",),
            expected_loan_status_raw="DEFERRED", deferral_type="full", deferral_phase="start",
            srf_status="ACTIVE", soup_status="ACTIVE",
        )
        recs, candidates, _ = _recommend([entry])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        self.assertEqual(outcome["applied_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        for r in excel_rows:
            self.assertEqual(_read_cell(ws, r, "Loan Status"), "DEFERRED")


class FieldIndependenceTests(unittest.TestCase):
    """Correcting one field must never change the other status field."""

    def test_delinquency_correction_leaves_loan_status_untouched(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ] * 3)
        recs, candidates, _ = _recommend([_correction_data_entry("1", excel_rows, calculated_dq="90 DAYS")])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        _, ws = _reload(outcome["file_bytes"])
        for r in excel_rows:
            self.assertEqual(_read_cell(ws, r, "Delinquency Status"), "90 DAYS")
            self.assertEqual(_read_cell(ws, r, "Loan Status"), "ACTIVE")  # unchanged


class OtherLoansUntouchedTests(unittest.TestCase):
    """A correction targeting one loan must leave every other loan unchanged."""

    def test_other_loan_rows_unchanged(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "2", "Loan Status": "ACTIVE", "Delinquency Status": "30 DAYS",
             "Beginning Balance": 40000, "Ending Balance": 39000, "Payment Received": 450,
             "Scheduled Payment": 500},
        ])
        loan1_rows = excel_rows[:2]
        loan2_row = excel_rows[2]
        recs, candidates, _ = _recommend([_correction_data_entry("1", loan1_rows, calculated_dq="60 DAYS")])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, loan2_row, "SBA GP Number"), "2")
        self.assertEqual(_read_cell(ws, loan2_row, "Loan Status"), "ACTIVE")
        self.assertEqual(_read_cell(ws, loan2_row, "Delinquency Status"), "30 DAYS")
        self.assertEqual(_read_cell(ws, loan2_row, "Beginning Balance"), 40000)
        self.assertEqual(_read_cell(ws, loan2_row, "Ending Balance"), 39000)
        self.assertEqual(_read_cell(ws, loan2_row, "Payment Received"), 450)


class FormattingAndFormulasPreservedTests(unittest.TestCase):
    """Number formats, formulas, and posting order survive a correction; only
    the target status cell's value changes."""

    def test_formulas_and_number_formats_unchanged_on_corrected_and_other_rows(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 95000, "Payment Received": 500,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 95000, "Ending Balance": 90000, "Payment Received": 500,
             "Scheduled Payment": 500},
            {"SBA GP Number": "2", "Loan Status": "ACTIVE", "Delinquency Status": "30 DAYS",
             "Beginning Balance": 40000, "Ending Balance": 39000, "Payment Received": 450,
             "Scheduled Payment": 500},
        ])
        _, ws_before = _reload(file_bytes)
        before = {
            r: {
                "net_formula": ws_before.cell(row=r, column=HEADERS.index("Net") + 1).value,
                "beg_fmt": ws_before.cell(row=r, column=HEADERS.index("Beginning Balance") + 1).number_format,
                "beg_val": _read_cell(ws_before, r, "Beginning Balance"),
                "ending_val": _read_cell(ws_before, r, "Ending Balance"),
                "gp": _read_cell(ws_before, r, "SBA GP Number"),
            }
            for r in excel_rows
        }
        recs, candidates, _ = _recommend([_correction_data_entry("1", excel_rows[:2], calculated_dq="90 DAYS")])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        _, ws_after = _reload(outcome["file_bytes"])
        for r in excel_rows:
            self.assertEqual(ws_after.cell(row=r, column=HEADERS.index("Net") + 1).value, before[r]["net_formula"])
            self.assertEqual(ws_after.cell(row=r, column=HEADERS.index("Beginning Balance") + 1).number_format, before[r]["beg_fmt"])
            self.assertEqual(_read_cell(ws_after, r, "Beginning Balance"), before[r]["beg_val"])
            self.assertEqual(_read_cell(ws_after, r, "Ending Balance"), before[r]["ending_val"])
            self.assertEqual(_read_cell(ws_after, r, "SBA GP Number"), before[r]["gp"])
        self.assertEqual(excel_rows, sorted(excel_rows))


class ReuploadConsistencyTests(unittest.TestCase):
    """Re-uploading the corrected workbook produces one consistent
    consolidated status matching the correction."""

    def test_reconsolidating_the_corrected_file_yields_the_corrected_status(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ] * 3)
        recs, candidates, _ = _recommend([_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        _, ws = _reload(outcome["file_bytes"])
        reuploaded_rows = [
            {h: _read_cell(ws, r, h) for h in HEADERS if h != "Net"}
            for r in excel_rows
        ]
        si_grouped = checks.consolidate_cycle_postings(pd.DataFrame(reuploaded_rows))
        self.assertEqual(len(si_grouped), 1)
        self.assertEqual(si_grouped.loc["1", "dq_status"], "60 DAYS")
        self.assertEqual(si_grouped.loc["1", "posting_count"], 3)


class CorrectionLogTests(unittest.TestCase):
    """The Correction Log records ONE row per loan-level correction with a
    Rows Updated count, and now also the review-workflow columns."""

    def test_three_posting_correction_produces_one_log_row_with_rows_updated_three(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ] * 3)
        recs, candidates, _ = _recommend([_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0], note="looks right")],
                            reviewer="K. Dutta", sha="abc123sha")

        headers, row, log_ws = _log_row(outcome["file_bytes"])
        self.assertEqual(headers, corrections.CORRECTION_LOG_HEADERS)
        self.assertEqual(log_ws.max_row, 2)  # header + exactly ONE correction row
        self.assertEqual(row["Rows Updated"], 3)
        self.assertEqual(row["SBA GP Number"], "1")
        self.assertEqual(row["Field Updated"], "Delinquency Status")
        self.assertEqual(row["Corrected Value"], "60 DAYS")
        for r in excel_rows:
            self.assertIn(str(r), row["Excel Row Numbers"])
        self.assertEqual(row["Original Value"].count("CURRENT"), 3)
        # new review-workflow columns
        self.assertEqual(row["Calculated Value"], "60 DAYS")
        self.assertEqual(row["Reviewer Approved Value"], "60 DAYS")
        self.assertEqual(row["Decision"], "Approved and applied")
        self.assertEqual(row["Review Status"], "Applied")
        self.assertEqual(row["Reviewer Note"], "looks right")
        self.assertEqual(row["Reviewed By (self-declared)"], "K. Dutta (unverified)")
        self.assertTrue(row["Reviewed On"])
        self.assertEqual(row["Source File SHA-256"], "abc123sha")
        self.assertIn(row["Notify Servicer"], (None, ""))  # openpyxl reads a blank cell as None

    def test_applied_count_reflects_loans_not_rows(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ] * 3)
        recs, candidates, _ = _recommend([_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        self.assertEqual(outcome["applied_count"], 1)


class RowLevelSafetyStillAppliesTests(unittest.TestCase):
    """Per-row GP verification still happens for EVERY row: a stale row that
    no longer matches the loan is skipped individually, the rest still apply."""

    def test_one_mismatched_row_is_skipped_others_still_applied(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ] * 2)
        stale_row = excel_rows[-1] + 5
        recs, candidates, _ = _recommend([_correction_data_entry("1", excel_rows + [stale_row], calculated_dq="60 DAYS")])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        self.assertEqual(outcome["applied_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        for r in excel_rows:
            self.assertEqual(_read_cell(ws, r, "Delinquency Status"), "60 DAYS")


class ExcelRowsDedupeAndValidationTests(unittest.TestCase):
    """excel_rows is deduplicated (order-preserving) and blank/null/nonnumeric/
    invalid entries are dropped without guessing, before any write."""

    def test_duplicate_row_reference_written_only_once(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ] * 2)
        dup_rows = [excel_rows[0], excel_rows[0], excel_rows[1]]
        recs, candidates, _ = _recommend([_correction_data_entry("1", dup_rows, calculated_dq="60 DAYS")])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        self.assertEqual(outcome["applied_count"], 1)
        _, row, _ = _log_row(outcome["file_bytes"])
        self.assertEqual(row["Rows Updated"], 2)  # not 3 — the duplicate collapses
        self.assertIn("duplicate", row["Skipped Row References"].lower())

    def test_blank_null_nonnumeric_and_negative_refs_all_ignored(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        messy_rows = [None, "", "abc", 0, -5, 12.5, excel_rows[0]]  # only the last is valid
        recs, candidates, _ = _recommend([_correction_data_entry("1", messy_rows, calculated_dq="60 DAYS")])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        self.assertEqual(outcome["applied_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[0], "Delinquency Status"), "60 DAYS")
        _, row, _ = _log_row(outcome["file_bytes"])
        self.assertEqual(row["Rows Updated"], 1)
        self.assertEqual(row["Skipped Row References"].count("invalid row reference"), 6)

    def test_order_preserved_after_dedupe(self):
        clean, skipped = corrections._clean_excel_rows([30, 10, 30, 20, 10])
        self.assertEqual(clean, [30, 10, 20])
        self.assertEqual(len(skipped), 2)


class ZeroValidRowsNotReviewableTests(unittest.TestCase):
    """A loan eligible by every business rule but with no valid posting row to
    write to is NOT reviewable — it surfaces in `skipped`, never as a
    recommendation, and never produces a corrected row."""

    def test_all_invalid_refs_is_skipped_not_recommended(self):
        _, _ = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        recs, candidates, skipped = _recommend([_correction_data_entry("1", [None, "", "abc", -1], calculated_dq="60 DAYS")])
        self.assertEqual(recs, [])
        self.assertEqual(candidates, [])
        self.assertEqual(len(skipped), 1)
        self.assertEqual(skipped[0]["SBA GP Number"], "1")
        self.assertIn("no valid posting row", skipped[0]["reason"].lower())

    def test_approved_but_all_rows_gp_mismatched_is_not_counted_and_reported(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        # Claim the correction belongs to loan "999", which never appears in
        # the sheet — every row reference fails its GP check at write time.
        recs, candidates, _ = _recommend([_correction_data_entry("999", excel_rows, calculated_dq="60 DAYS")])
        self.assertEqual(len(recs), 1)  # it IS reviewable (has row refs); the mismatch only shows up at write
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        self.assertEqual(outcome["applied_count"], 0)
        self.assertEqual(len(outcome["corrections_not_applied"]), 1)
        self.assertEqual(outcome["corrections_not_applied"][0]["gp"], "999")
        self.assertIn("GP mismatch", outcome["corrections_not_applied"][0]["reason"])
        # the original cell is left exactly as reported
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[0], "Delinquency Status"), "CURRENT")

    def test_one_unwritable_loan_does_not_block_another_loans_valid_correction(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "2", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 50000, "Ending Balance": 50000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        broken = _correction_data_entry("1", [None, "abc"], calculated_dq="60 DAYS")  # no valid rows -> skipped
        good = _correction_data_entry("2", [excel_rows[1]], calculated_dq="90 DAYS")
        recs, candidates, skipped = _recommend([broken, good])
        self.assertEqual(len(recs), 1)  # only loan 2 is reviewable
        self.assertEqual(recs[0]["gp"], "2")
        self.assertEqual(len(skipped), 1)
        self.assertEqual(skipped[0]["SBA GP Number"], "1")
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        self.assertEqual(outcome["applied_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[1], "Delinquency Status"), "90 DAYS")


class AppliedCountIsLoanAndFieldLevelTests(unittest.TestCase):
    """applied_count is one per (loan, field), never one per physical row."""

    def test_single_multirow_correction_counts_as_one(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ] * 3)
        recs, candidates, _ = _recommend([_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0])])
        self.assertEqual(outcome["applied_count"], 1)

    def test_two_fields_same_loan_counts_as_two(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ] * 2)
        entry = _correction_data_entry(
            "1", excel_rows, reasons=("DELINQUENCY STATUS MISMATCH", "LOAN STATUS MISMATCH"),
            calculated_dq="60 DAYS", expected_loan_status_raw="DEFERRED",
            deferral_type="full", deferral_phase="start", srf_status="ACTIVE", soup_status="ACTIVE",
        )
        recs, candidates, _ = _recommend([entry])
        self.assertEqual(len(recs), 2)  # DQ recommendation + Loan Status recommendation
        outcome = _generate(file_bytes, candidates, [_approve(r) for r in recs])
        self.assertEqual(outcome["applied_count"], 2)
        _, ws = _reload(outcome["file_bytes"])
        for r in excel_rows:
            self.assertEqual(_read_cell(ws, r, "Delinquency Status"), "60 DAYS")
            self.assertEqual(_read_cell(ws, r, "Loan Status"), "DEFERRED")


class ReviewerSubstitutionTests(unittest.TestCase):
    """A reviewer may substitute a value different from STAT-01's calculated
    one — as long as it is within the field's domain. The written cell and the
    log reflect the substituted value; Calculated Value still records what the
    engine computed."""

    def test_substituted_value_is_written_and_all_three_values_recorded(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        recs, candidates, _ = _recommend([_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")])
        # reviewer overrides 60 DAYS -> 90 DAYS (both valid ladder values)
        outcome = _generate(file_bytes, candidates, [_approve(recs[0], approved_value="90 DAYS")])
        self.assertEqual(outcome["applied_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[0], "Delinquency Status"), "90 DAYS")
        _, row, _ = _log_row(outcome["file_bytes"])
        self.assertEqual(row["Original Value"], "CURRENT")     # reported
        self.assertEqual(row["Calculated Value"], "60 DAYS")   # what STAT-01 computed
        self.assertEqual(row["Reviewer Approved Value"], "90 DAYS")  # what the reviewer authorised
        self.assertEqual(row["Corrected Value"], "90 DAYS")    # what was written


class RejectAndReturnLeaveCellUntouchedTests(unittest.TestCase):
    """Rejected and returned items never write to the workbook, and are logged
    with Rows Updated = 0 and a blank Corrected Value."""

    def test_reject_leaves_cell_untouched(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        recs, candidates, _ = _recommend([_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")])
        outcome = _generate(file_bytes, candidates, [{"id": recs[0]["id"], "decision": "reject"}])
        self.assertEqual(outcome["applied_count"], 0)
        self.assertEqual(outcome["rejected_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[0], "Delinquency Status"), "CURRENT")  # untouched
        _, row, _ = _log_row(outcome["file_bytes"])
        self.assertEqual(row["Rows Updated"], 0)
        self.assertEqual(row["Decision"], "Rejected")
        self.assertIn(row["Corrected Value"], (None, ""))  # nothing written -> blank cell reads as None
        self.assertEqual(row["Calculated Value"], "60 DAYS")

    def test_return_leaves_cell_untouched(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        recs, candidates, _ = _recommend([_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")])
        outcome = _generate(file_bytes, candidates,
                            [{"id": recs[0]["id"], "decision": "return", "notify_servicer": True}])
        self.assertEqual(outcome["applied_count"], 0)
        self.assertEqual(outcome["returned_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[0], "Delinquency Status"), "CURRENT")  # untouched
        _, row, _ = _log_row(outcome["file_bytes"])
        self.assertEqual(row["Rows Updated"], 0)
        self.assertEqual(row["Decision"], "Returned to servicer")
        self.assertEqual(row["Notify Servicer"], "Yes")


class AllThreeDecisionsInLogTests(unittest.TestCase):
    """Approve + reject + return across three loans produce three log rows
    with the right Rows Updated (n / 0 / 0), all attributed to the reviewer."""

    def test_approve_reject_return_all_logged(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "2", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "3", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        cd = [
            _correction_data_entry("1", [excel_rows[0]], calculated_dq="60 DAYS"),
            _correction_data_entry("2", [excel_rows[1]], calculated_dq="90 DAYS"),
            _correction_data_entry("3", [excel_rows[2]], calculated_dq="30 DAYS"),
        ]
        recs, candidates, _ = _recommend(cd)
        by_gp = {r["gp"]: r for r in recs}
        decisions = [
            _approve(by_gp["1"]),
            {"id": by_gp["2"]["id"], "decision": "reject"},
            {"id": by_gp["3"]["id"], "decision": "return"},
        ]
        outcome = _generate(file_bytes, candidates, decisions)
        self.assertEqual((outcome["applied_count"], outcome["rejected_count"], outcome["returned_count"]), (1, 1, 1))

        wb, _ = _reload(outcome["file_bytes"])
        log_ws = wb[corrections.CORRECTION_LOG_SHEET_NAME]
        self.assertEqual(log_ws.max_row, 4)  # header + 3 decisions
        headers = [c.value for c in log_ws[1]]
        rows = {
            log_ws.cell(row=r, column=headers.index("SBA GP Number") + 1).value: {
                h: log_ws.cell(row=r, column=headers.index(h) + 1).value
                for h in ("Rows Updated", "Decision")
            }
            for r in range(2, 5)
        }
        self.assertEqual(rows["1"], {"Rows Updated": 1, "Decision": "Approved and applied"})
        self.assertEqual(rows["2"], {"Rows Updated": 0, "Decision": "Rejected"})
        self.assertEqual(rows["3"], {"Rows Updated": 0, "Decision": "Returned to servicer"})
        # every row carries the reviewer identity
        for r in range(2, 5):
            self.assertEqual(
                log_ws.cell(row=r, column=headers.index("Reviewed By (self-declared)") + 1).value,
                "K. Dutta (unverified)",
            )


class ApprovedValueDomainRestrictionTests(unittest.TestCase):
    """A substituted approved value must come from the target field's own
    domain. A cross-domain or terminal substitution is refused — never
    written — and the cell is left untouched."""

    def test_cross_domain_delinquency_substitution_is_refused(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        recs, candidates, _ = _recommend([_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")])
        # FORCLOSURE is not on the delinquency ladder — a substitution cannot
        # reintroduce the exact cross-domain defect the review exists to catch.
        outcome = _generate(file_bytes, candidates, [_approve(recs[0], approved_value="FORCLOSURE")])
        self.assertEqual(outcome["applied_count"], 0)
        self.assertEqual(len(outcome["corrections_not_applied"]), 1)
        self.assertIn("not a valid Delinquency Status", outcome["corrections_not_applied"][0]["reason"])
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[0], "Delinquency Status"), "CURRENT")  # untouched

    def test_reviewer_may_now_substitute_a_terminal_loan_status(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        entry = _correction_data_entry(
            "1", excel_rows, reasons=("LOAN STATUS MISMATCH",),
            expected_loan_status_raw="DEFERRED", deferral_type="full", deferral_phase="start",
            srf_status="ACTIVE", soup_status="ACTIVE",
        )
        # SUPERSEDED 2026-08-08: LIQUIDATION used to be refused outright. A
        # reviewer who examines the evidence and concludes the servicer's
        # terminal report is correct must be able to say so, so an explicitly
        # SELECTED terminal status is now written. What the APPLICATION may
        # generate is unchanged and still excludes it.
        recs, candidates, _ = _recommend([entry])
        outcome = _generate(file_bytes, candidates, [_approve(recs[0], approved_value="LIQUIDATION")])
        self.assertEqual(outcome["applied_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[0], "Loan Status"), "LIQUIDATION")

    def test_validate_approved_value_directly(self):
        ok, norm = corrections.validate_approved_value("Delinquency Status", "90 days")
        self.assertTrue(ok)
        self.assertEqual(norm, "90 DAYS")
        self.assertFalse(corrections.validate_approved_value("Delinquency Status", "FORCLOSURE")[0])
        # SUPERSEDED 2026-08-08: PAID-IN-FULL is now a legitimate reviewer
        # selection. Cross-domain protection is unchanged.
        self.assertTrue(corrections.validate_approved_value("Loan Status", "PAID-IN-FULL")[0])
        self.assertFalse(corrections.validate_approved_value("Loan Status", "CURRENT")[0])  # cross-domain
        self.assertTrue(corrections.validate_approved_value("Loan Status", "DEFERRED")[0])


class CrossDomainRecommendationIsFlaggedButReviewableTests(unittest.TestCase):
    """A cross-domain reported value still produces a recommendation (so the
    defect is visible), carries a reviewFlag, and is never auto-applied — the
    reviewer can approve the calculated ladder value, which IS writable."""

    def test_flagged_recommendation_can_be_approved_to_calculated_value(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "4017399101", "Loan Status": "ACTIVE", "Delinquency Status": "FORCLOSURE",
             "Beginning Balance": 584023.61, "Ending Balance": 584023.61, "Payment Received": 9876.22,
             "Scheduled Payment": 4697.01},
        ])
        entry = _correction_data_entry(
            "4017399101", excel_rows, calculated_dq="30 DAYS", srf_dq="FORCLOSURE",
        )
        recs, candidates, _ = _recommend([entry])
        self.assertEqual(len(recs), 1)
        self.assertIsNotNone(recs[0]["reviewFlag"])       # both faults surfaced
        self.assertEqual(recs[0]["reported"], "FORCLOSURE")
        self.assertEqual(recs[0]["calculated"], "30 DAYS")
        # approving defaults to the calculated ladder value, which is writable
        outcome = _generate(file_bytes, candidates, [_approve(recs[0], notify=True)])
        self.assertEqual(outcome["applied_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[0], "Delinquency Status"), "30 DAYS")
        _, row, _ = _log_row(outcome["file_bytes"])
        self.assertEqual(row["Notify Servicer"], "Yes")
        self.assertTrue(row["Review Reason"])  # the cross-domain flag is recorded


class BuildEligibleCorrectionsSignatureTests(unittest.TestCase):
    """build_eligible_corrections has always returned a single list — that
    contract (a public function other suites may call) must not change."""

    def test_returns_a_plain_list_not_a_tuple(self):
        entry = _correction_data_entry("1", [10], calculated_dq="60 DAYS")
        result = corrections.build_eligible_corrections([entry])
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["gp"], "1")

    def test_returns_a_plain_list_even_when_every_row_reference_is_invalid(self):
        entry = _correction_data_entry("1", [None, "", "abc"], calculated_dq="60 DAYS")
        result = corrections.build_eligible_corrections([entry])
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["excel_rows"], [])

    def test_empty_correction_data_returns_empty_list(self):
        self.assertEqual(corrections.build_eligible_corrections([]), [])


if __name__ == "__main__":
    unittest.main()


class UnclassifiedDeferralTypeTests(unittest.TestCase):
    """A deferral whose memo text never states full-P&I vs principal-only.

    classify_deferral_type returns "unclassified" for it rather than assuming
    "full". Two separate protections must both hold, so that dropping either
    one still leaves the loan safe:

      1. STAT-01 raises DEFERRAL DATA QUALITY WARNING, which is in
         WARNING_REASONS and excludes the loan from BOTH correction types.
      2. The Loan Status -> DEFERRED gate is a strict `== "full"`, so an
         unclassified deferral cannot qualify even with no warning present.
    """

    def _entry(self, **overrides):
        base = dict(
            reasons=["LOAN STATUS MISMATCH"], dq_match=True, calculated_dq="CURRENT",
            status_match=False, expected_loan_status_raw="DEFERRED",
            srf_status="ACTIVE", soup_status="ACTIVE", deferral_phase="start",
        )
        base.update(overrides)
        return _correction_data_entry("1111111111", [4], **base)

    def test_full_deferral_still_produces_the_loan_status_candidate(self):
        # Control: the same shape with a confirmed full P&I deferral IS
        # eligible — otherwise the assertions below would pass vacuously.
        recs, _, _ = _recommend([self._entry(deferral_type="full")])
        self.assertEqual([r["field"] for r in recs], ["Loan Status"])
        self.assertEqual(recs[0]["calculated"], "DEFERRED")

    def test_unclassified_is_not_eligible_even_without_a_warning(self):
        # Protection 2 in isolation: no warning reason present at all.
        recs, _, _ = _recommend([self._entry(deferral_type="unclassified")])
        self.assertEqual(recs, [])

    def test_warning_alone_also_excludes_it(self):
        # Protection 1 in isolation: the warning excludes the loan from every
        # correction type, not just the deferral one.
        recs, _, _ = _recommend([self._entry(
            deferral_type="unclassified",
            reasons=["LOAN STATUS MISMATCH", "DEFERRAL DATA QUALITY WARNING"],
        )])
        self.assertEqual(recs, [])

    def test_warning_also_blocks_an_unrelated_delinquency_correction(self):
        # The exclusion is loan-level by design (corrections.py:346): a loan
        # with flagged uncertainty about itself isn't trusted for a different
        # field either.
        recs, _, _ = _recommend([_correction_data_entry(
            "2222222222", [5],
            reasons=["DELINQUENCY STATUS MISMATCH", "DEFERRAL DATA QUALITY WARNING"],
        )])
        self.assertEqual(recs, [])


class UnclassifiedDeferralArithmeticTests(unittest.TestCase):
    """Option B's core guarantee: "unclassified" changes what a reviewer SEES,
    never what the pipeline COMPUTES.

    An unclassified deferral rides the same full-P&I calculation path it rode
    when it was silently classified "full", so no delinquency value and no
    Portfolio Summary bucket moves. Switching it to the no-deferral path would
    skip the client-confirmed resume reset and assert delinquency on evidence
    just declared unknown — measured at $280,162.64 moving Current -> 60 Days
    on one real loan.
    """

    def _memo(self, servicer_action):
        df = pd.DataFrame([{
            "Lender": "Millennium Bank", "SBA Loan ID": "5658839107",
            "Lender Loan ID": "5658839107", "Borrower Name": "Test",
            "Securitization Action": "Deferred", "Servicer Action": servicer_action,
            "Deferral Start Date": pd.Timestamp("2026-03-01"),
            "Deferral End Date": pd.Timestamp("2026-05-01"),
            "Regular Payment Resume Date": pd.Timestamp("2026-06-01"),
            "Report Date": pd.Timestamp("2026-05-26"),
        }])
        df["gp_sba"] = df["SBA Loan ID"]
        df["gp_lender"] = df["Lender Loan ID"]
        df["deferral_type"] = df["Servicer Action"].apply(file_readers.classify_deferral_type)
        return df

    def test_unclassified_and_full_agree_on_every_phase(self):
        # The real loan's text ("three months deferment") against an explicit
        # full-P&I equivalent, across the whole deferral window.
        for cycle in ("2026-03-01", "2026-04-01", "2026-05-01", "2026-06-01"):
            with self.subTest(cycle=cycle):
                unclassified = checks._lookup_deferral(
                    self._memo("three months deferment"), "5658839107",
                    "Millennium Bank", pd.Timestamp(cycle), "DEFERRED")
                full = checks._lookup_deferral(
                    self._memo("three months principal and interest deferment"), "5658839107",
                    "Millennium Bank", pd.Timestamp(cycle), "DEFERRED")
                self.assertEqual(unclassified[0]["phase"], full[0]["phase"])
                self.assertEqual(unclassified[0]["months_since_start"],
                                 full[0]["months_since_start"])

    def test_resume_reset_still_fires_for_an_unclassified_deferral(self):
        # The client-confirmed reset is keyed on the resume cycle; losing it
        # here is exactly what moved $280k in the rejected variant.
        info, _, _, _ = checks._lookup_deferral(
            self._memo("three months deferment"), "5658839107",
            "Millennium Bank", pd.Timestamp("2026-06-01"), "DEFERRED")
        self.assertEqual(info["phase"], "resume")

    def test_unclassified_raises_the_data_quality_warning(self):
        info, warning, warning_type, evidence = checks._lookup_deferral(
            self._memo("three months deferment"), "5658839107",
            "Millennium Bank", pd.Timestamp("2026-04-01"), "DEFERRED")
        self.assertEqual(info["deferral_type"], "unclassified")
        self.assertEqual(warning_type, "deferral_dates")
        self.assertIn("does not state whether", warning)

    def test_a_typed_deferral_raises_no_such_warning(self):
        for text in ("three months principal and interest deferment",
                     "three months principal only deferment"):
            with self.subTest(text=text):
                _, warning, warning_type, _ = checks._lookup_deferral(
                    self._memo(text), "5658839107", "Millennium Bank",
                    pd.Timestamp("2026-04-01"), "DEFERRED")
                self.assertIsNone(warning)
                self.assertIsNone(warning_type)


class TerminalStatusRoutesToReviewTests(unittest.TestCase):
    """A servicer-reported terminal status routes corrections TO interactive
    review; it never excludes them FROM it (2026-08-08).

    The distinction the rules now draw: GENERATING a recommendation is
    permitted when independent evidence contradicts the servicer's claim;
    AUTOMATICALLY APPLYING one remains impossible. Every correction here is
    review_required and writes only on explicit approval.
    """

    def _entry(self, **over):
        base = dict(
            reasons=["LOAN STATUS MISMATCH", "DELINQUENCY STATUS MISMATCH", "TERMINAL STATUS REVIEW"],
            dq_match=False, calculated_dq="90 DAYS", srf_dq="120 DAYS",
            status_match=False, srf_status="LIQUIDATION", soup_status="PAID-IN-FULL",
            expected_loan_status_raw=None, expected_loan_status_effective="PAID-IN-FULL",
        )
        base.update(over)
        return _correction_data_entry("4338609104", [4], **base)

    def test_terminal_status_does_not_suppress_the_dq_candidate(self):
        recs, _, _ = _recommend([self._entry()])
        dq = [r for r in recs if r["field"] == "Delinquency Status"]
        self.assertEqual(len(dq), 1)
        self.assertEqual(dq[0]["calculated"], "90 DAYS")

    def test_dq_candidate_is_flagged_for_review_not_auto_applied(self):
        recs, _, _ = _recommend([self._entry()])
        dq = next(r for r in recs if r["field"] == "Delinquency Status")
        self.assertIsNotNone(dq["reviewFlag"])
        self.assertIn("terminal collection status", dq["reviewFlag"])

    def test_loan_status_candidate_uses_the_audits_effective_expected_value(self):
        recs, _, _ = _recommend([self._entry()])
        ls = next(r for r in recs if r["field"] == "Loan Status")
        self.assertEqual(ls["calculated"], "PAID-IN-FULL")
        self.assertIsNotNone(ls["reviewFlag"])

    def test_both_correction_types_appear_for_the_same_loan(self):
        recs, _, _ = _recommend([self._entry()])
        self.assertEqual({r["field"] for r in recs}, {"Loan Status", "Delinquency Status"})
        self.assertEqual(len({r["id"] for r in recs}), 2)

    def test_proposed_value_is_never_the_servicers_own_terminal_status(self):
        # The whole point: a terminal status must not validate itself.
        recs, _, _ = _recommend([self._entry()])
        for r in recs:
            self.assertNotIn("LIQUIDATION", str(r["calculated"]).upper())

    def test_stale_soup_is_not_proposed_when_the_audit_derived_something_else(self):
        # SOUP still says DEFERRED, but the memo/resume logic established
        # ACTIVE. The candidate must follow the audit's effective value, not
        # copy prior SOUP -- this is why the branch reads
        # expected_loan_status_effective rather than soup_status.
        recs, _, _ = _recommend([self._entry(
            soup_status="DEFERRED", expected_loan_status_effective="ACTIVE")])
        ls = next(r for r in recs if r["field"] == "Loan Status")
        self.assertEqual(ls["calculated"], "ACTIVE")

    def test_unverified_expected_value_produces_no_loan_status_candidate(self):
        # Insufficient or contradictory independent evidence -> no proposal.
        recs, _, _ = _recommend([self._entry(expected_loan_status_effective="UNVERIFIED")])
        self.assertEqual([r["field"] for r in recs], ["Delinquency Status"])

    def test_collection_status_is_never_proposed_even_if_it_is_the_expected_value(self):
        recs, _, _ = _recommend([self._entry(
            srf_status="ACTIVE", expected_loan_status_effective="LIQUIDATION",
            reasons=["LOAN STATUS MISMATCH", "TERMINAL STATUS REVIEW"], dq_match=True)])
        self.assertEqual([r["field"] for r in recs], [])

    def test_no_candidate_when_expected_matches_what_the_servicer_reported(self):
        recs, _, _ = _recommend([self._entry(expected_loan_status_effective="LIQUIDATION")])
        self.assertNotIn("Loan Status", [r["field"] for r in recs])


class TerminalWriteRequiresExplicitApprovalTests(unittest.TestCase):
    """The write guard must FAIL CLOSED: a terminal Loan Status is writable
    only when the correction carries explicit reviewer approval, and that flag
    is set in exactly one place — after the decision is confirmed to be an
    approval and the value has passed validation."""

    def _corr(self, value="PAID-IN-FULL", **over):
        c = {"gp": "1", "field": "Loan Status", "corrected_value": value,
             "excel_rows": [4], "sheet_name": "Servicer Inputs",
             "header_excel_row": HEADER_ROW, "skipped_row_refs": [],
             "eligibility_rule": "test fixture", "reported_value": "LIQUIDATION"}
        c.update(over)
        return c

    def _wb(self):
        file_bytes, rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "LIQUIDATION", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        return load_workbook(io.BytesIO(file_bytes), data_only=False), rows

    def test_paid_in_full_refused_without_the_approval_flag(self):
        wb, _ = self._wb()
        ok, detail = corrections._apply_one_correction(wb, self._corr())
        self.assertFalse(ok)
        self.assertIn("explicit reviewer approval", detail)

    def test_paid_in_full_written_with_the_approval_flag(self):
        wb, rows = self._wb()
        ok, _ = corrections._apply_one_correction(wb, self._corr(reviewer_approved=True))
        self.assertTrue(ok)
        self.assertEqual(wb["Servicer Inputs"].cell(row=rows[0], column=2).value, "PAID-IN-FULL")

    def test_collection_status_written_when_explicitly_approved(self):
        # SUPERSEDED 2026-08-08: these were refused unconditionally. The
        # application still never GENERATES them; a reviewer may select them.
        for value in ("LIQUIDATION", "FORECLOSURE", "REO"):
            with self.subTest(value=value):
                wb, rows = self._wb()
                ok, _ = corrections._apply_one_correction(
                    wb, self._corr(value=value, reviewer_approved=True))
                self.assertTrue(ok)
                self.assertEqual(wb["Servicer Inputs"].cell(row=rows[0], column=2).value, value)

    def test_every_loan_status_write_needs_the_approval_flag(self):
        # WIDENED 2026-08-08: the gate covers ALL Loan Status writes now, not
        # only terminal ones -- it is the whole of the "never automatically
        # applied" guarantee, so it can no longer be value-dependent.
        for value in ("ACTIVE", "DEFERRED", "LIQUIDATION", "PAID-IN-FULL"):
            with self.subTest(value=value):
                wb, _ = self._wb()
                ok, detail = corrections._apply_one_correction(wb, self._corr(value=value))
                self.assertFalse(ok)
                self.assertIn("explicit reviewer approval", detail)


class TerminalApprovedValueDomainTests(unittest.TestCase):
    """PAID-IN-FULL may be APPROVED from a generated recommendation, but never
    typed in as a freeform substitution."""

    def test_paid_in_full_accepted_when_not_substituted(self):
        ok, value = corrections.validate_approved_value("Loan Status", "PAID-IN-FULL", substituted=False)
        self.assertTrue(ok)
        self.assertEqual(value, "PAID-IN-FULL")

    def test_paid_in_full_permitted_as_a_reviewer_selection(self):
        # SUPERSEDED 2026-08-08.
        ok, value = corrections.validate_approved_value("Loan Status", "PAID-IN-FULL", substituted=True)
        self.assertTrue(ok)
        self.assertEqual(value, "PAID-IN-FULL")

    def test_soup_spelling_is_canonicalised_to_the_workbook_representation(self):
        # SOUP carries "Paid-in-Full" alongside "PAID-IN-FULL"; the servicer
        # workbooks are uniformly "PAID-IN-FULL".
        for spelling in ("Paid-in-Full", "paid in full", "PAID-IN-FULL"):
            with self.subTest(spelling=spelling):
                ok, value = corrections.validate_approved_value("Loan Status", spelling, substituted=False)
                self.assertTrue(ok)
                self.assertEqual(value, "PAID-IN-FULL")

    def test_collection_status_selectable_by_reviewer_but_not_generatable(self):
        # The rule the 2026-08-08 clarification draws: the APPLICATION may not
        # generate these (substituted=False is the app-derived route), but a
        # REVIEWER may explicitly select them.
        for value in ("LIQUIDATION", "FORECLOSURE", "REO"):
            with self.subTest(value=value):
                ok_gen, _ = corrections.validate_approved_value("Loan Status", value, substituted=False)
                self.assertFalse(ok_gen)
                ok_sel, out = corrections.validate_approved_value("Loan Status", value, substituted=True)
                self.assertTrue(ok_sel)
                self.assertEqual(out, value)

    def test_active_and_deferred_still_work_on_both_routes(self):
        for sub in (True, False):
            for value in ("ACTIVE", "DEFERRED"):
                with self.subTest(substituted=sub, value=value):
                    ok, out = corrections.validate_approved_value("Loan Status", value, substituted=sub)
                    self.assertTrue(ok)
                    self.assertEqual(out, value)

    def test_cross_domain_value_still_refused_when_not_substituted(self):
        ok, _ = corrections.validate_approved_value("Loan Status", "60 DAYS", substituted=False)
        self.assertFalse(ok)


class AuthoritativeLoanStatusDomainTests(unittest.TestCase):
    """The reviewer's Loan Status domain is the client's own, taken from the
    SRF template's "Reference Tables" sheet (identical across every servicer
    workbook inspected). Loan Status review works like Delinquency Status
    review: the reviewer picks from the real field domain.

    The rule this encodes (2026-08-08):
      the APPLICATION must never independently generate or automatically apply
      a terminal collection status;
      a HUMAN REVIEWER may explicitly select and approve any legitimate Loan
      Status -- but nothing is written without explicit approval.
    """

    AUTHORITATIVE = [
        "ACTIVE", "DEFERRED", "PAID-IN-FULL", "REO", "FORECLOSURE",
        "BANKRUPTCY", "SUBSTITUTION-IN", "SUBSTITUTION-OUT", "LIQUIDATION",
    ]
    # What the application itself may ever produce. Deliberately narrower.
    GENERATABLE = {"ACTIVE", "DEFERRED", "PAID-IN-FULL"}

    def _loan_status_candidate(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        entry = _correction_data_entry(
            "1", excel_rows, reasons=("LOAN STATUS MISMATCH",),
            expected_loan_status_raw="DEFERRED", deferral_type="full", deferral_phase="start",
            srf_status="ACTIVE", soup_status="ACTIVE",
        )
        recs, candidates, _ = _recommend([entry])
        return file_bytes, excel_rows, recs, candidates

    def test_backend_domain_matches_the_template_reference_list(self):
        self.assertEqual(corrections._PERMITTED_LOAN_STATUS, set(self.AUTHORITATIVE))

    def test_charge_off_is_not_in_the_domain(self):
        # Not in the template's Loan Status list -- it lives in the Servicer
        # Memo's Securitization Action field, a different field entirely.
        for spelling in ("CHARGE-OFF", "CHARGE OFF", "CHARGEOFF"):
            with self.subTest(spelling=spelling):
                self.assertFalse(corrections.validate_approved_value("Loan Status", spelling)[0])

    def test_all_nine_are_accepted_as_reviewer_selections(self):
        for value in self.AUTHORITATIVE:
            with self.subTest(value=value):
                ok, out = corrections.validate_approved_value("Loan Status", value, substituted=True)
                self.assertTrue(ok, f"{value} should be selectable")
                self.assertEqual(out, value)

    def test_all_nine_are_written_end_to_end_when_approved(self):
        for value in self.AUTHORITATIVE:
            with self.subTest(value=value):
                file_bytes, excel_rows, recs, candidates = self._loan_status_candidate()
                outcome = _generate(file_bytes, candidates, [_approve(recs[0], approved_value=value)])
                self.assertEqual(outcome["applied_count"], 1, outcome["corrections_not_applied"])
                _, ws = _reload(outcome["file_bytes"])
                self.assertEqual(_read_cell(ws, excel_rows[0], "Loan Status"), value)

    def test_only_the_narrow_domain_may_be_application_generated(self):
        for value in self.AUTHORITATIVE:
            with self.subTest(value=value):
                ok, _ = corrections.validate_approved_value("Loan Status", value, substituted=False)
                self.assertEqual(ok, value in self.GENERATABLE)

    def test_reviewer_selections_are_canonicalised(self):
        for typed, canonical in (("paid in full", "PAID-IN-FULL"),
                                 ("liquidation", "LIQUIDATION"),
                                 ("substitution out", "SUBSTITUTION-OUT"),
                                 ("Active", "ACTIVE")):
            with self.subTest(typed=typed):
                ok, out = corrections.validate_approved_value("Loan Status", typed, substituted=True)
                self.assertTrue(ok)
                self.assertEqual(out, canonical)

    def test_values_outside_the_domain_are_still_refused(self):
        for value in ("PAID IN FUL", "ACTIVE LOAN", "REOWNED", "", "UNVERIFIED"):
            with self.subTest(value=value):
                self.assertFalse(corrections.validate_approved_value("Loan Status", value)[0])

    def test_cross_domain_delinquency_values_are_still_refused(self):
        for value in ("CURRENT", "60 DAYS", "150+ DAYS", "DELINQUENT"):
            with self.subTest(value=value):
                ok, reason = corrections.validate_approved_value("Loan Status", value, substituted=True)
                self.assertFalse(ok)
                self.assertIn("Delinquency Status", reason)

    def test_no_reviewer_selection_is_written_without_approval(self):
        # The fail-closed guarantee, across the WHOLE domain: the write gate is
        # the entirety of "never automatically applied", so it must hold for
        # every value, not just terminal ones.
        for value in self.AUTHORITATIVE:
            with self.subTest(value=value):
                file_bytes, excel_rows, _, _ = self._loan_status_candidate()
                wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
                corr = {"gp": "1", "field": "Loan Status", "corrected_value": value,
                        "excel_rows": excel_rows, "sheet_name": "Servicer Inputs",
                        "header_excel_row": HEADER_ROW, "skipped_row_refs": [],
                        "eligibility_rule": "test fixture", "reported_value": "ACTIVE"}
                ok, detail = corrections._apply_one_correction(wb, corr)
                self.assertFalse(ok)
                self.assertIn("explicit reviewer approval", detail)
                self.assertEqual(wb["Servicer Inputs"].cell(row=excel_rows[0], column=2).value, "ACTIVE")

    def test_frontend_dropdown_matches_the_backend_domain(self):
        # The frontend comment promises these are kept in sync; this is what
        # makes that promise checkable rather than aspirational.
        import re as _re
        path = os.path.join(os.path.dirname(__file__), "..", "frontend", "src",
                            "components", "CorrectionReview.jsx")
        if not os.path.exists(path):
            self.skipTest("frontend not present")
        src = open(path).read()
        block = src[src.index("const LOAN_STATUS_DOMAIN"):]
        block = block[:block.index("]") + 1]
        self.assertEqual(set(_re.findall(r"'([^']+)'", block)),
                         corrections._PERMITTED_LOAN_STATUS)


class CollectionStatusReviewRoutingTests(unittest.TestCase):
    """FORECLOSURE, LIQUIDATION and REO must behave IDENTICALLY.

    Two notions of "terminal" live in corrections.py and overlap only on
    LIQUIDATION. Review routing follows the COLLECTION set
    (FORECLOSURE/LIQUIDATION/REO, matching checks._TERMINAL_COLLECTION_KEYWORDS
    so the review flag and the TERMINAL STATUS REVIEW reason can never
    disagree). The deferral safety guards follow a different set that includes
    PAID-IN-FULL, because "may we auto-derive DEFERRED for this loan?" is a
    different question from "is this loan in collection?".

    Keying review routing on the wrong set silently excluded FORECLOSURE and
    REO -- latent only because the delivered data contains LIQUIDATION alone.
    """

    COLLECTION = ["FORECLOSURE", "LIQUIDATION", "REO"]

    def _candidates_for(self, srf_status, soup_status="ACTIVE", expected="ACTIVE"):
        entry = _correction_data_entry(
            "1", [4],
            reasons=["LOAN STATUS MISMATCH", "DELINQUENCY STATUS MISMATCH", "TERMINAL STATUS REVIEW"],
            dq_match=False, calculated_dq="90 DAYS", srf_dq="120 DAYS", status_match=False,
            srf_status=srf_status, soup_status=soup_status,
            expected_loan_status_raw=None, expected_loan_status_effective=expected,
        )
        recs, _, _ = _recommend([entry])
        return {r["field"]: r for r in recs}

    def test_all_three_collection_statuses_generate_a_loan_status_candidate(self):
        for status in self.COLLECTION:
            with self.subTest(status=status):
                recs = self._candidates_for(status)
                self.assertIn("Loan Status", recs, f"{status} generated no Loan Status candidate")
                self.assertEqual(recs["Loan Status"]["calculated"], "ACTIVE")

    def test_all_three_collection_statuses_flag_the_loan_status_candidate(self):
        for status in self.COLLECTION:
            with self.subTest(status=status):
                flag = self._candidates_for(status)["Loan Status"]["reviewFlag"]
                self.assertIsNotNone(flag)
                self.assertIn("terminal collection status", flag)

    def test_all_three_collection_statuses_flag_the_dq_candidate(self):
        for status in self.COLLECTION:
            with self.subTest(status=status):
                flag = self._candidates_for(status)["Delinquency Status"]["reviewFlag"]
                self.assertIsNotNone(flag, f"{status} DQ candidate carried no review flag")
                self.assertIn("terminal collection status", flag)

    def test_the_three_produce_identical_shapes(self):
        shapes = []
        for status in self.COLLECTION:
            recs = self._candidates_for(status)
            shapes.append(tuple(sorted(
                (f, r["calculated"], bool(r["reviewFlag"])) for f, r in recs.items())))
        self.assertEqual(len(set(shapes)), 1, f"collection statuses diverged: {shapes}")

    def test_paid_in_full_does_not_trigger_the_collection_review_path(self):
        # PAID-IN-FULL is a finished loan, not an active collection proceeding.
        # It must not generate the terminal-review Loan Status candidate, and
        # its DQ candidate must not carry the terminal collection flag.
        recs = self._candidates_for("PAID-IN-FULL")
        self.assertNotIn("Loan Status", recs)
        dq_flag = recs["Delinquency Status"]["reviewFlag"]
        self.assertIsNone(dq_flag)

    def test_bankruptcy_does_not_trigger_the_collection_review_path(self):
        recs = self._candidates_for("BANKRUPTCY")
        self.assertNotIn("Loan Status", recs)
        self.assertIsNone(recs["Delinquency Status"]["reviewFlag"])

    def test_collection_set_matches_the_checks_module_definition(self):
        # If these ever drift, a loan could carry TERMINAL STATUS REVIEW while
        # its recommendation shows no flag -- the exact defect this fixes.
        self.assertEqual(set(corrections._COLLECTION_STATUS_KEYWORDS),
                         set(checks._TERMINAL_COLLECTION_KEYWORDS))

    def test_the_two_terminal_sets_remain_distinct(self):
        # Documents that the overlap is LIQUIDATION only, and that PAID-IN-FULL
        # is deliberately in the deferral-guard set and not the collection set.
        self.assertEqual(
            set(corrections._COLLECTION_STATUS_KEYWORDS) & set(corrections._LOAN_STATUS_TERMINAL_KEYWORDS),
            {"LIQUIDATION"})
        self.assertIn("PAID-IN-FULL", corrections._LOAN_STATUS_TERMINAL_KEYWORDS)
        self.assertNotIn("PAID-IN-FULL", corrections._COLLECTION_STATUS_KEYWORDS)

    def test_deferral_guard_still_blocks_auto_derived_deferred_for_all_of_them(self):
        # The separate question: the application must not AUTO-DERIVE
        # "-> DEFERRED" for a liquidating or paid-off loan. Unchanged.
        for status in ("LIQUIDATION", "PAID-IN-FULL"):
            with self.subTest(status=status):
                entry = _correction_data_entry(
                    "1", [4], reasons=("LOAN STATUS MISMATCH",),
                    expected_loan_status_raw="DEFERRED", deferral_type="full",
                    deferral_phase="start", srf_status=status, soup_status="ACTIVE",
                )
                recs, _, _ = _recommend([entry])
                self.assertNotIn("DEFERRED", [r["calculated"] for r in recs])

    def test_reviewer_may_still_select_any_of_the_nine_for_a_collection_loan(self):
        # The re-keying must not narrow the reviewer's domain.
        for value in sorted(corrections._PERMITTED_LOAN_STATUS):
            with self.subTest(value=value):
                self.assertTrue(
                    corrections.validate_approved_value("Loan Status", value, substituted=True)[0])
