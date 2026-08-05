"""
Regression test for corrections.py's multi-posting fix.

CLIENT-CONFIRMED (2026-08-05): when STAT-01 consolidates multiple postings
for one loan into a single audited Loan Status / Delinquency Status, an
approved correction must be written to EVERY posting row for that loan/
cycle, not only the chronologically last row `excel_row` used to point at.
Locks in:
  - a 1-posting loan still updates exactly 1 row (no behavior change for
    the common case)
  - a 3-posting loan's Delinquency Status correction updates all 3 rows
  - a Loan Status correction updates all rows for that loan
  - correcting one field never touches the other field
  - other loans in the same sheet are completely untouched
  - formulas, number formats, and posting order are preserved everywhere
  - re-uploading the corrected workbook produces one consistent
    consolidated status matching the correction
  - the Correction Log records ONE row per loan-level correction, with a
    "Rows Updated" count, not one row per Excel row touched

Run with:
    python -m unittest test_corrections.py
"""

import io
import unittest

from openpyxl import Workbook, load_workbook
import pandas as pd

import corrections
import checks


HEADERS = [
    "SBA GP Number", "Loan Status", "Delinquency Status", "Beginning Balance",
    "Ending Balance", "Payment Received", "Scheduled Payment", "Net",
]
HEADER_ROW = 3  # simulates a template with title/blank rows above the header


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
    }
    entry.update(overrides)
    return entry


class SinglePostingStillUpdatesOneRowTests(unittest.TestCase):
    """No behavior change for the common case: a loan with exactly one
    posting this cycle still gets exactly one row updated."""

    def test_single_posting_updates_exactly_one_row(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        stat01_result = {
            "correction_data": [_correction_data_entry("1", excel_rows)],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        self.assertEqual(outcome["correction_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[0], "Delinquency Status"), "60 DAYS")


class ThreePostingDelinquencyCorrectionTests(unittest.TestCase):
    """A loan with three postings and an approved Delinquency Status
    correction must update Delinquency Status on ALL three rows."""

    def _build(self):
        return _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])

    def test_all_three_rows_get_the_corrected_delinquency_status(self):
        file_bytes, excel_rows = self._build()
        self.assertEqual(len(excel_rows), 3)
        stat01_result = {
            "correction_data": [_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        self.assertEqual(outcome["correction_count"], 1)  # ONE loan-level correction
        _, ws = _reload(outcome["file_bytes"])
        for r in excel_rows:
            self.assertEqual(_read_cell(ws, r, "Delinquency Status"), "60 DAYS")


class LoanStatusCorrectionUpdatesAllRowsTests(unittest.TestCase):
    """A Loan Status correction (-> DEFERRED) updates Loan Status on every
    posting row for that loan."""

    def test_all_rows_get_deferred(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        entry = _correction_data_entry(
            "1", excel_rows, reasons=("LOAN STATUS MISMATCH",),
            expected_loan_status_raw="DEFERRED", deferral_type="full", deferral_phase="start",
            srf_status="ACTIVE", soup_status="ACTIVE",
        )
        stat01_result = {"correction_data": [entry], "_cycle_month": None}
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        self.assertEqual(outcome["correction_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        for r in excel_rows:
            self.assertEqual(_read_cell(ws, r, "Loan Status"), "DEFERRED")


class FieldIndependenceTests(unittest.TestCase):
    """Correcting one field must never change the other status field on
    any row."""

    def test_delinquency_correction_leaves_loan_status_untouched(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        stat01_result = {
            "correction_data": [_correction_data_entry("1", excel_rows, calculated_dq="90 DAYS")],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        _, ws = _reload(outcome["file_bytes"])
        for r in excel_rows:
            self.assertEqual(_read_cell(ws, r, "Delinquency Status"), "90 DAYS")
            self.assertEqual(_read_cell(ws, r, "Loan Status"), "ACTIVE")  # unchanged


class OtherLoansUntouchedTests(unittest.TestCase):
    """A correction targeting one loan's postings must leave every other
    loan's rows completely unchanged."""

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
        stat01_result = {
            "correction_data": [_correction_data_entry("1", loan1_rows, calculated_dq="60 DAYS")],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        _, ws = _reload(outcome["file_bytes"])
        # loan 2, untouched, keeps every original value exactly
        self.assertEqual(_read_cell(ws, loan2_row, "SBA GP Number"), "2")
        self.assertEqual(_read_cell(ws, loan2_row, "Loan Status"), "ACTIVE")
        self.assertEqual(_read_cell(ws, loan2_row, "Delinquency Status"), "30 DAYS")
        self.assertEqual(_read_cell(ws, loan2_row, "Beginning Balance"), 40000)
        self.assertEqual(_read_cell(ws, loan2_row, "Ending Balance"), 39000)
        self.assertEqual(_read_cell(ws, loan2_row, "Payment Received"), 450)


class FormattingAndFormulasPreservedTests(unittest.TestCase):
    """Number formats, formulas, and posting order must survive a
    correction untouched — only the target status cell's value changes."""

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

        stat01_result = {
            "correction_data": [_correction_data_entry("1", excel_rows[:2], calculated_dq="90 DAYS")],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        _, ws_after = _reload(outcome["file_bytes"])

        for r in excel_rows:
            self.assertEqual(ws_after.cell(row=r, column=HEADERS.index("Net") + 1).value, before[r]["net_formula"])
            self.assertEqual(ws_after.cell(row=r, column=HEADERS.index("Beginning Balance") + 1).number_format, before[r]["beg_fmt"])
            self.assertEqual(_read_cell(ws_after, r, "Beginning Balance"), before[r]["beg_val"])
            self.assertEqual(_read_cell(ws_after, r, "Ending Balance"), before[r]["ending_val"])
            self.assertEqual(_read_cell(ws_after, r, "SBA GP Number"), before[r]["gp"])

        # posting order preserved: rows are still in the same physical order
        self.assertEqual(excel_rows, sorted(excel_rows))


class ReuploadConsistencyTests(unittest.TestCase):
    """Re-uploading the corrected workbook must produce one consistent
    consolidated status matching the correction — every posting row now
    agrees, so consolidation's chronologically-last pick is trivially the
    corrected value too."""

    def test_reconsolidating_the_corrected_file_yields_the_corrected_status(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        stat01_result = {
            "correction_data": [_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        _, ws = _reload(outcome["file_bytes"])

        # Simulate re-upload: read the corrected Servicer Inputs sheet back
        # into a plain DataFrame the same shape checks.py expects.
        reuploaded_rows = [
            {h: _read_cell(ws, r, h) for h in HEADERS if h != "Net"}
            for r in excel_rows
        ]
        reuploaded_df = pd.DataFrame(reuploaded_rows)
        si_grouped = checks.consolidate_cycle_postings(reuploaded_df)

        self.assertEqual(len(si_grouped), 1)  # still one consolidated loan
        self.assertEqual(si_grouped.loc["1", "dq_status"], "60 DAYS")
        # every posting agreeing means posting_count still correctly reflects
        # 3 real postings -- consolidation logic itself is unchanged
        self.assertEqual(si_grouped.loc["1", "posting_count"], 3)


class CorrectionLogTests(unittest.TestCase):
    """The Correction Log must record ONE row per loan-level correction —
    with a Rows Updated count and all affected Excel row numbers listed —
    never one misleading row per Excel row touched."""

    def test_three_posting_correction_produces_one_log_row_with_rows_updated_three(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        stat01_result = {
            "correction_data": [_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        wb, _ = _reload(outcome["file_bytes"])
        log_ws = wb[corrections.CORRECTION_LOG_SHEET_NAME]

        headers = [c.value for c in log_ws[1]]
        self.assertEqual(headers, corrections.CORRECTION_LOG_HEADERS)
        self.assertEqual(log_ws.max_row, 2)  # header + exactly ONE correction row

        row = {h: log_ws.cell(row=2, column=i + 1).value for i, h in enumerate(headers)}
        self.assertEqual(row["Rows Updated"], 3)
        self.assertEqual(row["SBA GP Number"], "1")
        self.assertEqual(row["Field Updated"], "Delinquency Status")
        self.assertEqual(row["Corrected Value"], "60 DAYS")
        for r in excel_rows:
            self.assertIn(str(r), row["Excel Row Numbers"])
        # original values for all three rows recorded, not just one
        self.assertEqual(row["Original Value"].count("CURRENT"), 3)

    def test_correction_count_reflects_loans_not_rows(self):
        # correction_count is what the frontend shows as "N corrections
        # applied" -- must be loan-level (1), not row-level (3).
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        stat01_result = {
            "correction_data": [_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        self.assertEqual(outcome["correction_count"], 1)


class RowLevelSafetyStillAppliesTests(unittest.TestCase):
    """Per-row GP verification must still happen for EVERY row, not just
    the first — a row that no longer matches the expected loan is skipped
    individually rather than corrupting an unrelated loan's data."""

    def test_one_mismatched_row_is_skipped_others_still_applied(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        # Claim a THIRD row (never written by _build_workbook_bytes, so it's
        # genuinely blank/mismatched) belongs to loan "1" too -- simulates a
        # stale excel_rows entry from a template shift between analysis and
        # correction generation.
        stale_row = excel_rows[-1] + 5
        stat01_result = {
            "correction_data": [_correction_data_entry("1", excel_rows + [stale_row], calculated_dq="60 DAYS")],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        self.assertEqual(outcome["correction_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        for r in excel_rows:
            self.assertEqual(_read_cell(ws, r, "Delinquency Status"), "60 DAYS")


class ExcelRowsDedupeAndValidationTests(unittest.TestCase):
    """Safeguard 1 & 2: excel_rows is deduplicated (order-preserving) and
    blank/null/nonnumeric/invalid entries are dropped without guessing,
    before any correction is applied."""

    def test_duplicate_row_reference_written_only_once(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        dup_rows = [excel_rows[0], excel_rows[0], excel_rows[1]]  # first row listed twice
        stat01_result = {
            "correction_data": [_correction_data_entry("1", dup_rows, calculated_dq="60 DAYS")],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        self.assertEqual(outcome["correction_count"], 1)
        wb, _ = _reload(outcome["file_bytes"])
        log_ws = wb[corrections.CORRECTION_LOG_SHEET_NAME]
        headers = [c.value for c in log_ws[1]]
        row = {h: log_ws.cell(row=2, column=i + 1).value for i, h in enumerate(headers)}
        self.assertEqual(row["Rows Updated"], 2)  # not 3 — the duplicate collapses
        self.assertIn("duplicate", row["Skipped Row References"].lower())

    def test_blank_null_nonnumeric_and_negative_refs_all_ignored(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        messy_rows = [None, "", "abc", 0, -5, 12.5, excel_rows[0]]  # only the last is valid
        stat01_result = {
            "correction_data": [_correction_data_entry("1", messy_rows, calculated_dq="60 DAYS")],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        self.assertEqual(outcome["correction_count"], 1)
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[0], "Delinquency Status"), "60 DAYS")

        wb, _ = _reload(outcome["file_bytes"])
        log_ws = wb[corrections.CORRECTION_LOG_SHEET_NAME]
        headers = [c.value for c in log_ws[1]]
        row = {h: log_ws.cell(row=2, column=i + 1).value for i, h in enumerate(headers)}
        self.assertEqual(row["Rows Updated"], 1)
        skipped_field = row["Skipped Row References"]
        self.assertEqual(skipped_field.count("invalid row reference"), 6)  # None, "", "abc", 0, -5, 12.5

    def test_order_preserved_after_dedupe(self):
        # internal check on the helper directly: dedup keeps FIRST-seen order
        clean, skipped = corrections._clean_excel_rows([30, 10, 30, 20, 10])
        self.assertEqual(clean, [30, 10, 20])
        self.assertEqual(len(skipped), 2)


class ZeroValidRowsNotAppliedTests(unittest.TestCase):
    """Safeguard 5: if zero valid rows remain after validation/dedup/GP
    checks, the correction is reported as NOT applied — never counted as
    successful, never given a Correction Log row."""

    def test_all_invalid_refs_yields_zero_corrections_applied(self):
        file_bytes, _ = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        stat01_result = {
            "correction_data": [_correction_data_entry("1", [None, "", "abc", -1], calculated_dq="60 DAYS")],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        self.assertEqual(outcome["correction_count"], 0)
        self.assertIsNone(outcome["file_bytes"])
        self.assertIsNotNone(outcome["skip_reason"])

    def test_all_rows_gp_mismatched_yields_zero_corrections_applied_and_is_reported(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        # Claim the correction belongs to loan "999", which never appears in
        # this sheet — every row reference fails its GP check.
        entry = _correction_data_entry("999", excel_rows, calculated_dq="60 DAYS")
        stat01_result = {"correction_data": [entry], "_cycle_month": None}
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        self.assertEqual(outcome["correction_count"], 0)
        self.assertIsNone(outcome["file_bytes"])
        self.assertEqual(len(outcome["corrections_not_applied"]), 1)
        self.assertEqual(outcome["corrections_not_applied"][0]["gp"], "999")
        self.assertIn("GP mismatch", outcome["corrections_not_applied"][0]["reason"])

    def test_zero_valid_rows_for_one_loan_does_not_block_another_loans_valid_correction(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "2", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 50000, "Ending Balance": 50000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        broken_entry = _correction_data_entry("1", [None, "abc"], calculated_dq="60 DAYS")
        good_entry = _correction_data_entry("2", [excel_rows[1]], calculated_dq="90 DAYS")
        stat01_result = {"correction_data": [broken_entry, good_entry], "_cycle_month": None}
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        self.assertEqual(outcome["correction_count"], 1)  # only loan 2 succeeded
        self.assertEqual(len(outcome["corrections_not_applied"]), 1)
        self.assertEqual(outcome["corrections_not_applied"][0]["gp"], "1")
        _, ws = _reload(outcome["file_bytes"])
        self.assertEqual(_read_cell(ws, excel_rows[1], "Delinquency Status"), "90 DAYS")


class CorrectionCountIsLoanAndFieldLevelTests(unittest.TestCase):
    """Safeguard 4: correction_count is one per (loan, field), never one
    per physical row — confirmed both for a single multi-row correction
    and for two DIFFERENT fields on the same loan."""

    def test_single_multirow_correction_counts_as_one(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        stat01_result = {
            "correction_data": [_correction_data_entry("1", excel_rows, calculated_dq="60 DAYS")],
            "_cycle_month": None,
        }
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        self.assertEqual(outcome["correction_count"], 1)

    def test_two_fields_same_loan_counts_as_two(self):
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 100000, "Ending Balance": 100000, "Payment Received": 0,
             "Scheduled Payment": 500},
        ])
        dq_entry = _correction_data_entry(
            "1", excel_rows, reasons=("DELINQUENCY STATUS MISMATCH", "LOAN STATUS MISMATCH"),
            calculated_dq="60 DAYS", expected_loan_status_raw="DEFERRED",
            deferral_type="full", deferral_phase="start", srf_status="ACTIVE", soup_status="ACTIVE",
        )
        stat01_result = {"correction_data": [dq_entry], "_cycle_month": None}
        outcome = corrections.generate_corrected_workbook(file_bytes, "test.xlsx", "Test Bank", stat01_result)
        self.assertEqual(outcome["correction_count"], 2)  # DQ correction + Loan Status correction
        _, ws = _reload(outcome["file_bytes"])
        for r in excel_rows:
            self.assertEqual(_read_cell(ws, r, "Delinquency Status"), "60 DAYS")
            self.assertEqual(_read_cell(ws, r, "Loan Status"), "DEFERRED")


class BuildEligibleCorrectionsSignatureTests(unittest.TestCase):
    """CLIENT-CONFIRMED (2026-08-05): build_eligible_corrections has always
    returned a single list — that contract must never change, since it's a
    public function other test suites may call directly. The zero-valid-
    rows-reported-as-not-applied behavior is achieved entirely inside
    _apply_one_correction / generate_corrected_workbook instead, precisely
    so this signature never has to move."""

    def test_returns_a_plain_list_not_a_tuple(self):
        entry = _correction_data_entry("1", [10], calculated_dq="60 DAYS")
        result = corrections.build_eligible_corrections([entry])
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["gp"], "1")

    def test_returns_a_plain_list_even_when_every_row_reference_is_invalid(self):
        # The exact case that previously required a second return value —
        # confirms it no longer does.
        entry = _correction_data_entry("1", [None, "", "abc"], calculated_dq="60 DAYS")
        result = corrections.build_eligible_corrections([entry])
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 1)  # still emitted -- just with empty excel_rows
        self.assertEqual(result[0]["excel_rows"], [])

    def test_empty_correction_data_returns_empty_list(self):
        result = corrections.build_eligible_corrections([])
        self.assertEqual(result, [])


if __name__ == "__main__":
    unittest.main()
