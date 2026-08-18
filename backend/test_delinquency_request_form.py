"""Delinquency Information Request Form — the isolated report module.

CLIENT-CONFIRMED (2026-08-18). Covers the eligibility algorithm, the seven
columns and their sources, the receipt-date mapping, closed-at-zero suppression,
the workbook's reproduced structure, and the module's isolation from everything
it must not touch.

Endpoint wiring is a later commit and is not tested here.

THE ONE INVARIANT THAT OUTRANKS ALL OTHERS: the Servicing Notes / Additional
Information column is blank on every generated lender row. It is the lender's to
complete, and a value appearing there would be FBC answering its own question.
"""

import ast
import copy
import inspect
import io
import re
import unittest

import pandas as pd
from openpyxl import load_workbook

import checks
import corrections
import delinquency_request_form as drf


HEADER_ROW = 9


def _si(gp="1000000001", **overrides):
    """One servicer-inputs posting. Defaults are a clean, unremarkable loan."""
    row = {
        "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}",
        "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
        "Beginning Balance": 1000.0, "Ending Balance": 900.0,
        "Payment Received": 100.0, "Scheduled Payment": 100.0,
        "Servicing Fee": 0.0, "Net Interest": 10.0, "Principal": 90.0,
        "Liquidation Date": None, "Liquidation Loss": None,
        "Loan Payment Instance": 1, "Payment Receipt Date": "2026-07-05",
        "_excel_row": HEADER_ROW + 1, "_sheet_name": "Servicer Inputs",
        "_header_excel_row": HEADER_ROW,
    }
    row.update(overrides)
    return row


def _frames(postings, overrides=None):
    """(raw postings frame, final frame) — exactly what the caller supplies."""
    si = pd.DataFrame(postings)
    final = checks._apply_verified_overrides(
        checks.consolidate_cycle_postings(si), overrides or {})
    return si, final


def _rows(postings, overrides=None, excluded_gps=None):
    si, final = _frames(postings, overrides)
    return drf.build_lender_rows(final, si, excluded_gps=excluded_gps)


def _code_only(module):
    """Executable lines only. The module docstring legitimately names the things
    the module must not do in order to state that it does not do them."""
    source = inspect.getsource(module)
    body = source.split('"""')
    code = "".join(body[2::2]) if len(body) > 2 else source
    return "\n".join(l for l in code.splitlines() if not l.strip().startswith("#"))


def _sheet(blocks, cycle_month=None):
    buf = io.BytesIO()
    drf.generate_delinquency_request_form_bytes(blocks, buf, cycle_month)
    return load_workbook(buf)


# ---------------------------------------------------------------------------
# Eligibility — delinquency
# ---------------------------------------------------------------------------

class DelinquencyEligibilityTests(unittest.TestCase):
    """30 DAYS or worse, on the APPLICATION'S ladder."""

    EXPECTED = {
        "CURRENT": False,        # missed 0
        "DELINQUENT": False,     # missed 1 -- LESS severe than 30 DAYS
        "30 DAYS": True,         # missed 2 -- the threshold
        "60 DAYS": True,
        "90 DAYS": True,
        "120 DAYS": True,
        "150 DAYS": True,
        "150+ DAYS": True,
        "180 DAYS": True,
        "210 DAYS": True,
    }

    def test_every_permitted_value_individually(self):
        for status, expected in self.EXPECTED.items():
            with self.subTest(status=status):
                self.assertEqual(drf.delinquency_qualifies(status), expected)

    def test_the_expectations_cover_the_whole_permitted_domain(self):
        """So a value added upstream cannot slip past this suite unasserted."""
        self.assertEqual(set(self.EXPECTED), set(corrections._PERMITTED_DQ_STATUS))

    def test_the_boundary_sits_between_delinquent_and_30_days(self):
        self.assertEqual(checks._status_to_missed_count("DELINQUENT"), 1)
        self.assertEqual(checks._status_to_missed_count("30 DAYS"), 2)
        self.assertEqual(drf.MIN_MISSED_COUNT, 2)

    def test_case_and_whitespace_do_not_change_the_answer(self):
        for variant in ("60 days", "  60 DAYS  ", "60 Days"):
            with self.subTest(variant=variant):
                self.assertTrue(drf.delinquency_qualifies(variant))

    def test_an_off_domain_value_qualifies_rather_than_vanishing(self):
        """'60-DAYS' really occurs (2 July loans). It must never fail a numeric
        test in silence."""
        self.assertTrue(drf.delinquency_qualifies("60-DAYS"))

    def test_an_unparseable_typo_qualifies(self):
        """'DELIQNUENT' really occurs (1 July loan); the ladder returns None."""
        self.assertIsNone(checks._status_to_missed_count("DELIQNUENT"))
        self.assertTrue(drf.delinquency_qualifies("DELIQNUENT"))

    def test_a_cross_domain_value_qualifies(self):
        """'FORCLOSURE' really occurs in the Delinquency column (1 July loan)."""
        self.assertTrue(drf.delinquency_qualifies("FORCLOSURE"))

    def test_a_blank_does_not_qualify(self):
        for blank in (None, "", "   ", float("nan"), "nan", "None"):
            with self.subTest(blank=blank):
                self.assertFalse(drf.delinquency_qualifies(blank))


# ---------------------------------------------------------------------------
# Eligibility — loan status
# ---------------------------------------------------------------------------

class LoanStatusEligibilityTests(unittest.TestCase):
    """Anything other than ordinary ACTIVE."""

    def test_active_alone_does_not_qualify(self):
        self.assertFalse(drf.loan_status_qualifies("ACTIVE"))

    def test_every_other_permitted_value_qualifies(self):
        others = corrections._PERMITTED_LOAN_STATUS - {"ACTIVE"}
        self.assertEqual(len(others), 8)
        for status in sorted(others):
            with self.subTest(status=status):
                self.assertTrue(drf.loan_status_qualifies(status))

    def test_the_eight_are_the_client_approved_set(self):
        self.assertEqual(
            corrections._PERMITTED_LOAN_STATUS - {"ACTIVE"},
            {"DEFERRED", "LIQUIDATION", "FORECLOSURE", "REO", "BANKRUPTCY",
             "PAID-IN-FULL", "SUBSTITUTION-IN", "SUBSTITUTION-OUT"})

    def test_an_off_domain_loan_status_qualifies(self):
        """The rule is a negation, not an allow-list, precisely so an
        unexpected or misspelled status is surfaced rather than dropped."""
        self.assertTrue(drf.loan_status_qualifies("FORCLOSURE"))
        self.assertTrue(drf.loan_status_qualifies("CHARGE-OFF"))

    def test_a_blank_does_not_qualify(self):
        for blank in (None, "", "   ", float("nan")):
            with self.subTest(blank=blank):
                self.assertFalse(drf.loan_status_qualifies(blank))


# ---------------------------------------------------------------------------
# Eligibility — the union
# ---------------------------------------------------------------------------

class UnionTests(unittest.TestCase):

    def test_active_and_current_is_excluded(self):
        self.assertEqual(_rows([_si()]), [])

    def test_active_qualifies_when_the_delinquency_does(self):
        rows = _rows([_si(**{"Delinquency Status": "60 DAYS"})])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Loan Status"], "ACTIVE")

    def test_a_non_active_status_qualifies_while_current(self):
        rows = _rows([_si(**{"Loan Status": "DEFERRED"})])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Delinquency Status"], "CURRENT")

    def test_a_loan_qualifying_on_both_appears_exactly_once(self):
        rows = _rows([_si(**{"Loan Status": "LIQUIDATION",
                             "Delinquency Status": "120 DAYS"})])
        self.assertEqual(len(rows), 1)

    def test_delinquent_alone_is_excluded(self):
        self.assertEqual(_rows([_si(**{"Delinquency Status": "DELINQUENT"})]), [])

    def test_delinquent_with_a_non_active_status_is_included(self):
        rows = _rows([_si(**{"Loan Status": "BANKRUPTCY",
                             "Delinquency Status": "DELINQUENT"})])
        self.assertEqual(len(rows), 1)

    def test_rows_are_ordered_by_sba_gp_number(self):
        rows = _rows([_si("3000000003", **{"Loan Status": "REO"}),
                      _si("1000000001", **{"Loan Status": "REO"}),
                      _si("2000000002", **{"Loan Status": "REO"})])
        self.assertEqual([r["SBA GP Number"] for r in rows],
                         ["1000000001", "2000000002", "3000000003"])

    def test_no_second_status_ladder_is_defined_in_this_module(self):
        """Severity must be asked of the application, never restated."""
        code = _code_only(drf)
        self.assertIn("checks._status_to_missed_count", code)
        for restated in ("90 DAYS", "120 DAYS", "150 DAYS", "180 DAYS", "210 DAYS"):
            with self.subTest(value=restated):
                self.assertNotIn(restated, code)


# ---------------------------------------------------------------------------
# Final values, not submitted ones
# ---------------------------------------------------------------------------

class FinalValueTests(unittest.TestCase):
    """The form reflects FBC's determination after Human Review."""

    def test_a_loan_corrected_into_scope_appears(self):
        rows = _rows([_si()], {("1000000001", "Delinquency Status"): "60 DAYS"})
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Delinquency Status"], "60 DAYS")

    def test_a_loan_corrected_out_of_scope_does_not_appear(self):
        rows = _rows([_si(**{"Delinquency Status": "60 DAYS"})],
                     {("1000000001", "Delinquency Status"): "CURRENT"})
        self.assertEqual(rows, [])

    def test_a_loan_status_correction_brings_a_loan_into_scope(self):
        rows = _rows([_si()], {("1000000001", "Loan Status"): "DEFERRED"})
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Loan Status"], "DEFERRED")

    def test_a_loan_status_correction_can_take_a_loan_out_of_scope(self):
        rows = _rows([_si(**{"Loan Status": "LIQUIDATION"})],
                     {("1000000001", "Loan Status"): "ACTIVE"})
        self.assertEqual(rows, [])

    def test_the_balance_is_the_final_overridden_value(self):
        rows = _rows([_si(**{"Loan Status": "REO"})],
                     {("1000000001", "Ending Balance"): 145729.87})
        self.assertEqual(rows[0]["Securitization Balance"], 145729.87)

    def test_an_off_domain_value_is_preserved_verbatim(self):
        rows = _rows([_si(**{"Delinquency Status": "60-DAYS"})])
        self.assertEqual(rows[0]["Delinquency Status"], "60-DAYS")


# ---------------------------------------------------------------------------
# Column sources
# ---------------------------------------------------------------------------

class ColumnSourceTests(unittest.TestCase):

    def setUp(self):
        self.rows = _rows([_si(**{"Loan Status": "FORECLOSURE"})])

    def test_the_seven_columns_are_exactly_the_approved_set(self):
        self.assertEqual(drf.COLUMNS, [
            "SBA GP Number", "Lender Loan Number", "Securitization Balance",
            "Loan Status", "Delinquency Status", "Last Payment Receipt Date",
            "Servicing Notes / Additional Information"])

    def test_gp_lender_number_and_balance(self):
        row = self.rows[0]
        self.assertEqual(row["SBA GP Number"], "1000000001")
        self.assertEqual(row["Lender Loan Number"], "L-1000000001")
        self.assertEqual(row["Securitization Balance"], 900.0)

    def test_a_missing_lender_loan_number_is_blank_never_fabricated(self):
        rows = _rows([_si(**{"Loan Status": "REO", "Lender Loan Number": None})])
        self.assertEqual(rows[0]["Lender Loan Number"], "")

    def test_a_numeric_lender_loan_number_loses_its_float_artifact(self):
        """Found on real July data: pandas delivers 2010300146.0, and
        "2010300146.0" is not the lender's own reference."""
        rows = _rows([_si(**{"Loan Status": "REO",
                             "Lender Loan Number": 2010300146.0})])
        self.assertEqual(rows[0]["Lender Loan Number"], "2010300146")

    def test_a_string_lender_loan_number_keeps_its_leading_zeros(self):
        """The client's own form carries "0000775"."""
        rows = _rows([_si(**{"Loan Status": "REO", "Lender Loan Number": "0000775"})])
        self.assertEqual(rows[0]["Lender Loan Number"], "0000775")

    def test_a_fractional_value_is_left_alone_rather_than_rounded(self):
        rows = _rows([_si(**{"Loan Status": "REO", "Lender Loan Number": 12.5})])
        self.assertEqual(rows[0]["Lender Loan Number"], "12.5")

    def test_the_sba_gp_number_is_unaffected(self):
        self.assertEqual(self.rows[0]["SBA GP Number"], "1000000001")


# ---------------------------------------------------------------------------
# Last Payment Receipt Date
# ---------------------------------------------------------------------------

class ReceiptDateTests(unittest.TestCase):
    """Read from the posting consolidation ALREADY chose as authoritative."""

    def test_it_comes_from_the_authoritative_posting(self):
        rows = _rows([_si(**{"Loan Status": "REO"})])
        self.assertEqual(str(rows[0]["Last Payment Receipt Date"])[:10], "2026-07-05")

    def test_with_several_postings_it_is_the_last_one_consolidation_used(self):
        postings = [
            _si(**{"Loan Payment Instance": 1, "Payment Receipt Date": "2026-07-05",
                   "_excel_row": 10}),
            _si(**{"Loan Payment Instance": 2, "Payment Receipt Date": "2026-07-20",
                   "_excel_row": 11, "Loan Status": "REO"}),
        ]
        si, final = _frames(postings)
        # Consolidation's own choice of authoritative row.
        self.assertEqual(final.at["1000000001", "_excel_row"], 11)
        rows = drf.build_lender_rows(final, si)
        self.assertEqual(str(rows[0]["Last Payment Receipt Date"])[:10], "2026-07-20")

    def test_a_blank_receipt_date_stays_blank(self):
        """A non-paying loan has no receipt date. Nothing reaches backwards into
        an earlier posting to fill the cell."""
        postings = [
            _si(**{"Loan Payment Instance": 1, "Payment Receipt Date": "2026-06-05",
                   "_excel_row": 10}),
            _si(**{"Loan Payment Instance": 2, "Payment Receipt Date": None,
                   "_excel_row": 11, "Loan Status": "REO"}),
        ]
        si, final = _frames(postings)
        rows = drf.build_lender_rows(final, si)
        self.assertIsNone(rows[0]["Last Payment Receipt Date"])

    def test_it_is_blank_when_the_column_is_absent_entirely(self):
        postings = [_si(**{"Loan Status": "REO"})]
        del postings[0]["Payment Receipt Date"]
        si, final = _frames(postings)
        rows = drf.build_lender_rows(final, si)
        self.assertIsNone(rows[0]["Last Payment Receipt Date"])

    def test_it_is_blank_when_no_raw_frame_is_supplied(self):
        _si_df, final = _frames([_si(**{"Loan Status": "REO"})])
        rows = drf.build_lender_rows(final, None)
        self.assertIsNone(rows[0]["Last Payment Receipt Date"])

    def test_rows_on_another_sheet_cannot_supply_the_date(self):
        """The index is keyed on sheet AND row, so a repeated row number on a
        second sheet cannot leak its date onto this loan."""
        si, final = _frames([_si(**{"Loan Status": "REO"})])
        si.loc[0, "_sheet_name"] = "Some Other Sheet"
        rows = drf.build_lender_rows(final, si)
        self.assertIsNone(rows[0]["Last Payment Receipt Date"])


# ---------------------------------------------------------------------------
# Closed-at-zero suppression
# ---------------------------------------------------------------------------

class SuppressionTests(unittest.TestCase):

    def test_a_suppressed_loan_is_omitted_even_when_it_qualifies(self):
        rows = _rows([_si(**{"Loan Status": "PAID-IN-FULL"})],
                     excluded_gps=["1000000001"])
        self.assertEqual(rows, [])

    def test_suppression_matches_on_the_normalised_gp(self):
        rows = _rows([_si(**{"Loan Status": "PAID-IN-FULL"})],
                     excluded_gps=["  1000000001  "])
        self.assertEqual(rows, [])

    def test_an_empty_or_absent_set_changes_nothing(self):
        for excluded in ([], None):
            with self.subTest(excluded=excluded):
                rows = _rows([_si(**{"Loan Status": "REO"})], excluded_gps=excluded)
                self.assertEqual(len(rows), 1)

    def test_the_suppression_set_is_never_recomputed_here(self):
        code = _code_only(drf)
        for forbidden in ("loans_closed_at_zero", "suppressed_closed_loans",
                          "_CLOSED_AT_ZERO"):
            with self.subTest(name=forbidden):
                self.assertNotIn(forbidden, code)


# ---------------------------------------------------------------------------
# The notes column
# ---------------------------------------------------------------------------

class NotesColumnTests(unittest.TestCase):
    """ALWAYS blank. The lender completes it."""

    def test_the_row_carries_an_empty_note(self):
        rows = _rows([_si(**{"Loan Status": "REO"})])
        self.assertEqual(rows[0][drf.NOTES_COLUMN], "")

    def test_the_cell_is_blank_on_every_generated_row(self):
        rows = _rows([_si(f"100000000{i}", **{"Loan Status": "REO"}) for i in range(1, 6)])
        self.assertEqual(len(rows), 5)
        ws = _sheet([("Lendistry 2025", rows)])["Lendistry 2025"]
        notes_col = drf.FIRST_COLUMN + 6
        for i in range(len(rows)):
            with self.subTest(row=i):
                self.assertIsNone(ws.cell(row=drf.FIRST_DATA_ROW + i,
                                          column=notes_col).value)

    def test_it_stays_blank_even_if_a_row_dict_carries_text(self):
        """Defence in depth: the writer blanks the column regardless of input."""
        rows = _rows([_si(**{"Loan Status": "REO"})])
        rows[0][drf.NOTES_COLUMN] = "FBC SHOULD NOT WRITE THIS"
        ws = _sheet([("Lendistry 2025", rows)])["Lendistry 2025"]
        self.assertIsNone(ws.cell(row=drf.FIRST_DATA_ROW,
                                  column=drf.FIRST_COLUMN + 6).value)

    def test_the_heading_still_describes_the_column(self):
        ws = _sheet([("Lendistry 2025", [])])["Lendistry 2025"]
        self.assertIn("Servicing Notes",
                      ws.cell(row=drf.HEADER_ROW, column=drf.FIRST_COLUMN + 6).value)


# ---------------------------------------------------------------------------
# Workbook structure and formatting
# ---------------------------------------------------------------------------

class WorkbookTests(unittest.TestCase):

    def setUp(self):
        self.rows = _rows([_si("100000000%d" % i, **{"Loan Status": "REO"})
                           for i in range(1, 4)])
        self.wb = _sheet([("Lendistry 2025", self.rows)], pd.Timestamp("2026-07-01"))
        self.ws = self.wb["Lendistry 2025"]

    def test_the_title_is_the_supplied_one(self):
        cell = self.ws.cell(row=drf.TITLE_ROW, column=drf.FIRST_COLUMN)
        self.assertEqual(cell.value, "Delinquency Information Request Form")
        self.assertEqual(cell.font.size, 16)
        self.assertTrue(cell.font.bold)
        self.assertEqual(cell.font.name, "Aptos Narrow")

    def test_headers_sit_on_row_4_and_data_starts_on_row_5(self):
        self.assertEqual(self.ws.cell(row=4, column=2).value, "SBA GP Number")
        self.assertEqual(self.ws.cell(row=5, column=2).value, "1000000001")

    def test_column_a_is_the_narrow_spacer(self):
        self.assertIsNone(self.ws.cell(row=4, column=1).value)
        self.assertAlmostEqual(self.ws.column_dimensions["A"].width, 1.86, places=2)

    def test_every_column_width_matches_the_supplied_form(self):
        for letter, width in drf.COLUMN_WIDTHS.items():
            with self.subTest(column=letter):
                self.assertAlmostEqual(
                    self.ws.column_dimensions[letter].width, width, places=2)

    def test_the_supplied_row_heights_are_reproduced(self):
        for row, height in drf.ROW_HEIGHTS.items():
            with self.subTest(row=row):
                self.assertAlmostEqual(self.ws.row_dimensions[row].height, height,
                                       places=2)

    def test_the_notes_heading_is_verbatim_including_its_missing_space(self):
        value = self.ws.cell(row=4, column=drf.FIRST_COLUMN + 6).value
        self.assertTrue(value.startswith("Servicing Notes /Additional Information - "))
        self.assertIn("last date of Borrower contact", value)

    def test_the_balance_heading_carries_the_month_end(self):
        self.assertEqual(self.ws.cell(row=4, column=drf.FIRST_COLUMN + 2).value,
                         "Securitization Balance as of 7/31/26")

    def test_number_formats(self):
        self.assertEqual(self.ws.cell(row=5, column=drf.FIRST_COLUMN + 2).number_format,
                         "#,##0")
        self.assertEqual(self.ws.cell(row=5, column=drf.FIRST_COLUMN + 5).number_format,
                         "mm-dd-yy")

    def test_data_rows_are_banded_on_even_rows(self):
        self.assertIsNone(self.ws.cell(row=5, column=2).fill.fill_type)
        self.assertEqual(self.ws.cell(row=6, column=2).fill.fill_type, "solid")
        self.assertIsNone(self.ws.cell(row=7, column=2).fill.fill_type)

    def test_data_cells_are_bordered_and_centred(self):
        cell = self.ws.cell(row=5, column=2)
        self.assertEqual(cell.border.left.style, "thin")
        self.assertEqual(cell.border.bottom.style, "thin")
        self.assertEqual(cell.alignment.horizontal, "center")
        self.assertEqual(cell.font.size, 10)

    def test_headers_are_wrapped_and_filled(self):
        cell = self.ws.cell(row=4, column=2)
        self.assertTrue(cell.alignment.wrap_text)
        self.assertEqual(cell.fill.fill_type, "solid")
        self.assertFalse(cell.font.bold)
        self.assertEqual(cell.font.size, 11)

    def test_the_balance_is_written_as_a_number_not_text(self):
        """Numeric, so #,##0 formats it and Excel can total the column. openpyxl
        reads a whole value back as int, so the assertion is numeric-ness rather
        than float specifically."""
        value = self.ws.cell(row=5, column=drf.FIRST_COLUMN + 2).value
        self.assertIsInstance(value, (int, float))
        self.assertNotIsInstance(value, bool)
        self.assertEqual(float(value), 900.0)

    def test_the_receipt_date_is_written_as_a_date_not_text(self):
        import datetime
        self.assertIsInstance(self.ws.cell(row=5, column=drf.FIRST_COLUMN + 5).value,
                              datetime.datetime)


class BalanceHeadingTests(unittest.TestCase):

    def test_month_ends_are_derived_not_guessed(self):
        for month, expected in (("2026-07-01", "7/31/26"), ("2026-02-01", "2/28/26"),
                                ("2026-12-01", "12/31/26"), ("2024-02-01", "2/29/24")):
            with self.subTest(month=month):
                self.assertEqual(drf.balance_heading(pd.Timestamp(month)),
                                 f"Securitization Balance as of {expected}")

    def test_an_unknown_cycle_month_degrades_rather_than_guessing(self):
        self.assertEqual(drf.balance_heading(None), "Securitization Balance")
        self.assertEqual(drf.balance_heading("not a date"), "Securitization Balance")


# ---------------------------------------------------------------------------
# Worksheet structure
# ---------------------------------------------------------------------------

class WorksheetStructureTests(unittest.TestCase):

    def test_one_sheet_per_lender_and_deal_in_upload_order(self):
        wb = _sheet([("Community Bank & Trust 2025", []), ("Lendistry 2024", []),
                     ("Lendistry 2025", [])])
        self.assertEqual(wb.sheetnames,
                         ["Community Bank & Trust 2025", "Lendistry 2024",
                          "Lendistry 2025"])

    def test_one_lender_in_two_deals_keeps_two_separate_sheets(self):
        rows_2024 = _rows([_si("1000000001", **{"Loan Status": "REO"})])
        rows_2025 = _rows([_si("2000000002", **{"Loan Status": "REO"})])
        wb = _sheet([("Lendistry 2024", rows_2024), ("Lendistry 2025", rows_2025)])
        self.assertEqual(wb["Lendistry 2024"].cell(row=5, column=2).value, "1000000001")
        self.assertEqual(wb["Lendistry 2025"].cell(row=5, column=2).value, "2000000002")
        self.assertNotEqual(wb["Lendistry 2024"].cell(row=5, column=2).value,
                            wb["Lendistry 2025"].cell(row=5, column=2).value)

    def test_a_lender_with_no_qualifying_loans_keeps_its_sheet_and_says_so(self):
        wb = _sheet([("Poppy Bank", [])])
        ws = wb["Poppy Bank"]
        self.assertIn("Poppy Bank", wb.sheetnames)
        self.assertEqual(ws.cell(row=drf.FIRST_DATA_ROW, column=drf.FIRST_COLUMN).value,
                         "No loans requiring additional information.")
        self.assertEqual(ws.cell(row=4, column=2).value, "SBA GP Number")

    def test_illegal_sheet_characters_are_removed(self):
        wb = _sheet([("Bank/Trust: [2025]*?", [])])
        self.assertEqual(wb.sheetnames, ["BankTrust 2025"])

    def test_a_long_label_is_truncated_to_excels_limit(self):
        wb = _sheet([("A" * 60, [])])
        self.assertEqual(len(wb.sheetnames[0]), 31)

    def test_colliding_labels_get_distinct_sheets(self):
        wb = _sheet([("Lendistry", []), ("Lendistry", [])])
        self.assertEqual(len(wb.sheetnames), 2)
        self.assertEqual(wb.sheetnames[0], "Lendistry")
        self.assertNotEqual(wb.sheetnames[0], wb.sheetnames[1])

    def test_a_workbook_with_no_lenders_still_saves(self):
        wb = _sheet([])
        self.assertEqual(len(wb.sheetnames), 1)


class FilenameTests(unittest.TestCase):

    def test_the_approved_convention(self):
        self.assertEqual(
            drf.delinquency_request_form_filename(pd.Timestamp("2026-07-01")),
            "Delinquency_Information_Request_Form_July_2026.xlsx")

    def test_an_unknown_period_is_omitted_never_guessed(self):
        self.assertEqual(drf.delinquency_request_form_filename(None),
                         "Delinquency_Information_Request_Form.xlsx")
        self.assertEqual(drf.delinquency_request_form_filename("nonsense"),
                         "Delinquency_Information_Request_Form.xlsx")


# ---------------------------------------------------------------------------
# Isolation
# ---------------------------------------------------------------------------

class IsolationTests(unittest.TestCase):
    """A reporting consumer, and provably nothing more."""

    def test_it_imports_only_what_a_reporting_consumer_needs(self):
        names = set()
        for node in ast.walk(ast.parse(inspect.getsource(drf))):
            if isinstance(node, ast.Import):
                names.update(a.name.split(".")[0] for a in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                names.add(node.module.split(".")[0])
        self.assertEqual(names, {"re", "pandas", "openpyxl", "checks", "corrections"})

    def test_it_never_imports_main_or_any_other_report(self):
        code = _code_only(drf)
        for forbidden in ("import main", "monthly_servicer_report", "mismatch_report",
                          "import report", "preflight", "chatbot"):
            with self.subTest(name=forbidden):
                self.assertNotIn(forbidden, code)

    def test_it_calls_no_human_review_function(self):
        code = _code_only(drf)
        for forbidden in ("generate_reviewed_workbook", "build_recommendations",
                          "validate_approved_value", "build_eligible_corrections",
                          "classify_recommendation"):
            with self.subTest(name=forbidden):
                self.assertNotIn(forbidden, code)

    def test_it_runs_no_audit_check_and_writes_no_correction(self):
        code = _code_only(drf)
        for forbidden in ("run_servicer_reconciliation", "consolidate_cycle_postings",
                          "_apply_verified_overrides", "STAT-01", "POP-01", "BAL-01"):
            with self.subTest(name=forbidden):
                self.assertNotIn(forbidden, code)

    def test_it_touches_no_session_or_request_state(self):
        """`request` is matched as a WHOLE WORD: the module's own name carries
        it inside `delinquency_request_form` and
        `Delinquency_Information_Request_Form`, where a substring test would
        fail on the filename constant rather than on real web-layer state."""
        code = _code_only(drf)
        for forbidden in ("SESSION_CACHE", "fastapi", "JSONResponse"):
            with self.subTest(name=forbidden):
                self.assertNotIn(forbidden, code)
        # Lowercase only: "Request" is a legitimate word in this form's own
        # title, "Delinquency Information Request Form". The import-set test
        # above is what actually makes web-layer state unreachable.
        self.assertIsNone(re.search(r"\brequest\b", code))

    def test_no_entry_point_takes_a_request_or_session_argument(self):
        for fn in (drf.build_lender_rows, drf.generate_delinquency_request_form_bytes,
                   drf.delinquency_request_form_filename, drf.worksheet_title,
                   drf.balance_heading):
            params = set(inspect.signature(fn).parameters)
            with self.subTest(fn=fn.__name__):
                self.assertEqual(params & {"request", "session", "session_id", "app"},
                                 set())

    def test_main_uses_only_this_module_s_public_entry_points(self):
        """main.py IS the consumer as of Step 2.

        This replaces Step 1's "main does not reference the module yet", which
        pinned that commit's not-yet-wired claim and expired the moment the
        wiring arrived. The invariant worth keeping is the DIRECTION of the
        dependency: main reaches in only through the documented entry points, so
        this module's internals stay free to change and the separability the
        tests above assert is not quietly undone by a caller binding to a
        private name."""
        import main
        used = set(re.findall(r"drf_mod\.(\w+)", inspect.getsource(main)))
        self.assertTrue(used)
        self.assertEqual(used - {"build_lender_rows",
                                 "generate_delinquency_request_form_bytes",
                                 "delinquency_request_form_filename",
                                 "NO_QUALIFYING_LOANS_TEXT", "COLUMNS"}, set())
        self.assertEqual([n for n in used if n.startswith("_")], [])


class PurityTests(unittest.TestCase):
    """Inputs are read, never written."""

    def test_the_final_frame_is_not_mutated(self):
        si, final = _frames([_si(**{"Loan Status": "REO"})])
        before = final.copy(deep=True)
        drf.build_lender_rows(final, si)
        pd.testing.assert_frame_equal(final, before)

    def test_the_raw_postings_frame_is_not_mutated(self):
        si, final = _frames([_si(**{"Loan Status": "REO"})])
        before = si.copy(deep=True)
        drf.build_lender_rows(final, si)
        pd.testing.assert_frame_equal(si, before)

    def test_the_suppression_set_is_not_mutated(self):
        excluded = ["1000000001"]
        before = copy.deepcopy(excluded)
        _rows([_si(**{"Loan Status": "REO"})], excluded_gps=excluded)
        self.assertEqual(excluded, before)

    def test_an_empty_frame_produces_no_rows(self):
        self.assertEqual(drf.build_lender_rows(pd.DataFrame(), None), [])
        self.assertEqual(drf.build_lender_rows(None, None), [])


if __name__ == "__main__":
    unittest.main()
