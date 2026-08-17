"""Submitted Information vs. FBC Mismatch Report — CLIENT-CONFIRMED (2026-08-17).

The report answers one question: where does the information submitted to FBC
differ from FBC's final determination after Human Review?

These tests run WITHOUT main.py. Every entry point takes plain values -- two
DataFrames, a servicer key, a list of review records, a buffer -- so the module
is exercised directly and Commit 3's wiring stays provably separable. One test
asserts that separability explicitly.
"""

import ast
import inspect
import io
import re
import unittest

import pandas as pd
from openpyxl import load_workbook

import checks
import mismatch_report as mr


def _si(**overrides):
    """One servicer-inputs posting. Defaults are a clean, unremarkable loan."""
    row = {
        "SBA GP Number": "1000000001", "Lender Loan Number": "L-1",
        "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
        "Beginning Balance": 1000.0, "Ending Balance": 900.0,
        "Payment Received": 100.0, "Scheduled Payment": 100.0,
        "Servicing Fee": 0.0, "Net Interest": 10.0, "Principal": 90.0,
        "Liquidation Date": None, "Liquidation Loss": None,
        "Loan Payment Instance": 1,
    }
    row.update(overrides)
    return row


def _frames(postings, overrides=None):
    """(submitted, final) — exactly the two frames the application builds."""
    submitted = checks.consolidate_cycle_postings(pd.DataFrame(postings))
    final = checks._apply_verified_overrides(submitted, overrides or {})
    return submitted, final


def _code_only(module):
    """Executable lines only, docstrings and comments stripped. The module
    docstring legitimately names the things the module must not do -- in-sheet
    labels, audit checks, business rules -- in order to state that it does not
    do them."""
    source = inspect.getsource(module)
    body = source.split('"""')
    code = "".join(body[2::2]) if len(body) > 2 else source
    return "\n".join(l for l in code.splitlines() if not l.strip().startswith("#"))


def _log(gp, field, reason=None, note=None):
    return {"gp": gp, "field": field, "review_reason": reason, "reviewer_note": note}


def _rows(postings, overrides=None, log_rows=None, gaps=None, soup_key="Lendistry"):
    submitted, final = _frames(postings, overrides)
    return mr.build_mismatch_rows(submitted, final, soup_key, log_rows, gaps)


# ---------------------------------------------------------------------------
# Detection
# ---------------------------------------------------------------------------

class DetectionTests(unittest.TestCase):

    def test_an_approved_change_produces_a_row(self):
        rows = _rows([_si()], {("1000000001", "Loan Status"): "DEFERRED"})
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Mismatch Type"], "Loan Status")
        self.assertEqual(rows[0]["Submitted Information"], "ACTIVE")
        self.assertEqual(rows[0]["FBC Determination"], "DEFERRED")

    def test_a_rejection_produces_no_row(self):
        """A rejection contributes no override, so submitted == final."""
        rows = _rows([_si()], {})
        self.assertEqual(rows, [])

    def test_a_reviewer_third_value_is_compared_against_the_final_value(self):
        """Neither the submitted value nor the engine's recommendation is the
        determination -- the reviewer's own selection is."""
        rows = _rows([_si(**{"Loan Status": "LIQUIDATION"})],
                     {("1000000001", "Loan Status"): "BANKRUPTCY"})
        self.assertEqual(rows[0]["Submitted Information"], "LIQUIDATION")
        self.assertEqual(rows[0]["FBC Determination"], "BANKRUPTCY")

    def test_all_five_field_types_are_detected(self):
        postings = [_si(**{"Loan Status": "LIQUIDATION", "Delinquency Status": "120 DAYS",
                           "Ending Balance": 0.0, "Liquidation Date": "2026-03-27",
                           "Liquidation Loss": 145729.87})]
        overrides = {
            ("1000000001", "Loan Status"): "ACTIVE",
            ("1000000001", "Delinquency Status"): "CURRENT",
            ("1000000001", "Ending Balance"): 145729.87,
            ("1000000001", "Liquidation Date"): "",
            ("1000000001", "Liquidation Loss"): None,
        }
        rows = _rows(postings, overrides)
        self.assertEqual([r["Mismatch Type"] for r in rows], list(mr.MISMATCH_FIELDS))

    def test_a_liquidation_reversal_yields_four_separate_rows(self):
        """One reviewer decision, four real field differences. Cascade
        consequences are never collapsed."""
        postings = [_si(**{"Loan Status": "LIQUIDATION", "Ending Balance": 0.0,
                           "Liquidation Date": "2026-03-27", "Liquidation Loss": 145729.87})]
        overrides = {
            ("1000000001", "Loan Status"): "ACTIVE",
            ("1000000001", "Ending Balance"): 145729.87,
            ("1000000001", "Liquidation Date"): "",
            ("1000000001", "Liquidation Loss"): None,
        }
        rows = _rows(postings, overrides)
        self.assertEqual([r["Mismatch Type"] for r in rows],
                         ["Loan Status", "Ending Balance", "Liquidation Date",
                          "Liquidation Loss"])

    def test_an_unchanged_field_produces_nothing(self):
        rows = _rows([_si()], {("1000000001", "Loan Status"): "ACTIVE"})
        self.assertEqual(rows, [])

    def test_a_loan_with_no_changes_produces_nothing(self):
        rows = _rows([_si(), _si(**{"SBA GP Number": "1000000002"})], {})
        self.assertEqual(rows, [])


# ---------------------------------------------------------------------------
# Normalisation
# ---------------------------------------------------------------------------

class NormalisationTests(unittest.TestCase):
    """Formatting must not create a false mismatch; financial meaning must not
    be quietly changed."""

    def test_none_blank_whitespace_and_nan_are_all_the_same_absence(self):
        for a in (None, "", "   ", float("nan")):
            for b in (None, "", "   ", float("nan")):
                with self.subTest(a=repr(a), b=repr(b)):
                    self.assertEqual(mr._comparable("Liquidation Date", a),
                                     mr._comparable("Liquidation Date", b))

    def test_a_case_only_difference_is_not_a_mismatch(self):
        rows = _rows([_si()], {("1000000001", "Loan Status"): "Active"})
        self.assertEqual(rows, [])

    def test_a_whitespace_only_difference_is_not_a_mismatch(self):
        rows = _rows([_si()], {("1000000001", "Loan Status"): "  ACTIVE  "})
        self.assertEqual(rows, [])

    def test_float_noise_within_a_cent_is_not_a_mismatch(self):
        rows = _rows([_si(**{"Ending Balance": 145729.87})],
                     {("1000000001", "Ending Balance"): 145729.870000001})
        self.assertEqual(rows, [])

    def test_zero_versus_absent_IS_a_mismatch(self):
        """The line that must not be crossed: 0.00 and "no value" are different
        financial facts."""
        rows = _rows([_si(**{"Liquidation Loss": 0.0})],
                     {("1000000001", "Liquidation Loss"): None})
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Submitted Information"], "0.00")
        self.assertEqual(rows[0]["FBC Determination"], "")

    def test_absent_versus_zero_IS_a_mismatch_in_the_other_direction(self):
        rows = _rows([_si(**{"Liquidation Loss": None})],
                     {("1000000001", "Liquidation Loss"): 0.0})
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Submitted Information"], "")
        self.assertEqual(rows[0]["FBC Determination"], "0.00")

    def test_a_date_is_equal_across_string_timestamp_and_date(self):
        base = mr._comparable("Liquidation Date", "2026-03-27")
        for other in (pd.Timestamp("2026-03-27"), pd.Timestamp("2026-03-27").date(),
                      "2026-03-27 00:00:00"):
            with self.subTest(other=repr(other)):
                self.assertEqual(mr._comparable("Liquidation Date", other), base)

    def test_a_date_present_versus_absent_IS_a_mismatch(self):
        rows = _rows([_si(**{"Liquidation Date": "2026-03-27"})],
                     {("1000000001", "Liquidation Date"): ""})
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Submitted Information"], "2026-03-27")
        self.assertEqual(rows[0]["FBC Determination"], "")

    def test_an_unparseable_date_falls_back_to_text_comparison(self):
        self.assertEqual(mr._comparable("Liquidation Date", "not a date"),
                         ("text", "NOT A DATE"))
        self.assertNotEqual(mr._comparable("Liquidation Date", "not a date"),
                            mr._comparable("Liquidation Date", "also not a date"))

    def test_money_never_compares_equal_to_the_same_text(self):
        self.assertNotEqual(mr._comparable("Ending Balance", 0.0),
                            mr._comparable("Loan Status", "0.00"))

    def test_the_display_value_is_the_raw_value_not_the_comparison_key(self):
        """Normalisation may suppress a difference; it may never change what a
        reader sees."""
        rows = _rows([_si(**{"Loan Status": "Paid-in-Full"})],
                     {("1000000001", "Loan Status"): "ACTIVE"})
        self.assertEqual(rows[0]["Submitted Information"], "Paid-in-Full")


# ---------------------------------------------------------------------------
# Identity
# ---------------------------------------------------------------------------

class IdentityTests(unittest.TestCase):

    def test_lender_is_the_supplied_soup_key(self):
        rows = _rows([_si()], {("1000000001", "Loan Status"): "DEFERRED"},
                     soup_key="Community Bank & Trust")
        self.assertEqual(rows[0]["Lender"], "Community Bank & Trust")

    def test_lender_is_never_taken_from_an_in_sheet_label(self):
        code = _code_only(mr)
        for label in ("► Servicer", "Servicer:", "Lender Name"):
            with self.subTest(label=label):
                self.assertNotIn(label, code)

    def test_the_gp_comes_from_the_frame_index(self):
        rows = _rows([_si()], {("1000000001", "Loan Status"): "DEFERRED"})
        self.assertEqual(rows[0]["SBA GP Number"], "1000000001")

    def test_the_lender_loan_number_is_read_from_the_frame(self):
        rows = _rows([_si(**{"Lender Loan Number": "ABC-99"})],
                     {("1000000001", "Loan Status"): "DEFERRED"})
        self.assertEqual(rows[0]["Lender Loan Number"], "ABC-99")

    def test_a_missing_lender_loan_number_is_blank_never_fabricated(self):
        rows = _rows([_si(**{"Lender Loan Number": None})],
                     {("1000000001", "Loan Status"): "DEFERRED"})
        self.assertEqual(rows[0]["Lender Loan Number"], "")


# ---------------------------------------------------------------------------
# Comments
# ---------------------------------------------------------------------------

class CommentTests(unittest.TestCase):

    def test_the_comparison_sentence_format(self):
        rows = _rows([_si(**{"Loan Status": "LIQUIDATION"})],
                     {("1000000001", "Loan Status"): "ACTIVE"})
        self.assertEqual(
            rows[0]["Comments"],
            "Submitted Loan Status: LIQUIDATION. FBC Determination: ACTIVE. "
            + mr.NO_REVIEW_RECORD)

    def test_an_existing_review_reason_is_appended_verbatim(self):
        rows = _rows([_si(**{"Loan Status": "LIQUIDATION"})],
                     {("1000000001", "Loan Status"): "ACTIVE"},
                     log_rows=[_log("1000000001", "Loan Status",
                                    reason="LENDER REPORTED LIQUIDATION")])
        self.assertIn("LENDER REPORTED LIQUIDATION", rows[0]["Comments"])
        self.assertNotIn(mr.NO_REVIEW_RECORD, rows[0]["Comments"])

    def test_the_reviewer_note_is_never_included(self):
        """Internal free-form text stays internal; this report goes to a third
        party."""
        rows = _rows([_si(**{"Loan Status": "LIQUIDATION"})],
                     {("1000000001", "Loan Status"): "ACTIVE"},
                     log_rows=[_log("1000000001", "Loan Status",
                                    reason="LENDER REPORTED LIQUIDATION",
                                    note="spoke to servicer, confirmed by email")])
        self.assertNotIn("spoke to servicer", rows[0]["Comments"])
        self.assertNotIn("reviewer_note", inspect.getsource(mr.build_mismatch_rows))

    def test_an_absent_value_reads_as_none_in_the_sentence(self):
        rows = _rows([_si(**{"Liquidation Loss": 100.0})],
                     {("1000000001", "Liquidation Loss"): None})
        self.assertIn(f"FBC Determination: {mr.ABSENT_DISPLAY}", rows[0]["Comments"])

    def test_a_missing_review_record_still_emits_the_row_and_reports_the_gap(self):
        gaps = []
        rows = _rows([_si(**{"Loan Status": "LIQUIDATION"})],
                     {("1000000001", "Loan Status"): "ACTIVE"}, log_rows=[], gaps=gaps)
        self.assertEqual(len(rows), 1)
        self.assertIn(mr.NO_REVIEW_RECORD, rows[0]["Comments"])
        self.assertEqual(len(gaps), 1)
        self.assertEqual(gaps[0]["field"], "Loan Status")


# ---------------------------------------------------------------------------
# Deal grouping
# ---------------------------------------------------------------------------

class DealGroupingTests(unittest.TestCase):

    def test_the_two_worksheet_labels(self):
        self.assertEqual(mr.WORKSHEETS, ("SOUP 2024 FBC1", "SOUP 2025 FBC1"))
        self.assertEqual(mr.worksheet_label("SOUP2024_1"), "SOUP 2024 FBC1")
        self.assertEqual(mr.worksheet_label("SOUP2025_1"), "SOUP 2025 FBC1")

    def test_both_worksheets_are_always_present(self):
        blocks = mr.group_by_deal([[]], ["SOUP2024_1"])
        self.assertEqual(sorted(blocks), sorted(mr.WORKSHEETS))

    def test_rows_go_to_their_own_deal_only(self):
        a = _rows([_si()], {("1000000001", "Loan Status"): "DEFERRED"})
        b = _rows([_si(**{"SBA GP Number": "2000000002"})],
                  {("2000000002", "Loan Status"): "DEFERRED"})
        blocks = mr.group_by_deal([a, b], ["SOUP2024_1", "SOUP2025_1"])
        self.assertEqual([r["SBA GP Number"] for r in blocks["SOUP 2024 FBC1"]],
                         ["1000000001"])
        self.assertEqual([r["SBA GP Number"] for r in blocks["SOUP 2025 FBC1"]],
                         ["2000000002"])

    def test_an_unassignable_deal_is_returned_never_guessed_or_misfiled(self):
        for deal in (None, "", "SOUP2023_1", "not a deal"):
            with self.subTest(deal=deal):
                rows = _rows([_si()], {("1000000001", "Loan Status"): "DEFERRED"})
                unassigned = []
                blocks = mr.group_by_deal([rows], [deal], unassigned=unassigned)
                self.assertEqual(blocks["SOUP 2024 FBC1"], [])
                self.assertEqual(blocks["SOUP 2025 FBC1"], [])
                self.assertEqual(len(unassigned), 1)
                self.assertEqual(unassigned[0]["deal"], deal)

    def test_no_third_worksheet_is_ever_created(self):
        unassigned = []
        blocks = mr.group_by_deal(
            [_rows([_si()], {("1000000001", "Loan Status"): "DEFERRED"})],
            ["SOUP2023_1"], unassigned=unassigned)
        self.assertEqual(set(blocks), set(mr.WORKSHEETS))


# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------

def _render(blocks):
    buf = io.BytesIO()
    mr.generate_mismatch_report_bytes(blocks, buf)
    return load_workbook(io.BytesIO(buf.getvalue()))


class WorkbookTests(unittest.TestCase):

    def setUp(self):
        rows = _rows([_si(**{"Loan Status": "LIQUIDATION"})],
                     {("1000000001", "Loan Status"): "ACTIVE"})
        self.blocks = mr.group_by_deal([rows], ["SOUP2025_1"])
        self.wb = _render(self.blocks)

    def test_exactly_two_worksheets_with_the_exact_names(self):
        self.assertEqual(self.wb.sheetnames, ["SOUP 2024 FBC1", "SOUP 2025 FBC1"])

    def test_the_seven_columns_in_order(self):
        ws = self.wb["SOUP 2025 FBC1"]
        self.assertEqual([ws.cell(row=1, column=c).value for c in range(1, 8)],
                         mr.MISMATCH_COLUMNS)

    def test_a_deal_with_no_mismatches_keeps_its_sheet_and_says_so(self):
        ws = self.wb["SOUP 2024 FBC1"]
        self.assertEqual([ws.cell(row=1, column=c).value for c in range(1, 8)],
                         mr.MISMATCH_COLUMNS)
        self.assertEqual(ws.cell(row=2, column=1).value, mr.NO_MISMATCHES_TEXT)

    def test_the_row_is_written_to_the_correct_sheet(self):
        ws = self.wb["SOUP 2025 FBC1"]
        self.assertEqual(ws.cell(row=2, column=1).value, "Lendistry")
        self.assertEqual(ws.cell(row=2, column=6).value, "Loan Status")

    def test_both_sheets_render_when_both_are_empty(self):
        wb = _render(mr.group_by_deal([[]], ["SOUP2024_1"]))
        for label in mr.WORKSHEETS:
            with self.subTest(sheet=label):
                self.assertEqual(wb[label].cell(row=2, column=1).value,
                                 mr.NO_MISMATCHES_TEXT)

    def test_output_is_deterministic_across_repeated_runs(self):
        def cells():
            wb = _render(self.blocks)
            return {n: [[c.value for c in row] for row in wb[n].iter_rows()]
                    for n in wb.sheetnames}
        self.assertEqual(cells(), cells())


class OrderingTests(unittest.TestCase):

    def test_rows_sort_by_gp_then_declared_field_order(self):
        postings = [_si(**{"SBA GP Number": "2000000002", "Loan Status": "LIQUIDATION",
                           "Liquidation Loss": 5.0}),
                    _si(**{"SBA GP Number": "1000000001", "Loan Status": "LIQUIDATION",
                           "Ending Balance": 0.0})]
        overrides = {("2000000002", "Loan Status"): "ACTIVE",
                     ("2000000002", "Liquidation Loss"): None,
                     ("1000000001", "Loan Status"): "ACTIVE",
                     ("1000000001", "Ending Balance"): 500.0}
        rows = _rows(postings, overrides)
        self.assertEqual([(r["SBA GP Number"], r["Mismatch Type"]) for r in rows],
                         [("1000000001", "Loan Status"), ("1000000001", "Ending Balance"),
                          ("2000000002", "Loan Status"), ("2000000002", "Liquidation Loss")])

    def test_servicer_order_is_preserved_within_a_worksheet(self):
        a = _rows([_si()], {("1000000001", "Loan Status"): "DEFERRED"}, soup_key="Zebra Bank")
        b = _rows([_si(**{"SBA GP Number": "2000000002"})],
                  {("2000000002", "Loan Status"): "DEFERRED"}, soup_key="Alpha Bank")
        blocks = mr.group_by_deal([a, b], ["SOUP2025_1", "SOUP2025_1"])
        self.assertEqual([r["Lender"] for r in blocks["SOUP 2025 FBC1"]],
                         ["Zebra Bank", "Alpha Bank"])


# ---------------------------------------------------------------------------
# Purity
# ---------------------------------------------------------------------------

class PurityTests(unittest.TestCase):
    """A reporting consumer, not a calculation engine."""

    def test_the_input_frames_are_never_mutated(self):
        postings = [_si(**{"Loan Status": "LIQUIDATION"})]
        submitted, final = _frames(postings, {("1000000001", "Loan Status"): "ACTIVE"})
        before_s, before_f = submitted.copy(deep=True), final.copy(deep=True)
        mr.build_mismatch_rows(submitted, final, "Lendistry")
        pd.testing.assert_frame_equal(submitted, before_s)
        pd.testing.assert_frame_equal(final, before_f)

    def test_the_override_map_is_never_mutated(self):
        overrides = {("1000000001", "Loan Status"): "ACTIVE"}
        before = dict(overrides)
        _rows([_si(**{"Loan Status": "LIQUIDATION"})], overrides)
        self.assertEqual(overrides, before)

    def test_the_module_imports_only_the_permitted_set(self):
        tree = ast.parse(inspect.getsource(mr))
        imported = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                imported.update(a.name.split(".")[0] for a in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                imported.add(node.module.split(".")[0])
        self.assertEqual(imported, {"decimal", "pandas", "openpyxl", "checks", "report"})

    def test_it_contains_no_business_rule_or_domain_constant(self):
        code = _code_only(mr)
        for forbidden in ("DELINQUENT", "30 DAYS", "60 DAYS", "PAID-IN-FULL",
                          "BANKRUPTCY", "FORECLOSURE", "THRESHOLD", "0.80",
                          "_PERMITTED_LOAN_STATUS", "_classify", "ladder"):
            with self.subTest(symbol=forbidden):
                self.assertNotIn(forbidden, code)

    def test_it_references_no_audit_check(self):
        code = _code_only(mr)
        for check in ("STAT-01", "PORT-01", "BAL-01", "POP-01", "CONT-01", "MEMO-01"):
            with self.subTest(check=check):
                self.assertNotIn(check, code)

    def test_it_touches_no_session_request_or_frontend(self):
        code = _code_only(mr)
        for symbol in ("SESSION_CACHE", "request", "axios", "fetch(", "JSONResponse",
                       "fastapi"):
            with self.subTest(symbol=symbol):
                self.assertNotIn(symbol, code)

    def test_it_calls_no_human_review_or_correction_logic(self):
        code = _code_only(mr)
        for symbol in ("generate_reviewed_workbook", "build_recommendations",
                       "build_eligible_corrections", "classify_recommendation",
                       "validate_approved_value", "consolidate_cycle_postings",
                       "_apply_verified_overrides"):
            with self.subTest(symbol=symbol):
                self.assertNotIn(symbol, code)

    def test_the_field_scope_cannot_drift_from_the_override_map(self):
        self.assertEqual(set(mr.MISMATCH_FIELDS), set(checks._VERIFIED_OVERRIDE_COLUMNS))



class TestableWithoutMainTests(unittest.TestCase):
    """Commit 3's wiring is provably separable from this module."""

    def test_the_module_itself_never_imports_main(self):
        """The production module needs nothing from the endpoint layer, which
        is what makes it testable on its own."""
        tree = ast.parse(inspect.getsource(mr))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                self.assertNotIn("main", [a.name for a in node.names])
            elif isinstance(node, ast.ImportFrom):
                self.assertNotEqual(node.module, "main")

    def test_no_entry_point_takes_a_request_or_session_argument(self):
        for fn in (mr.build_mismatch_rows, mr.group_by_deal,
                   mr.generate_mismatch_report_bytes, mr.mismatch_report_filename):
            params = set(inspect.signature(fn).parameters)
            with self.subTest(fn=fn.__name__):
                self.assertEqual(params & {"request", "session", "session_id", "app"}, set())

    def test_main_uses_only_this_module_s_public_entry_points(self):
        """main.py IS the consumer as of Commit 3.

        This replaces Commit 2's "main does not reference the module yet",
        which pinned that commit's not-yet-wired claim and expired the moment
        the consumer arrived. The invariant worth keeping is the direction of
        the dependency: main reaches in only through the documented entry
        points, so this module's internals stay free to change and the
        separability the tests above assert is not quietly undone by a caller
        binding to a private name."""
        import main
        source = _code_only(main)
        used = set(re.findall(r"mismatch_mod\.(\w+)", source))
        self.assertTrue(used)
        self.assertEqual(used - {"build_mismatch_rows", "group_by_deal",
                                 "generate_mismatch_report_bytes",
                                 "mismatch_report_filename", "WORKSHEET_2024",
                                 "WORKSHEET_2025", "WORKSHEETS",
                                 "MISMATCH_FIELDS", "NO_MISMATCHES_TEXT"}, set())
        self.assertEqual([n for n in used if n.startswith("_")], [])


if __name__ == "__main__":
    unittest.main()
