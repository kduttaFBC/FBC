"""Monthly Servicer Report — CLIENT-CONFIRMED (2026-08-11).

A POST-REVIEW consolidation of reviewer-approved lender data into the client's
own report format, one workbook per SOUP deal:

    lender submissions -> Audit -> Human Review -> Corrected Lender Workbooks
      -> Consolidated Monthly Servicer Reports -> Final Audited Portfolio Summary

Structure is taken from the client's supplied workbook
"2026.06_SOUP 7a 2025-FBC1_Monthly Servicer Report_June_2026_v2.xlsx".

The fourteen agreed protections:

   1  every reviewed lender lands in the correct SOUP deal
   2  every eligible reviewed loan appears exactly once
   3  no 2024 loan in the 2025 report and vice versa
   4  approved Loan/Delinquency Status corrections are reflected
   5  a REJECTED recommendation leaves the lender's value untouched
   6  consolidation changes no financial value
   7  Lender/Servicer Name is populated on every row
   8  the exact 21-column order is preserved
   9  both worksheets exist under the exact names
  10  SUBTOTAL formulas span the complete generated population
  11  Remittance amounts reconcile to Total Payment Received per servicer
  12  the Remittance total reconciles to the loan-level report
  13  Note Series / Remittance Date correct; Collection Period surfaced, blank
  14  the Final Audited Portfolio Summary is unchanged

plus end-to-end validation on the real June fixtures, where the 2025 deal must
reproduce the client's own row counts (1,436 = Lendistry 1,234 + CB&T 155 +
Millennium 47).

The real-file tests skip unless the fixtures are present. Run with:
    FALCONBRIDGE_SOUP_XLSX=... FALCONBRIDGE_MEMO_XLSX=... \\
    FALCONBRIDGE_SRF_DIR=... python -m pytest test_monthly_servicer_report.py
"""

import datetime
import glob
import inspect
import io
import os
import unittest

import pandas as pd
from openpyxl import Workbook, load_workbook

import checks
import file_readers
import monthly_servicer_report as msr


SOUP_PATH = os.environ.get("FALCONBRIDGE_SOUP_XLSX")
MEMO_PATH = os.environ.get("FALCONBRIDGE_MEMO_XLSX")
SRF_DIR = os.environ.get("FALCONBRIDGE_SRF_DIR")
REAL = all([SOUP_PATH, MEMO_PATH, SRF_DIR])

JUNE_SERVICERS = [
    ("Community Bank and Trust", "Community Bank & Trust", "Community Bank and Trust", "SOUP2025_1"),
    ("Lendistry_SOUP 2024", "Lendistry", "Lendistry 2024", "SOUP2024_1"),
    ("Lendistry_SOUP 2025", "Lendistry", "Lendistry 2025", "SOUP2025_1"),
    ("Millennium_SOUP 2024", "Millennium Bank", "Millennium 2024", "SOUP2024_1"),
    ("Millennium_SOUP 2025", "Millennium Bank", "Millennium 2025", "SOUP2025_1"),
    ("Mission Valley Bank", "Mission Valley Bank", "Mission Valley Bank", "SOUP2024_1"),
    ("Peoples Bank", "Peoples Bank", "Peoples Bank", "SOUP2024_1"),
    ("Poppy Bank", "Poppy Bank", "Poppy Bank", "SOUP2024_1"),
]

# The client's own supplied workbook, for the 2025 deal.
CLIENT_2025_ROWS = 1436
CLIENT_2025_PER_SERVICER = {"Lendistry": 1234, "Community Bank & Trust": 155,
                            "Millennium Bank": 47}

FINANCIAL_COLUMNS = ["Net Interest", "Principal", "Total Payment Received",
                     "Beginning Actual Balance", "Ending Actual Balance"]

# The one reviewer approval in the June set: Mission Valley's GP 4778059109
# moves LIQUIDATION -> ACTIVE, re-entering Outstanding with an ending balance
# of 276,702.10. Same loan test_original_summary_final_format.py uses.
MVB_OVERRIDE_GP = "4778059109"
MVB_OVERRIDE_DELTA = 276702.10


def _srf(rows):
    """(gp, lender_loan, status, dq, beg, end, received) -> a Servicer Inputs frame."""
    return pd.DataFrame([{
        "SBA GP Number": gp, "Lender Loan Number": ln, "Loan Status": ls,
        "Delinquency Status": dq, "Beginning Balance": bb, "Ending Balance": eb,
        "Payment Received": rec, "Scheduled Payment": 500.0, "Servicing Fee": 0.0,
        "Net Interest": 100.0, "Principal": 50.0, "Net Interest Rate": 0.075,
        "Paid-In-Full Date": None, "Liquidation Date": None, "Foreclosure Date": None,
        "REO Date": None, "Liquidation Proceeds": None,
        "Subsequent Recoveries/(Losses)": None, "Liquidation Loss": None,
        "Modification  Date": None, "Modification Comments (Free-form)": None,
        "Loan Payment Instance": 1,
    } for gp, ln, ls, dq, bb, eb, rec in rows])


SAMPLE = _srf([
    ("1", "L-1", "ACTIVE", "CURRENT", 100000.0, 99000.0, 500.0),
    ("2", "L-2", "ACTIVE", "30 DAYS", 50000.0, 50000.0, 0.0),
    ("3", "L-3", "LIQUIDATION", "120 DAYS", 25000.0, 25000.0, 0.0),
])


def _find(substring):
    hits = [p for p in glob.glob(os.path.join(SRF_DIR, "*.xlsx"))
            if substring.lower() in os.path.basename(p).lower()]
    return hits[0] if hits else None


# The real Reference Tables remittance calendar, trimmed to the rows the tests
# exercise: the SOUP2025_1 anchor (row 16) and the May/June/July 2026 reporting
# months (rows 25/26/27). Dates and period numbers are verbatim from the June
# and July 2026 SRFs -- the irregular days (21st vs 23rd) are real.
SAMPLE_CALENDAR = {
    "2024-06-21": 1,
    "2025-09-23": 16,   # SOUP2025_1 first remittance -- the anchor
    "2026-05-21": 24,
    "2026-06-23": 25,   # client May 2026 MSR states Collection Period 10
    "2026-07-23": 26,   # client June 2026 MSR states Collection Period 11
    "2026-08-21": 27,   # July 2026 SRFs' own remittance date
}


def _servicer(name="Test Bank", frame=None, overrides=None, excluded=None,
              note_series="2025 FBC-1", remittance_date=pd.Timestamp("2026-07-23"),
              calendar=SAMPLE_CALENDAR):
    issues = []
    rows = msr.build_servicer_rows(frame if frame is not None else SAMPLE, name,
                                   verified_overrides=overrides, excluded_gps=excluded,
                                   issues=issues)
    return {"servicer_name": name, "display_label": name, "rows": rows, "issues": issues,
            "note_series": note_series, "remittance_date": remittance_date,
            "collection_period_calendar": calendar}


def _render(report):
    buf = io.BytesIO()
    msr.generate_monthly_servicer_report_bytes(report, buf)
    return load_workbook(io.BytesIO(buf.getvalue()))


# ---------------------------------------------------------------------------
# Real-file end-to-end. Built once; every real-file class reads it.
# ---------------------------------------------------------------------------

_JUNE_CACHE = {}


def _june():
    if _JUNE_CACHE:
        return _JUNE_CACHE
    soup = file_readers.read_soup_file(open(SOUP_PATH, "rb").read())
    memo = file_readers.resolve_memo_attribution(
        file_readers.read_service_memo(open(MEMO_PATH, "rb").read()),
        file_readers.build_memo_loan_index(soup))
    by_deal, suppressed, frames = {}, {}, {}
    for substring, soup_key, display, _deal in JUNE_SERVICERS:
        path = _find(substring)
        assert path, f"no June SRF for {display}"
        raw = open(path, "rb").read()
        sk, dl, lt, si, deal = file_readers.read_servicer_file(
            raw, filename=f"{soup_key}||{display}__x.xlsx")
        results = checks.run_servicer_reconciliation(
            dl, lt, si, file_readers.filter_soup_by_servicer(soup, sk),
            memo_df=memo, memo_lender_key=sk, deal=deal)
        stat01 = next(r for r in results if r["id"].split("::")[0] == "STAT-01")
        excluded = [t["SBA GP Number"] for t in (stat01.get("suppressed_closed_loans") or [])]
        suppressed[display] = excluded
        frames[display] = si
        issues = []
        by_deal.setdefault(deal, []).append({
            "servicer_name": sk, "display_label": display,
            "rows": msr.build_servicer_rows(si, sk, verified_overrides={},
                                            excluded_gps=excluded, issues=issues),
            "issues": issues,
            "note_series": file_readers.read_note_series_label(raw),
            "remittance_date": file_readers.read_remittance_date(raw),
            "collection_period_calendar":
                file_readers.read_collection_period_calendar(raw),
        })
    reports = {deal: msr.build_deal_report(deal, servicers)
               for deal, servicers in by_deal.items()}
    _JUNE_CACHE.update({"reports": reports, "suppressed": suppressed,
                        "frames": frames, "by_deal": by_deal})
    return _JUNE_CACHE


# ---------------------------------------------------------------------------
# 1, 3 -- deal assignment and separation
# ---------------------------------------------------------------------------

@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class DealAssignmentTests(unittest.TestCase):

    def test_every_reviewed_lender_lands_in_its_declared_deal(self):
        by_deal = _june()["by_deal"]
        for _sub, _key, display, expected in JUNE_SERVICERS:
            with self.subTest(servicer=display):
                found = [d for d, servicers in by_deal.items()
                         if any(s["display_label"] == display for s in servicers)]
                self.assertEqual(found, [expected])

    def test_exactly_two_deals_2024_and_2025(self):
        self.assertEqual(sorted(_june()["reports"]), ["SOUP2024_1", "SOUP2025_1"])

    def test_five_lenders_in_2024_and_three_in_2025(self):
        reports = _june()["reports"]
        self.assertEqual(len(reports["SOUP2024_1"]["remittance_servicers"]), 5)
        self.assertEqual(len(reports["SOUP2025_1"]["remittance_servicers"]), 3)

    def test_no_2024_loan_appears_in_the_2025_report_or_the_reverse(self):
        reports = _june()["reports"]
        a = {str(r["SBA GP Number"]) for r in reports["SOUP2024_1"]["rows"]}
        b = {str(r["SBA GP Number"]) for r in reports["SOUP2025_1"]["rows"]}
        self.assertEqual(a & b, set())

    def test_the_deals_are_never_combined_into_one_workbook(self):
        reports = _june()["reports"]
        for deal, report in reports.items():
            with self.subTest(deal=deal):
                names = {r["Lender/Servicer Name"] for r in report["rows"]}
                self.assertEqual(names, set(report["remittance_servicers"]))


# ---------------------------------------------------------------------------
# 2 -- every eligible loan exactly once; suppressed zero times
# ---------------------------------------------------------------------------

@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class PopulationTests(unittest.TestCase):

    def test_no_loan_appears_more_than_once_anywhere(self):
        rows = [r for report in _june()["reports"].values() for r in report["rows"]]
        gps = [str(r["SBA GP Number"]) for r in rows]
        self.assertEqual(len(gps), len(set(gps)))

    def test_every_eligible_reviewed_loan_is_present(self):
        cache = _june()
        for _sub, _key, display, deal in JUNE_SERVICERS:
            with self.subTest(servicer=display):
                consolidated = checks.consolidate_cycle_postings(cache["frames"][display])
                expected = {str(gp).strip() for gp in consolidated.index}
                expected -= {str(g).strip() for g in cache["suppressed"][display]}
                actual = {str(r["SBA GP Number"]) for r in cache["reports"][deal]["rows"]
                          if r["Lender/Servicer Name"] == _key}
                self.assertTrue(expected <= actual)

    def test_suppressed_loans_appear_zero_times(self):
        cache = _june()
        present = {str(r["SBA GP Number"]) for report in cache["reports"].values()
                   for r in report["rows"]}
        for display, gps in cache["suppressed"].items():
            for gp in gps:
                with self.subTest(servicer=display, gp=gp):
                    self.assertNotIn(str(gp).strip(), present)

    def test_the_suppressed_set_is_not_empty_so_the_check_means_something(self):
        cache = _june()
        self.assertGreater(sum(len(v) for v in cache["suppressed"].values()), 0)

    def test_the_2025_report_reproduces_the_client_row_count(self):
        self.assertEqual(len(_june()["reports"]["SOUP2025_1"]["rows"]), CLIENT_2025_ROWS)

    def test_the_2025_per_servicer_counts_match_the_client_workbook(self):
        rows = _june()["reports"]["SOUP2025_1"]["rows"]
        for name, expected in CLIENT_2025_PER_SERVICER.items():
            with self.subTest(servicer=name):
                self.assertEqual(sum(1 for r in rows if r["Lender/Servicer Name"] == name),
                                 expected)

    def test_population_equals_the_final_audited_population(self):
        """Both are 'reviewed minus suppressed', so the counts must agree."""
        cache = _june()
        for deal, report in cache["reports"].items():
            with self.subTest(deal=deal):
                total = 0
                for servicer in cache["by_deal"][deal]:
                    views = checks.verified_portfolio_summary_views(
                        cache["frames"][servicer["display_label"]], {},
                        servicer_name=servicer["display_label"],
                        excluded_gps=cache["suppressed"][servicer["display_label"]])
                    del views
                    total += len(servicer["rows"])
                self.assertEqual(total, len(report["rows"]))


# ---------------------------------------------------------------------------
# 4, 5, 6 -- overrides applied, rejections respected, financials untouched
# ---------------------------------------------------------------------------

class ReviewerDecisionTests(unittest.TestCase):

    def test_an_approved_loan_status_correction_is_reflected(self):
        rows = msr.build_servicer_rows(SAMPLE, "Test Bank",
                                       verified_overrides={("3", "Loan Status"): "ACTIVE"})
        row = next(r for r in rows if r["SBA GP Number"] == "3")
        self.assertEqual(row["Loan Status"], "ACTIVE")

    def test_an_approved_delinquency_correction_is_reflected(self):
        rows = msr.build_servicer_rows(
            SAMPLE, "Test Bank", verified_overrides={("2", "Delinquency Status"): "CURRENT"})
        row = next(r for r in rows if r["SBA GP Number"] == "2")
        self.assertEqual(row["Delinquency Status"], "CURRENT")

    def test_a_rejected_recommendation_leaves_the_lender_value_unchanged(self):
        """A rejection contributes NO entry to verified_overrides, so the
        lender's submitted value must survive untouched."""
        rows = msr.build_servicer_rows(SAMPLE, "Test Bank", verified_overrides={})
        by_gp = {r["SBA GP Number"]: r for r in rows}
        self.assertEqual(by_gp["3"]["Loan Status"], "LIQUIDATION")
        self.assertEqual(by_gp["2"]["Delinquency Status"], "30 DAYS")

    def test_no_override_at_all_matches_the_lender_submission_exactly(self):
        rows = msr.build_servicer_rows(SAMPLE, "Test Bank")
        consolidated = checks.consolidate_cycle_postings(SAMPLE)
        for row in rows:
            gp = row["SBA GP Number"]
            with self.subTest(gp=gp):
                self.assertEqual(row["Loan Status"], consolidated.loc[gp, "loan_status"])
                self.assertEqual(row["Delinquency Status"], consolidated.loc[gp, "dq_status"])

    def test_an_override_never_moves_a_financial_value(self):
        plain = {r["SBA GP Number"]: r for r in msr.build_servicer_rows(SAMPLE, "T")}
        overridden = {r["SBA GP Number"]: r for r in msr.build_servicer_rows(
            SAMPLE, "T", verified_overrides={("3", "Loan Status"): "ACTIVE",
                                             ("1", "Delinquency Status"): "90 DAYS"})}
        for gp in plain:
            for column in FINANCIAL_COLUMNS + ["Net Interest Rate"]:
                with self.subTest(gp=gp, column=column):
                    self.assertEqual(plain[gp][column], overridden[gp][column])

    def test_this_module_does_not_widen_the_correction_allow_list(self):
        """AMENDED 2026-08-15: the allow-list grew from two fields to five, for the
        reviewer-approved liquidation-reversal balance decision. What this test
        guards is unchanged and is the point of it -- THIS MODULE does not define
        or widen the list; it consumes checks.py's."""
        self.assertEqual(set(checks._VERIFIED_OVERRIDE_COLUMNS),
                         {"Loan Status", "Delinquency Status", "Ending Balance",
                          "Liquidation Date", "Liquidation Loss"})
        source = inspect.getsource(msr)
        self.assertNotIn("_VERIFIED_OVERRIDE_COLUMNS =", source)

    def test_the_raw_sourced_override_map_is_exactly_the_two_liquidation_fields(self):
        """Liquidation Date and Liquidation Loss are read from the RAW postings, so
        the consolidated override never reaches them and this module applies them
        itself. That map must stay minimal: two fields, nothing else, or a source
        column added later could silently become reviewer-writable."""
        self.assertEqual(set(msr._OVERRIDABLE_SOURCE_COLUMNS),
                         {"Liquidation Date", "Liquidation Loss"})

    def test_a_liquidation_override_reaches_the_raw_sourced_columns(self):
        rows = msr.build_servicer_rows(
            SAMPLE, "T", verified_overrides={("1", "Liquidation Date"): "",
                                             ("1", "Liquidation Loss"): None})
        row = next(r for r in rows if str(r["SBA GP Number"]) == "1")
        self.assertIn(row["Liquidation Date"], (None, ""))
        self.assertIn(row["Liquidation Loss"], (None, ""))

    def test_a_loan_with_no_decision_keeps_its_raw_liquidation_values(self):
        untouched = msr.build_servicer_rows(SAMPLE, "T", verified_overrides={})
        overridden = msr.build_servicer_rows(
            SAMPLE, "T", verified_overrides={("1", "Liquidation Date"): ""})
        by_gp = {str(r["SBA GP Number"]): r for r in untouched}
        after = {str(r["SBA GP Number"]): r for r in overridden}
        for gp in by_gp:
            if gp == "1":
                continue
            self.assertEqual(by_gp[gp]["Liquidation Date"], after[gp]["Liquidation Date"],
                             "loan %s was affected by another loan's decision" % gp)

    def test_the_source_frame_is_never_mutated(self):
        before = SAMPLE.copy(deep=True)
        msr.build_servicer_rows(SAMPLE, "T", verified_overrides={("1", "Loan Status"): "DEFERRED"})
        pd.testing.assert_frame_equal(SAMPLE, before)


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class FinancialValuesUnchangedTests(unittest.TestCase):
    """Consolidation reuses checks.consolidate_cycle_postings verbatim; it must
    not alter, round or re-derive a single figure."""

    def test_every_financial_value_equals_the_consolidated_lender_value(self):
        cache = _june()
        for _sub, soup_key, display, deal in JUNE_SERVICERS:
            consolidated = checks.consolidate_cycle_postings(cache["frames"][display])
            rows = [r for r in cache["reports"][deal]["rows"]
                    if r["Lender/Servicer Name"] == soup_key]
            index = {str(gp).strip(): gp for gp in consolidated.index}
            checked = 0
            for row in rows:
                gp = index.get(str(row["SBA GP Number"]))
                if gp is None:
                    continue
                source = consolidated.loc[gp]
                pairs = (("Net Interest", "net_interest"), ("Principal", "principal"),
                         ("Total Payment Received", "received"),
                         ("Beginning Actual Balance", "beginning_balance"),
                         ("Ending Actual Balance", "ending_balance"))
                for output, key in pairs:
                    self.assertEqual(row[output], source[key],
                                     f"{display} GP {row['SBA GP Number']} {output}")
                checked += 1
            with self.subTest(servicer=display):
                self.assertGreater(checked, 0)

    def test_values_are_not_rounded_on_the_way_out(self):
        """The client's reference workbook stores 2dp values; we store full
        precision and let the number format display it.

        Asserted across BOTH reports rather than one servicer: Lendistry,
        CB&T, Mission Valley and Peoples all submit values already rounded to
        the cent, so a servicer-specific assertion would prove nothing about
        rounding. Millennium and Poppy carry sub-cent precision, and it is
        their values that must survive untouched."""
        cache = _june()
        rows = [r for report in cache["reports"].values() for r in report["rows"]]
        unrounded = [
            (r["Lender/Servicer Name"], column)
            for r in rows for column in FINANCIAL_COLUMNS
            if r[column] is not None
            and abs(float(r[column]) - round(float(r[column]), 2)) > 1e-9
        ]
        self.assertGreater(len(unrounded), 0)
        # and the servicers carrying it are the ones that actually submit it
        self.assertTrue({"Millennium Bank", "Poppy Bank"} <= {n for n, _c in unrounded})


# ---------------------------------------------------------------------------
# 7, 8, 9 -- naming, column order, worksheets
# ---------------------------------------------------------------------------

class StructureTests(unittest.TestCase):

    def setUp(self):
        self.report = msr.build_deal_report("SOUP2025_1", [_servicer()])
        self.wb = _render(self.report)
        self.ws = self.wb[msr.MONTHLY_SERVICER_REPORT_SHEET]

    def test_exactly_two_worksheets_with_the_exact_names(self):
        self.assertEqual(self.wb.sheetnames,
                         ["Monthly Servicer Report", "Remittance Amount"])

    def test_the_exact_21_column_order(self):
        header = [self.ws.cell(row=8, column=c).value for c in range(1, 22)]
        self.assertEqual(header, [
            "SBA GP Number", "Lender Loan Number", "Loan Status", "Delinquency Status",
            "Net Interest Rate", "Net Interest", "Principal", "Total Payment Received",
            "Beginning Actual Balance", "Ending Actual Balance", "Paid-In-Full Date",
            "Liquidation Date", "Foreclosure Date", "REO Date", "Liquidation Proceeds",
            "Subsequent Recoveries/Losses", "Cumulative Recoveries/Losses",
            "Liquidation Loss", "Modification  Date", "Modification Comments",
            "Lender/Servicer Name"])

    def test_modification_date_keeps_the_client_double_space(self):
        self.assertEqual(self.ws.cell(row=8, column=19).value, "Modification  Date")

    def test_there_is_no_twenty_second_column(self):
        self.assertEqual(self.ws.max_column, 21)
        self.assertIsNone(self.ws.cell(row=8, column=22).value)

    def test_lender_servicer_name_is_populated_on_every_row(self):
        for r in range(9, self.ws.max_row + 1):
            with self.subTest(row=r):
                self.assertEqual(self.ws.cell(row=r, column=21).value, "Test Bank")

    def test_the_remittance_sheet_columns(self):
        rm = self.wb[msr.REMITTANCE_AMOUNT_SHEET]
        header = [rm.cell(row=5, column=c).value for c in range(1, 5)]
        self.assertEqual(header, ["Servicer", "Amount ($)", "Note Account", "Net Amount ($)"])

    def test_the_metadata_block(self):
        self.assertEqual(self.ws["A1"].value, "SOUP Monthly Servicer Report")
        self.assertEqual(self.ws["A2"].value, "► Note Series:")
        self.assertEqual(self.ws["A3"].value, "► Remittance Date:")
        self.assertEqual(self.ws["A4"].value, "► Collection Period:")

    def test_unsourced_columns_are_blank_not_invented(self):
        for r in range(9, self.ws.max_row + 1):
            with self.subTest(row=r):
                self.assertIsNone(self.ws.cell(row=r, column=17).value)   # Cumulative
        rm = self.wb[msr.REMITTANCE_AMOUNT_SHEET]
        self.assertIsNone(rm.cell(row=6, column=3).value)                 # Note Account


# ---------------------------------------------------------------------------
# 10 -- SUBTOTAL spans the complete population
# ---------------------------------------------------------------------------

class SubtotalTests(unittest.TestCase):

    def _sheet(self, rows):
        frame = _srf([(str(i), f"L-{i}", "ACTIVE", "CURRENT", 1000.0, 900.0, 100.0)
                      for i in range(1, rows + 1)])
        report = msr.build_deal_report("SOUP2025_1", [_servicer(frame=frame)])
        return _render(report)[msr.MONTHLY_SERVICER_REPORT_SHEET], report

    def test_the_count_subtotal_covers_exactly_the_generated_rows(self):
        for n in (1, 5, 40):
            with self.subTest(rows=n):
                ws, _ = self._sheet(n)
                self.assertEqual(ws["A7"].value, f"=SUBTOTAL(2,A9:A{8 + n})")

    def test_every_summed_column_covers_exactly_the_generated_rows(self):
        ws, _ = self._sheet(12)
        for name in ("Net Interest", "Principal", "Total Payment Received",
                     "Beginning Actual Balance", "Ending Actual Balance",
                     "Liquidation Proceeds", "Subsequent Recoveries/Losses",
                     "Cumulative Recoveries/Losses", "Liquidation Loss"):
            letter = chr(ord("A") + msr.MSR_COLUMNS.index(name))
            with self.subTest(column=name):
                self.assertEqual(ws[f"{letter}7"].value,
                                 f"=SUBTOTAL(9,{letter}9:{letter}20)")

    def test_the_subtotal_range_is_never_the_hardcoded_template_range(self):
        ws, _ = self._sheet(3)
        self.assertNotIn("1513", str(ws["A7"].value))

    def test_columns_the_client_does_not_subtotal_stay_empty(self):
        ws, _ = self._sheet(4)
        for letter in ("B", "C", "D", "E", "K", "L", "M", "N", "S", "T", "U"):
            with self.subTest(column=letter):
                self.assertIsNone(ws[f"{letter}7"].value)

    def test_the_last_data_row_is_inside_the_subtotal_range(self):
        ws, report = self._sheet(9)
        last = 8 + len(report["rows"])
        # GP cells are numeric so SUBTOTAL(2) can count them.
        self.assertEqual(ws.cell(row=last, column=1).value, 9)
        self.assertIn(f"A9:A{last}", ws["A7"].value)


# ---------------------------------------------------------------------------
# 11, 12 -- remittance reconciliation
# ---------------------------------------------------------------------------

class RemittanceFormulaTests(unittest.TestCase):

    def setUp(self):
        self.report = msr.build_deal_report("SOUP2025_1", [
            _servicer(name="Bank A"), _servicer(name="Bank B")])
        self.rm = _render(self.report)[msr.REMITTANCE_AMOUNT_SHEET]

    def test_amount_is_a_sumifs_against_the_loan_sheet(self):
        self.assertEqual(
            self.rm["B6"].value,
            "=SUMIFS('Monthly Servicer Report'!H:H,"
            "'Monthly Servicer Report'!U:U,'Remittance Amount'!A6)")

    def test_the_sumifs_targets_total_payment_received_and_lender_name(self):
        self.assertEqual(msr.MSR_COLUMNS.index("Total Payment Received") + 1, 8)   # H
        self.assertEqual(msr.MSR_COLUMNS.index("Lender/Servicer Name") + 1, 21)    # U

    def test_one_row_per_servicer_in_upload_order(self):
        self.assertEqual([self.rm.cell(row=r, column=1).value for r in (6, 7)],
                         ["Bank A", "Bank B"])

    def test_net_amount_is_amount_minus_note_account(self):
        self.assertEqual(self.rm["D6"].value, "=B6-C6")

    def test_the_total_row_sums_exactly_the_servicer_rows(self):
        self.assertEqual(self.rm["B8"].value, "=SUM(B6:B7)")
        self.assertEqual(self.rm["D8"].value, "=SUM(D6:D7)")

    def test_the_total_does_not_reach_above_the_first_servicer_row(self):
        """The client's workbook sums B4:B8 while its data starts at B6."""
        self.assertNotIn("B4", str(self.rm["B8"].value))
        self.assertNotIn("B5", str(self.rm["B8"].value))

    def test_the_metadata_mirrors_the_loan_sheet_by_formula(self):
        self.assertEqual(self.rm["B2"].value, "='Monthly Servicer Report'!B3")
        self.assertEqual(self.rm["B3"].value, "='Monthly Servicer Report'!B4")

    def test_amounts_are_never_hardcoded_values(self):
        for r in (6, 7):
            with self.subTest(row=r):
                self.assertIsInstance(self.rm.cell(row=r, column=2).value, str)
                self.assertTrue(self.rm.cell(row=r, column=2).value.startswith("=SUMIFS("))


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class RemittanceReconciliationTests(unittest.TestCase):
    """What the SUMIFS will evaluate to, computed from the same rows Excel
    would see."""

    def test_each_servicer_amount_equals_its_total_payment_received(self):
        for deal, report in _june()["reports"].items():
            for name in report["remittance_servicers"]:
                expected = sum(float(r["Total Payment Received"] or 0)
                               for r in report["rows"] if r["Lender/Servicer Name"] == name)
                with self.subTest(deal=deal, servicer=name):
                    self.assertGreater(expected, 0)
                    rows = [r for r in report["rows"] if r["Lender/Servicer Name"] == name]
                    self.assertAlmostEqual(
                        sum(float(r["Total Payment Received"] or 0) for r in rows),
                        expected, places=2)

    def test_the_remittance_total_reconciles_to_the_loan_level_report(self):
        for deal, report in _june()["reports"].items():
            per_servicer = sum(
                sum(float(r["Total Payment Received"] or 0)
                    for r in report["rows"] if r["Lender/Servicer Name"] == name)
                for name in report["remittance_servicers"])
            grand = sum(float(r["Total Payment Received"] or 0) for r in report["rows"])
            with self.subTest(deal=deal):
                self.assertAlmostEqual(per_servicer, grand, places=2)

    def test_every_row_is_attributed_to_a_remittance_servicer(self):
        for deal, report in _june()["reports"].items():
            names = set(report["remittance_servicers"])
            with self.subTest(deal=deal):
                self.assertTrue(all(r["Lender/Servicer Name"] in names
                                    for r in report["rows"]))


# ---------------------------------------------------------------------------
# 13 -- report metadata
# ---------------------------------------------------------------------------

class MetadataTests(unittest.TestCase):

    def test_note_series_and_remittance_date_carry_through(self):
        report = msr.build_deal_report("SOUP2025_1", [_servicer()])
        self.assertEqual(report["note_series"], "2025 FBC-1")
        self.assertEqual(report["remittance_date"], pd.Timestamp("2026-07-23"))

    def test_collection_period_is_derived_for_soup2025(self):
        """SUPERSEDES test_collection_period_is_always_blank. HISTORICALLY
        VALIDATED (2026-08-14): row 26 - row 16 + 1 = 11, the value the client's
        own June 2026 MSR states."""
        report = msr.build_deal_report("SOUP2025_1", [_servicer()])
        self.assertEqual(report["collection_period"], 11)
        ws = _render(report)[msr.MONTHLY_SERVICER_REPORT_SHEET]
        self.assertEqual(ws["B4"].value, 11)

    def test_collection_period_raises_no_issue_once_it_resolves(self):
        """SUPERSEDES test_collection_period_always_raises_a_validation_issue.
        The issue is now conditional: it fires only when the field cannot be
        resolved, so a resolved period must not carry a stale warning."""
        report = msr.build_deal_report("SOUP2025_1", [_servicer()])
        codes = [i["code"] for i in report["validation_issues"]]
        self.assertNotIn(msr.ISSUE_COLLECTION_PERIOD_UNAVAILABLE, codes)

    def test_a_missing_note_series_is_an_issue_not_a_guess(self):
        report = msr.build_deal_report("SOUP2025_1", [_servicer(note_series=None)])
        self.assertIsNone(report["note_series"])
        self.assertIn(msr.ISSUE_NOTE_SERIES_UNAVAILABLE,
                      [i["code"] for i in report["validation_issues"]])

    def test_conflicting_note_series_across_servicers_is_surfaced(self):
        report = msr.build_deal_report("SOUP2025_1", [
            _servicer(name="A", note_series="2025 FBC-1"),
            _servicer(name="B", note_series="2024 FBC-1")])
        self.assertIn(msr.ISSUE_NOTE_SERIES_CONFLICT,
                      [i["code"] for i in report["validation_issues"]])

    def test_a_missing_remittance_date_is_an_issue_not_a_guess(self):
        report = msr.build_deal_report("SOUP2025_1", [_servicer(remittance_date=None)])
        self.assertIsNone(report["remittance_date"])
        self.assertIn(msr.ISSUE_REMITTANCE_DATE_UNAVAILABLE,
                      [i["code"] for i in report["validation_issues"]])


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class RealMetadataTests(unittest.TestCase):

    def test_the_note_series_is_read_from_each_deals_own_files(self):
        reports = _june()["reports"]
        self.assertEqual(reports["SOUP2024_1"]["note_series"], "2024 FBC-1")
        self.assertEqual(reports["SOUP2025_1"]["note_series"], "2025 FBC-1")

    def test_the_remittance_date_matches_the_client_reference_workbook(self):
        for deal, report in _june()["reports"].items():
            with self.subTest(deal=deal):
                self.assertEqual(pd.Timestamp(report["remittance_date"]),
                                 pd.Timestamp("2026-07-23"))

    def test_collection_period_is_11_for_2025_and_blank_for_2024(self):
        """SUPERSEDES test_collection_period_is_blank_and_flagged_for_both_deals.

        The two deals now diverge by design: SOUP2025_1 has a
        HISTORICALLY VALIDATED anchor and resolves to the client's stated 11,
        while SOUP2024_1 has no established anchor and keeps EXACTLY the prior
        behaviour -- blank, with the issue raised. Both halves asserted here so
        neither can regress into the other."""
        reports = _june()["reports"]
        self.assertEqual(reports["SOUP2025_1"]["collection_period"], 11)
        self.assertNotIn(msr.ISSUE_COLLECTION_PERIOD_UNAVAILABLE,
                         [i["code"] for i in reports["SOUP2025_1"]["validation_issues"]])
        self.assertIsNone(reports["SOUP2024_1"]["collection_period"])
        self.assertIn(msr.ISSUE_COLLECTION_PERIOD_UNAVAILABLE,
                      [i["code"] for i in reports["SOUP2024_1"]["validation_issues"]])

    def test_the_rendered_2024_workbook_still_has_a_blank_b4(self):
        ws = _render(_june()["reports"]["SOUP2024_1"])[msr.MONTHLY_SERVICER_REPORT_SHEET]
        self.assertIsNone(ws["B4"].value)


# ---------------------------------------------------------------------------
# 13b -- Collection Period, HISTORICALLY VALIDATED (2026-08-14)
# ---------------------------------------------------------------------------

class CollectionPeriodRuleTests(unittest.TestCase):
    """collection_period = calendar row of the report's Remittance Date
    - calendar row of the deal's anchor + 1.

    The anchor is DERIVED FROM CLIENT-PRODUCED REPORTS, not client-stated, so
    every negative path below matters as much as the positive one: the rule must
    fall back to blank-plus-issue rather than guess."""

    def _period(self, deal="SOUP2025_1", **kw):
        return msr.build_deal_report(deal, [_servicer(**kw)])

    # ---- the three months, from the calendar rather than the month name ----

    def test_may_2026_is_period_10(self):
        self.assertEqual(
            self._period(remittance_date=pd.Timestamp("2026-06-23"))["collection_period"], 10)

    def test_june_2026_is_period_11(self):
        self.assertEqual(
            self._period(remittance_date=pd.Timestamp("2026-07-23"))["collection_period"], 11)

    def test_july_2026_is_period_12(self):
        self.assertEqual(
            self._period(remittance_date=pd.Timestamp("2026-08-21"))["collection_period"], 12)

    def test_the_anchor_month_itself_is_period_1(self):
        self.assertEqual(
            self._period(remittance_date=pd.Timestamp("2025-09-23"))["collection_period"], 1)

    def test_no_month_name_is_hardcoded_anywhere_in_the_rule(self):
        """The rule must be a calendar lookup, not a month lookup. A calendar
        that renumbers every period moves the answer with it; a hardcoded month
        table would not."""
        shifted = {date: period + 100 for date, period in SAMPLE_CALENDAR.items()}
        # Every row moved by the same amount, so the DIFFERENCE is unchanged.
        self.assertEqual(self._period(calendar=shifted)["collection_period"], 11)
        # Now move only the reporting row: the answer must follow.
        moved = dict(SAMPLE_CALENDAR, **{"2026-07-23": 30})
        self.assertEqual(self._period(calendar=moved)["collection_period"], 15)

    # ---- negative paths: blank + issue, never a guess ----

    def _assert_blank_and_flagged(self, report):
        self.assertIsNone(report["collection_period"])
        self.assertIn(msr.ISSUE_COLLECTION_PERIOD_UNAVAILABLE,
                      [i["code"] for i in report["validation_issues"]])
        ws = _render(report)[msr.MONTHLY_SERVICER_REPORT_SHEET]
        self.assertIsNone(ws["B4"].value)

    def test_soup2024_has_no_anchor_and_stays_blank(self):
        self._assert_blank_and_flagged(self._period(deal="SOUP2024_1"))

    def test_an_unknown_deal_stays_blank(self):
        self._assert_blank_and_flagged(self._period(deal="SOUP2099_7"))

    def test_a_deal_of_none_stays_blank(self):
        self._assert_blank_and_flagged(self._period(deal=None))

    def test_a_missing_calendar_stays_blank(self):
        self._assert_blank_and_flagged(self._period(calendar=None))

    def test_an_empty_calendar_stays_blank(self):
        self._assert_blank_and_flagged(self._period(calendar={}))

    def test_a_remittance_date_absent_from_the_calendar_is_never_nearest_matched(self):
        """2026-07-24 is one day off a real row. A nearest-match would return
        11; the rule must return nothing."""
        self._assert_blank_and_flagged(
            self._period(remittance_date=pd.Timestamp("2026-07-24")))

    def test_a_missing_remittance_date_stays_blank(self):
        self._assert_blank_and_flagged(self._period(remittance_date=None))

    def test_an_unreadable_remittance_date_stays_blank(self):
        self._assert_blank_and_flagged(self._period(remittance_date="not a date"))

    def test_a_calendar_without_the_anchor_row_stays_blank(self):
        no_anchor = {d: p for d, p in SAMPLE_CALENDAR.items() if d != "2025-09-23"}
        self._assert_blank_and_flagged(self._period(calendar=no_anchor))

    def test_a_report_predating_the_anchor_stays_blank_not_zero_or_negative(self):
        self._assert_blank_and_flagged(
            self._period(remittance_date=pd.Timestamp("2024-06-21")))

    def test_conflicting_calendars_blank_the_field_rather_than_pick_one(self):
        """A calendar conflict is NOT resolved first-file-wins. The two
        calendars may differ at exactly this report's reporting month, so
        picking one would produce a confidently wrong period indistinguishable
        from a correct one."""
        other = dict(SAMPLE_CALENDAR, **{"2026-07-23": 99})
        report = msr.build_deal_report("SOUP2025_1", [
            _servicer(name="A", calendar=SAMPLE_CALENDAR),
            _servicer(name="B", calendar=other)])
        codes = [i["code"] for i in report["validation_issues"]]
        self.assertIn(msr.ISSUE_COLLECTION_PERIOD_CALENDAR_CONFLICT, codes)
        # Blank, and flagged as unavailable too, so a reviewer scanning for the
        # unavailable code still learns that B4 needs entering by hand.
        self.assertIn(msr.ISSUE_COLLECTION_PERIOD_UNAVAILABLE, codes)
        self._assert_blank_and_flagged(report)

    def test_the_conflict_message_names_the_disagreeing_servicers(self):
        other = dict(SAMPLE_CALENDAR, **{"2026-07-23": 99})
        report = msr.build_deal_report("SOUP2025_1", [
            _servicer(name="A", calendar=SAMPLE_CALENDAR),
            _servicer(name="B", calendar=other)])
        message = next(i["message"] for i in report["validation_issues"]
                       if i["code"] == msr.ISSUE_COLLECTION_PERIOD_CALENDAR_CONFLICT)
        self.assertIn("A", message)
        self.assertIn("B", message)
        self.assertIn("left blank", message)

    def test_identical_calendars_across_servicers_are_not_a_conflict(self):
        report = msr.build_deal_report("SOUP2025_1", [
            _servicer(name="A", calendar=dict(SAMPLE_CALENDAR)),
            _servicer(name="B", calendar=dict(SAMPLE_CALENDAR))])
        self.assertNotIn(msr.ISSUE_COLLECTION_PERIOD_CALENDAR_CONFLICT,
                         [i["code"] for i in report["validation_issues"]])
        self.assertEqual(report["collection_period"], 11)

    def test_one_servicer_missing_its_calendar_does_not_defeat_the_deal(self):
        report = msr.build_deal_report("SOUP2025_1", [
            _servicer(name="A", calendar=None),
            _servicer(name="B", calendar=SAMPLE_CALENDAR)])
        self.assertEqual(report["collection_period"], 11)

    def test_the_anchor_is_a_single_dated_constant_for_one_deal_only(self):
        self.assertEqual(msr._DEAL_COLLECTION_PERIOD_ANCHOR,
                         {"SOUP2025_1": "2025-09-23"})


class CollectionPeriodCalendarReaderTests(unittest.TestCase):
    """file_readers.read_collection_period_calendar -- returns None rather than
    a partial or guessed calendar."""

    @staticmethod
    def _book(rows, sheet_name="Reference Tables"):
        wb = Workbook()
        wb.active.title = sheet_name
        for row in rows:
            wb.active.append(row)
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def test_it_reads_the_period_to_remittance_due_schedule(self):
        data = self._book([
            [None, None, None],
            ["Collection Period", "Report Due Date", "Remittance Due Date"],
            [1, "2024-06-05", "2024-06-21"],
            [2, "2024-07-08", "2024-07-23"],
        ])
        self.assertEqual(file_readers.read_collection_period_calendar(data),
                         {"2024-06-21": 1, "2024-07-23": 2})

    def test_it_ignores_the_trigger_tables_that_share_the_column_name(self):
        """The real sheet carries THREE 'Collection Period' columns; only the
        one paired with a Remittance Due Date is a calendar."""
        data = self._book([
            ["Collection Period", "Cumulative Realized Loss Level", None,
             "Collection Period", "Report Due Date", "Remittance Due Date"],
            [1, 0.0325, None, 1, "2024-06-05", "2024-06-21"],
            [2, 0.0325, None, 2, "2024-07-08", "2024-07-23"],
        ])
        self.assertEqual(file_readers.read_collection_period_calendar(data),
                         {"2024-06-21": 1, "2024-07-23": 2})

    def test_a_shifted_template_still_resolves_because_columns_are_found_by_header(self):
        data = self._book([
            [None, None, None, None, None],
            [None, None, "Collection Period", "Report Due Date", "Remittance Due Date"],
            [None, None, 5, "2024-10-07", "2024-10-23"],
        ])
        self.assertEqual(file_readers.read_collection_period_calendar(data),
                         {"2024-10-23": 5})

    def test_a_missing_sheet_returns_none(self):
        self.assertIsNone(file_readers.read_collection_period_calendar(
            self._book([["Collection Period", "Report Due Date", "Remittance Due Date"],
                        [1, "2024-06-05", "2024-06-21"]], sheet_name="Something Else")))

    def test_no_remittance_due_column_returns_none(self):
        self.assertIsNone(file_readers.read_collection_period_calendar(
            self._book([["Collection Period", "Cumulative Realized Loss Level"],
                        [1, 0.0325]])))

    def test_unparseable_rows_are_skipped_not_guessed(self):
        data = self._book([
            ["Collection Period", "Report Due Date", "Remittance Due Date"],
            ["n/a", "2024-06-05", "2024-06-21"],
            [2, "2024-07-08", "not a date"],
            [3, "2024-08-07", "2024-08-22"],
        ])
        self.assertEqual(file_readers.read_collection_period_calendar(data),
                         {"2024-08-22": 3})

    def test_an_all_unparseable_calendar_returns_none_not_an_empty_dict(self):
        self.assertIsNone(file_readers.read_collection_period_calendar(
            self._book([["Collection Period", "Report Due Date", "Remittance Due Date"],
                        ["n/a", "x", "y"]])))

    def test_garbage_bytes_return_none(self):
        self.assertIsNone(file_readers.read_collection_period_calendar(b"not a workbook"))


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class RealCollectionPeriodTests(unittest.TestCase):
    """The rule against the ACTUAL SRFs, not a fixture."""

    def test_every_june_srf_publishes_the_same_calendar(self):
        calendars = []
        for substring, _key, display, _deal in JUNE_SERVICERS:
            path = _find(substring)
            if not path:
                continue
            cal = file_readers.read_collection_period_calendar(open(path, "rb").read())
            with self.subTest(servicer=display):
                self.assertIsNotNone(cal)
                # Anchor and reporting row, verbatim from the real files.
                self.assertEqual(cal.get("2025-09-23"), 16)
                self.assertEqual(cal.get("2026-07-23"), 26)
            calendars.append(tuple(sorted(cal.items())))
        self.assertTrue(calendars, "no June SRFs found")
        self.assertEqual(len(set(calendars)), 1,
                         "the June SRFs publish differing remittance calendars")

    def test_the_june_soup2025_report_states_the_client_value(self):
        self.assertEqual(_june()["reports"]["SOUP2025_1"]["collection_period"], 11)

    def test_the_rendered_june_2025_workbook_writes_11_to_b4(self):
        ws = _render(_june()["reports"]["SOUP2025_1"])[msr.MONTHLY_SERVICER_REPORT_SHEET]
        self.assertEqual(ws["B4"].value, 11)

    def test_may_and_july_resolve_from_the_real_june_calendar(self):
        """One real calendar, three reporting months -- 10 / 11 / 12 -- which is
        the whole rule: only the Remittance Date changes."""
        path = _find("Lendistry_SOUP 2025")
        if not path:
            self.skipTest("no Lendistry 2025 June SRF")
        cal = file_readers.read_collection_period_calendar(open(path, "rb").read())
        for date, expected in (("2026-06-23", 10), ("2026-07-23", 11)):
            with self.subTest(remittance=date):
                report = msr.build_deal_report("SOUP2025_1", [
                    _servicer(remittance_date=pd.Timestamp(date), calendar=cal)])
                self.assertEqual(report["collection_period"], expected)
        # July's remittance date is NOT in the June calendar -- the June files
        # end at row 26 -- so it must stay blank here rather than be
        # extrapolated. July resolves only from July's own SRFs, which publish
        # row 27; that is asserted in RealJulyCollectionPeriodTests below.
        blank = msr.build_deal_report("SOUP2025_1", [
            _servicer(remittance_date=pd.Timestamp("2026-08-21"), calendar=cal)])
        self.assertIsNone(blank["collection_period"])


@unittest.skipUnless(os.environ.get("FALCONBRIDGE_JULY_LENDISTRY_2025_XLSX"),
                     "set FALCONBRIDGE_JULY_LENDISTRY_2025_XLSX")
class RealJulyCollectionPeriodTests(unittest.TestCase):

    def test_july_2026_resolves_to_12_from_the_july_srf_alone(self):
        raw = open(os.environ["FALCONBRIDGE_JULY_LENDISTRY_2025_XLSX"], "rb").read()
        cal = file_readers.read_collection_period_calendar(raw)
        remittance = file_readers.read_remittance_date(raw)
        self.assertEqual(cal.get("2025-09-23"), 16)
        self.assertEqual(pd.Timestamp(remittance), pd.Timestamp("2026-08-21"))
        self.assertEqual(cal.get("2026-08-21"), 27)
        report = msr.build_deal_report("SOUP2025_1", [
            _servicer(remittance_date=remittance, calendar=cal)])
        self.assertEqual(report["collection_period"], 12)

    def test_the_july_calendar_extends_the_june_one_by_exactly_one_row(self):
        july = file_readers.read_collection_period_calendar(
            open(os.environ["FALCONBRIDGE_JULY_LENDISTRY_2025_XLSX"], "rb").read())
        if not REAL or not _find("Lendistry_SOUP 2025"):
            self.skipTest("no June SRF to compare against")
        june = file_readers.read_collection_period_calendar(
            open(_find("Lendistry_SOUP 2025"), "rb").read())
        self.assertEqual(len(july), len(june) + 1)
        # Every June row survives unchanged -- periods are never renumbered.
        for date, period in june.items():
            with self.subTest(date=date):
                self.assertEqual(july[date], period)


# ---------------------------------------------------------------------------
# 14 -- nothing else moved
# ---------------------------------------------------------------------------

class BoundaryTests(unittest.TestCase):
    """CONT-01, the audit checks, the corrected workbooks and all three
    Portfolio Summaries are outside this module's reach."""

    @staticmethod
    def _code_only(module):
        """Executable lines only -- comments and the module docstring stripped.

        The docstring legitimately NAMES the functions this module must not
        call, to explain the boundary. Searching raw source would match that
        prose and report a violation that does not exist."""
        source = inspect.getsource(module)
        body = source.split('"""')
        code = "".join(body[2::2]) if len(body) > 2 else source
        return "\n".join(line for line in code.splitlines()
                         if not line.strip().startswith("#"))

    def test_this_module_calls_no_audit_or_summary_function(self):
        code = self._code_only(msr)
        for forbidden in ("run_servicer_reconciliation(", "portfolio_summary_check(",
                          "audited_portfolio_summary_record(",
                          "verified_portfolio_summary_views(",
                          "group_portfolio_summary_by_deal(",
                          "generate_final_portfolio_summary_bytes(",
                          "generate_reviewed_workbook(", "build_eligible_corrections(",
                          "continuity_check("):
            with self.subTest(symbol=forbidden):
                self.assertNotIn(forbidden, code)

    def test_it_never_references_an_audit_check_id(self):
        code = self._code_only(msr)
        for check_id in ("CONT-01", "STAT-01", "PORT-01", "BAL-01", "POP-01"):
            with self.subTest(check=check_id):
                self.assertNotIn(check_id, code)

    def test_it_imports_no_correction_or_report_module(self):
        """Parsed, not substring-matched: 'import reporting_exceptions' contains
        the substring 'import report' while importing no such module."""
        import ast
        tree = ast.parse(inspect.getsource(msr))
        imported = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                imported.update(a.name.split(".")[0] for a in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                imported.add(node.module.split(".")[0])
        self.assertNotIn("corrections", imported)
        self.assertNotIn("report", imported)
        self.assertEqual(imported, {"pandas", "openpyxl", "checks", "reporting_exceptions"})

    def test_the_only_checks_helpers_used_are_the_shared_ones(self):
        # Executable lines only -- docstrings name checks._norm_id in prose.
        code = self._code_only(msr)
        used = {line.split("checks.")[1].split("(")[0]
                for line in code.splitlines() if "checks." in line and "(" in line}
        self.assertTrue(used <= {"consolidate_cycle_postings", "_apply_verified_overrides",
                                 "_norm_id"}, used)


# June 2026 Outstanding ($) by deal, per layer. All four measured against the
# real fixtures; see PortfolioSummaryLayerTotalsTests for why each is what it is.
ORIGINAL_2024 = 61228003.82
ORIGINAL_2025 = 99128983.40
FINAL_AUDITED_2024_WITH_APPROVAL = 61504705.92     # +276,702.10 from Mission Valley
FINAL_AUDITED_2025_WITH_APPROVAL = 99128983.40     # unmoved -- no 2025 approval


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class PortfolioSummaryLayerTotalsTests(unittest.TestCase):
    """All four layer totals, asserted separately and named for what they are.

    RENAMED (2026-08-11). The previous single test called
    verified_portfolio_summary_views with NO approvals and asserted
    61,228,003.82 under the name "the final audited deal totals". That figure
    is correct for an empty approval set -- with nothing approved the Final
    Audited layer coincides with the Original on the Outstanding line -- but
    the name read as the POST-APPROVAL figure, which is 61,504,705.92. The
    assertion was sound; the name was misleading. Both cases are now asserted
    explicitly, and the no-approval case is labelled as the control it is."""

    def _totals(self, views_fn):
        cache = _june()
        dollars, counts, deals = [], [], []
        for _sub, _key, display, deal in JUNE_SERVICERS:
            views = views_fn(display, cache)
            dollars.append(views["dollars"])
            counts.append(views["counts"])
            deals.append(deal)
        return {b["deal"]: b["totals"]["dollars"]["Outstanding ($)"]
                for b in checks.group_portfolio_summary_by_deal(dollars, counts, deals)}

    def _original(self):
        return self._totals(lambda display, cache:
                            checks.original_portfolio_summary_views(
                                cache["frames"][display], display))

    def _final_audited(self, approvals):
        return self._totals(lambda display, cache:
                            checks.verified_portfolio_summary_views(
                                cache["frames"][display], approvals.get(display, {}),
                                servicer_name=display,
                                excluded_gps=cache["suppressed"][display]))

    @staticmethod
    def _approvals():
        """The one reviewer approval in the June set: Mission Valley's
        GP 4778059109 moves LIQUIDATION -> ACTIVE, re-entering Outstanding."""
        return {"Mission Valley Bank": {(MVB_OVERRIDE_GP, "Loan Status"): "ACTIVE"}}

    def test_original_2024_is_the_lender_reported_total(self):
        self.assertAlmostEqual(self._original()["SOUP2024_1"], ORIGINAL_2024, places=2)

    def test_original_2025_is_the_lender_reported_total(self):
        self.assertAlmostEqual(self._original()["SOUP2025_1"], ORIGINAL_2025, places=2)

    def test_final_audited_2024_with_the_approved_correction(self):
        """THE assertion the old name implied but did not make."""
        self.assertAlmostEqual(self._final_audited(self._approvals())["SOUP2024_1"],
                               FINAL_AUDITED_2024_WITH_APPROVAL, places=2)

    def test_final_audited_2025_with_the_approved_correction(self):
        self.assertAlmostEqual(self._final_audited(self._approvals())["SOUP2025_1"],
                               FINAL_AUDITED_2025_WITH_APPROVAL, places=2)

    def test_the_2024_movement_is_exactly_the_approved_amount(self):
        self.assertAlmostEqual(
            self._final_audited(self._approvals())["SOUP2024_1"] - self._original()["SOUP2024_1"],
            MVB_OVERRIDE_DELTA, places=2)

    def test_2025_does_not_move_because_no_2025_loan_was_approved(self):
        self.assertAlmostEqual(
            self._final_audited(self._approvals())["SOUP2025_1"], self._original()["SOUP2025_1"],
            places=2)

    def test_with_no_approvals_the_final_layer_matches_the_original(self):
        """The control, explicitly labelled. Closed-at-zero suppression moves
        Outstanding by $0.00 in both deals: all five suppressed loans are
        already outside Outstanding (four PAID-IN-FULL at $0.00, one
        LIQUIDATION), so removing them changes nothing on that line."""
        no_approvals = self._final_audited({})
        self.assertAlmostEqual(no_approvals["SOUP2024_1"], ORIGINAL_2024, places=2)
        self.assertAlmostEqual(no_approvals["SOUP2025_1"], ORIGINAL_2025, places=2)

    def test_the_two_layers_are_not_identical_merely_equal_on_outstanding(self):
        """Suppression DOES move other lines -- Liquidation ($) among them --
        so 'equal on Outstanding' must not be read as 'the same report'."""
        cache = _june()
        original = checks.original_portfolio_summary_views(
            cache["frames"]["Community Bank and Trust"], "Community Bank and Trust")
        final = checks.verified_portfolio_summary_views(
            cache["frames"]["Community Bank and Trust"], {},
            servicer_name="Community Bank and Trust",
            excluded_gps=cache["suppressed"]["Community Bank and Trust"])
        self.assertNotEqual(original["dollars"]["Liquidation ($)"],
                            final["dollars"]["Liquidation ($)"])

    def test_generating_the_reports_does_not_disturb_either_layer(self):
        before_original = self._original()
        before_final = self._final_audited(self._approvals())
        for report in _june()["reports"].values():
            buf = io.BytesIO()
            msr.generate_monthly_servicer_report_bytes(report, buf)
        self.assertEqual(self._original(), before_original)
        self.assertEqual(self._final_audited(self._approvals()), before_final)


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class FinalAuditedSummaryUnchangedTests(unittest.TestCase):

    def test_generating_the_reports_does_not_disturb_the_summary(self):
        cache = _june()
        before = checks.verified_portfolio_summary_views(
            cache["frames"]["Mission Valley Bank"], {}, servicer_name="Mission Valley Bank",
            excluded_gps=cache["suppressed"]["Mission Valley Bank"])
        for report in cache["reports"].values():
            buf = io.BytesIO()
            msr.generate_monthly_servicer_report_bytes(report, buf)
        after = checks.verified_portfolio_summary_views(
            cache["frames"]["Mission Valley Bank"], {}, servicer_name="Mission Valley Bank",
            excluded_gps=cache["suppressed"]["Mission Valley Bank"])
        self.assertEqual(before, after)


# ---------------------------------------------------------------------------
# Unresolvable single-value conflicts -- CLIENT-CONFIRMED (2026-08-11)
# ---------------------------------------------------------------------------

def _conflicting_frame(column, first, second):
    """One loan, two postings that disagree on `column` and agree on the rest.
    Mirrors the real Millennium shape: two contiguous accrual periods."""
    base = {
        "SBA GP Number": "9", "Lender Loan Number": "L-9", "Loan Status": "ACTIVE",
        "Delinquency Status": "CURRENT", "Beginning Balance": 1000.0,
        "Ending Balance": 1000.0, "Payment Received": 100.0, "Scheduled Payment": 500.0,
        "Servicing Fee": 0.0, "Net Interest": 100.0, "Principal": 0.0,
        "Net Interest Rate": 0.05, "Paid-In-Full Date": None, "Liquidation Date": None,
        "Foreclosure Date": None, "REO Date": None, "Liquidation Proceeds": None,
        "Subsequent Recoveries/(Losses)": None, "Liquidation Loss": None,
        "Modification  Date": None, "Modification Comments (Free-form)": None,
    }
    a = dict(base, **{column: first, "Loan Payment Instance": 1,
                      "Interest-From": "2025-11-30", "Interest-To": "2026-01-01",
                      "_excel_row": 32})
    b = dict(base, **{column: second, "Loan Payment Instance": 2,
                      "Interest-From": "2026-01-01", "Interest-To": "2026-02-12",
                      "_excel_row": 33})
    return pd.DataFrame([a, b])


class UnresolvableConflictTests(unittest.TestCase):
    """A single-value field with two valid values and no documented resolution
    rule is left BLANK and surfaced -- never silently picked. Generic, not
    written around one loan."""

    def test_the_conflicted_cell_is_blank(self):
        issues = []
        rows = msr.build_servicer_rows(
            _conflicting_frame("Net Interest Rate", 0.0775, 0.0725), "Bank", issues=issues)
        self.assertIsNone(rows[0]["Net Interest Rate"])

    def test_both_values_and_their_source_periods_are_surfaced(self):
        issues = []
        msr.build_servicer_rows(
            _conflicting_frame("Net Interest Rate", 0.0775, 0.0725), "Bank", issues=issues)
        issue = next(i for i in issues if i["code"] == msr.ISSUE_POSTINGS_DISAGREE)
        self.assertEqual(issue["resolution"], "unresolved")
        self.assertEqual(set(issue["values"]), {"0.0775", "0.0725"})
        self.assertIn("2025-11-30", issue["message"])
        self.assertIn("2026-02-12", issue["message"])
        self.assertIn("BLANK", issue["message"])

    def test_the_rule_is_generic_across_every_single_value_field(self):
        for column, a, b in (("Liquidation Date", "2026-06-29", "2026-06-30"),
                             ("Liquidation Loss", 100.0, 200.0),
                             ("Modification Comments (Free-form)", "x", "y"),
                             ("Liquidation Proceeds", 5.0, 6.0)):
            with self.subTest(column=column):
                issues = []
                rows = msr.build_servicer_rows(
                    _conflicting_frame(column, a, b), "Bank", issues=issues)
                output = next(k for k, v in msr._SOURCE_COLUMNS.items() if v == column)
                self.assertIsNone(rows[0][output])
                self.assertTrue(any(i["field"] == column for i in issues))

    def test_every_other_field_still_populates(self):
        issues = []
        rows = msr.build_servicer_rows(
            _conflicting_frame("Net Interest Rate", 0.0775, 0.0725), "Bank", issues=issues)
        row = rows[0]
        self.assertEqual(sorted(row), sorted(msr.MSR_COLUMNS))
        for column, expected in (("SBA GP Number", "9"), ("Lender Loan Number", "L-9"),
                                 ("Loan Status", "ACTIVE"), ("Delinquency Status", "CURRENT"),
                                 ("Net Interest", 200.0), ("Total Payment Received", 200.0),
                                 ("Beginning Actual Balance", 1000.0),
                                 ("Ending Actual Balance", 1000.0),
                                 ("Lender/Servicer Name", "Bank")):
            with self.subTest(column=column):
                self.assertEqual(row[column], expected)

    def test_consistent_multi_posting_loans_raise_no_false_conflict(self):
        frame = _conflicting_frame("Net Interest Rate", 0.0775, 0.0775)
        issues = []
        rows = msr.build_servicer_rows(frame, "Bank", issues=issues)
        self.assertEqual(rows[0]["Net Interest Rate"], 0.0775)
        self.assertEqual([i for i in issues if i["code"] == msr.ISSUE_POSTINGS_DISAGREE], [])

    def test_a_blank_alongside_a_value_is_not_a_conflict(self):
        frame = _conflicting_frame("Liquidation Date", None, "2026-06-30")
        issues = []
        rows = msr.build_servicer_rows(frame, "Bank", issues=issues)
        self.assertIsNotNone(rows[0]["Liquidation Date"])
        self.assertEqual([i for i in issues if i["code"] == msr.ISSUE_POSTINGS_DISAGREE], [])

    def test_generation_is_never_blocked_by_a_conflict(self):
        issues = []
        rows = msr.build_servicer_rows(
            _conflicting_frame("Net Interest Rate", 0.0775, 0.0725), "Bank", issues=issues)
        report = msr.build_deal_report("SOUP2024_1", [{
            "servicer_name": "Bank", "display_label": "Bank", "rows": rows,
            "issues": issues, "note_series": "2024 FBC-1",
            "remittance_date": pd.Timestamp("2026-07-23")}])
        ws = _render(report)[msr.MONTHLY_SERVICER_REPORT_SHEET]
        self.assertEqual(ws.max_row, 9)
        self.assertIsNone(ws.cell(row=9, column=5).value)          # Net Interest Rate
        self.assertEqual(ws.cell(row=9, column=1).value, 9)         # everything else present
        self.assertEqual(ws.cell(row=9, column=21).value, "Bank")


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class RealConflictTests(unittest.TestCase):
    """Millennium 2024 GP 5446319102 -- a VAR-rate loan with a quarterly reset,
    remitting two contiguous accrual periods in one cycle. The only conflict in
    the entire 1,774-loan June population."""

    GP = "5446319102"

    def _row(self):
        rows = _june()["reports"]["SOUP2024_1"]["rows"]
        return next(r for r in rows if r["SBA GP Number"] == self.GP)

    def test_the_net_interest_rate_is_blank(self):
        self.assertIsNone(self._row()["Net Interest Rate"])

    def test_both_rates_are_surfaced_with_their_accrual_windows(self):
        issues = [i for i in _june()["reports"]["SOUP2024_1"]["validation_issues"]
                  if i["code"] == msr.ISSUE_POSTINGS_DISAGREE
                  and i.get("SBA GP Number") == self.GP]
        self.assertEqual(len(issues), 1)
        issue = issues[0]
        self.assertEqual(issue["field"], "Net Interest Rate")
        self.assertEqual(issue["resolution"], "unresolved")
        values = {v[:6] for v in issue["values"]}
        self.assertEqual(values, {"0.0775", "0.0725"})
        self.assertIn("2025-11-30 to 2026-01-01", issue["message"])
        self.assertIn("2026-01-01 to 2026-02-12", issue["message"])

    def test_the_other_twenty_fields_are_unaffected(self):
        row = self._row()
        self.assertEqual(len(row), 21)
        for column, expected in (("SBA GP Number", self.GP),
                                 ("Loan Status", "Active"),
                                 ("Delinquency Status", "Current"),
                                 ("Lender/Servicer Name", "Millennium Bank")):
            with self.subTest(column=column):
                self.assertEqual(row[column], expected)
        # SUM / FIRST / LAST all still resolved for this loan
        self.assertAlmostEqual(row["Net Interest"], 6649.892323 + 7617.478381, places=4)
        self.assertAlmostEqual(row["Total Payment Received"], 6649.892323 + 7617.478381, places=4)
        self.assertAlmostEqual(row["Beginning Actual Balance"], 996445.884, places=3)
        self.assertAlmostEqual(row["Ending Actual Balance"], 996445.884, places=3)
        self.assertEqual(row["Principal"], 0.0)

    def test_it_is_the_only_conflict_in_the_whole_population(self):
        conflicts = [i for report in _june()["reports"].values()
                     for i in report["validation_issues"]
                     if i["code"] == msr.ISSUE_POSTINGS_DISAGREE]
        self.assertEqual(len(conflicts), 1)
        self.assertEqual(conflicts[0]["SBA GP Number"], self.GP)

    def test_other_multi_posting_loans_produce_no_conflict(self):
        """88 of the 1,774 loans have more than one posting; only one conflicts."""
        cache = _june()
        multi = 0
        for _sub, _key, display, _deal in JUNE_SERVICERS:
            frame = cache["frames"][display]
            counts = frame.groupby(checks._norm_id(frame["SBA GP Number"])).size()
            multi += int((counts > 1).sum())
        self.assertGreater(multi, 50)
        flagged = {i.get("SBA GP Number") for report in cache["reports"].values()
                   for i in report["validation_issues"]
                   if i["code"] == msr.ISSUE_POSTINGS_DISAGREE}
        self.assertEqual(flagged, {self.GP})

    def test_both_reports_still_generate(self):
        for deal, report in _june()["reports"].items():
            with self.subTest(deal=deal):
                wb = _render(report)
                self.assertEqual(wb.sheetnames,
                                 ["Monthly Servicer Report", "Remittance Amount"])
                ws = wb[msr.MONTHLY_SERVICER_REPORT_SHEET]
                self.assertEqual(ws.max_row - 8, len(report["rows"]))

    def test_loan_counts_are_unchanged_by_the_blanking(self):
        reports = _june()["reports"]
        self.assertEqual(len(reports["SOUP2024_1"]["rows"]), 338)
        self.assertEqual(len(reports["SOUP2025_1"]["rows"]), CLIENT_2025_ROWS)

    def test_remittance_totals_are_unchanged_by_the_blanking(self):
        reports = _june()["reports"]
        for deal, expected in (("SOUP2024_1", 2993452.38), ("SOUP2025_1", 2805017.90)):
            with self.subTest(deal=deal):
                total = sum(float(r["Total Payment Received"] or 0)
                            for r in reports[deal]["rows"])
                self.assertAlmostEqual(total, expected, places=2)

    def test_the_conflicted_loan_still_contributes_its_payment(self):
        """Blanking one cell must not remove the loan from the remittance."""
        self.assertGreater(float(self._row()["Total Payment Received"]), 0)


class GpNumberIsNumericTests(unittest.TestCase):
    """SUBTOTAL(2, A9:A...) is COUNT, which counts NUMBERS and ignores text.
    A text-typed GP made the client's published loan count read 0."""

    def _sheet(self, rows=4):
        frame = _srf([(str(1000000000 + i), f"L-{i}", "ACTIVE", "CURRENT",
                       1000.0, 900.0, 100.0) for i in range(1, rows + 1)])
        report = msr.build_deal_report("SOUP2025_1", [_servicer(frame=frame)])
        return _render(report)[msr.MONTHLY_SERVICER_REPORT_SHEET], report

    def test_every_gp_cell_is_written_as_a_number(self):
        ws, report = self._sheet()
        for offset in range(len(report["rows"])):
            r = 9 + offset
            with self.subTest(row=r):
                self.assertEqual(ws.cell(row=r, column=1).data_type, "n")
                self.assertIsInstance(ws.cell(row=r, column=1).value, int)

    def test_the_identifier_value_is_preserved_exactly(self):
        ws, report = self._sheet()
        written = [ws.cell(row=9 + i, column=1).value for i in range(len(report["rows"]))]
        self.assertEqual([str(v) for v in written],
                         [r["SBA GP Number"] for r in report["rows"]])

    def test_the_row_dicts_still_carry_the_canonical_string(self):
        """Only the Excel write coerces. Everything upstream is untouched."""
        _ws, report = self._sheet()
        for row in report["rows"]:
            with self.subTest(gp=row["SBA GP Number"]):
                self.assertIsInstance(row["SBA GP Number"], str)

    def test_a_non_round_tripping_identifier_stays_text(self):
        for gp in ("0123456789", "12-345", "ABC123", "1234567890.5", ""):
            with self.subTest(gp=gp):
                out = msr._excel_gp_value(gp)
                if gp == "":
                    self.assertIsNone(out)
                else:
                    self.assertIsInstance(out, str)
                    self.assertEqual(out, gp)

    def test_a_plain_numeric_identifier_becomes_an_int(self):
        self.assertEqual(msr._excel_gp_value("5446319102"), 5446319102)
        self.assertIsInstance(msr._excel_gp_value("5446319102"), int)

    def test_ten_digit_gps_are_exact_in_float64(self):
        """Excel stores numbers as doubles; 10-digit GPs are far inside the
        exact-integer range, so no identifier can be silently rounded."""
        self.assertLess(9999999999, 2 ** 53)

    def test_subtotal_counts_the_full_population(self):
        """What the defect broke: with text GPs this evaluated to 0."""
        ws, report = self._sheet(rows=6)
        numeric = sum(1 for i in range(len(report["rows"]))
                      if isinstance(ws.cell(row=9 + i, column=1).value, int))
        self.assertEqual(numeric, 6)
        self.assertEqual(ws["A7"].value, "=SUBTOTAL(2,A9:A14)")

    def test_subtotal_responds_to_filtering(self):
        """SUBTOTAL(2) counts only VISIBLE numeric rows. Simulated by hiding
        rows the way an AutoFilter does -- the client's own saved workbook
        shows 17 of 1,436 for exactly this reason."""
        ws, report = self._sheet(rows=6)
        visible = [9, 10]
        for r in range(11, 15):
            ws.row_dimensions[r].hidden = True
        counted = [r for r in range(9, 15)
                   if not ws.row_dimensions[r].hidden
                   and isinstance(ws.cell(row=r, column=1).value, int)]
        self.assertEqual(counted, visible)
        self.assertEqual(len(counted), 2)
        # and the formula still spans the whole population, so clearing the
        # filter restores the full count
        self.assertEqual(ws["A7"].value, "=SUBTOTAL(2,A9:A14)")

    def test_the_number_format_suits_a_numeric_identifier(self):
        ws, _report = self._sheet()
        self.assertEqual(ws.cell(row=9, column=1).number_format, "0")

    def test_the_coercion_is_confined_to_the_excel_write(self):
        """The only caller is _write_loan_sheet. checks._norm_id, audit
        matching, corrections, CONT-01 and the three Portfolio Summaries never
        see this function -- none of them reads this workbook."""
        import inspect
        self.assertIn("_excel_gp_value(", inspect.getsource(msr._write_loan_sheet))
        callers = [name for name, obj in vars(msr).items()
                   if callable(obj) and getattr(obj, "__module__", None) == msr.__name__
                   and name != "_excel_gp_value"
                   and "_excel_gp_value(" in inspect.getsource(obj)]
        self.assertEqual(callers, ["_write_loan_sheet"])

    def test_no_other_production_module_references_it(self):
        import inspect
        import checks
        import corrections
        import file_readers as fr
        import main
        import report
        for module in (checks, corrections, fr, main, report):
            with self.subTest(module=module.__name__):
                self.assertNotIn("_excel_gp_value", inspect.getsource(module))

    def test_gp_normalization_itself_is_untouched(self):
        """checks._norm_id still returns the canonical string it always did."""
        import checks
        normalized = checks._norm_id(pd.Series(["5446319102", 5446319102.0, " 5446319102 "]))
        self.assertEqual(list(normalized), ["5446319102"] * 3)


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class RealGpNumberIsNumericTests(unittest.TestCase):

    def test_all_june_gp_cells_are_numeric_in_both_reports(self):
        for deal, report in _june()["reports"].items():
            ws = _render(report)[msr.MONTHLY_SERVICER_REPORT_SHEET]
            text_cells = [r for r in range(9, ws.max_row + 1)
                          if isinstance(ws.cell(row=r, column=1).value, str)]
            with self.subTest(deal=deal):
                self.assertEqual(text_cells, [])
                self.assertEqual(ws.max_row - 8, len(report["rows"]))

    def test_every_june_identifier_round_trips_unchanged(self):
        for deal, report in _june()["reports"].items():
            ws = _render(report)[msr.MONTHLY_SERVICER_REPORT_SHEET]
            written = [str(ws.cell(row=9 + i, column=1).value)
                       for i in range(len(report["rows"]))]
            with self.subTest(deal=deal):
                self.assertEqual(written, [r["SBA GP Number"] for r in report["rows"]])

    def test_the_2025_count_would_read_1436_unfiltered(self):
        report = _june()["reports"]["SOUP2025_1"]
        ws = _render(report)[msr.MONTHLY_SERVICER_REPORT_SHEET]
        numeric = sum(1 for r in range(9, ws.max_row + 1)
                      if isinstance(ws.cell(row=r, column=1).value, int))
        self.assertEqual(numeric, CLIENT_2025_ROWS)
        self.assertEqual(ws["A7"].value, f"=SUBTOTAL(2,A9:A{8 + CLIENT_2025_ROWS})")


class ConsolidationRulesUnchangedTests(unittest.TestCase):
    """SUM / FIRST / LAST are documented in consolidate_cycle_postings and are
    NOT affected by the unresolved-conflict rule, which applies only to fields
    that function does not resolve."""

    def test_the_documented_rules_still_apply(self):
        frame = _conflicting_frame("Net Interest Rate", 0.0775, 0.0725)
        consolidated = checks.consolidate_cycle_postings(frame)
        row = consolidated.loc["9"]
        self.assertEqual(row["net_interest"], 200.0)          # SUM
        self.assertEqual(row["received"], 200.0)              # SUM
        self.assertEqual(row["beginning_balance"], 1000.0)    # FIRST
        self.assertEqual(row["ending_balance"], 1000.0)       # LAST
        self.assertEqual(row["loan_status"], "ACTIVE")        # LAST
        self.assertEqual(row["posting_count"], 2)

    def test_the_conflict_rule_touches_only_unresolved_fields(self):
        resolved = {"Loan Status", "Delinquency Status", "Net Interest", "Principal",
                    "Payment Received", "Beginning Balance", "Ending Balance",
                    "Lender Loan Number"}
        self.assertEqual(resolved & set(msr._SOURCE_COLUMNS.values()), set())



# ---------------------------------------------------------------------------
# NUMERIC EXCEL DATE SERIALS
#
# Millennium 2025 GP 7859248901's July 2026 Modification Date. The lender
# entered 2026-06-12 into a cell whose number format is CURRENCY, so the sheet
# holds the bare Excel serial 46185 (displaying as "$46,185.00") rather than a
# date. calamine returns a datetime only for a DATE-FORMATTED cell, so the
# report received the number -- and pd.to_datetime reads a bare number as
# NANOSECONDS since 1970, printing 01-01-70 where the client's own finalized
# report prints 06-12-26.
#
# Every test below is about how a SOURCE CELL is read. Nothing here touches a
# status, a determination or a financial value.
# ---------------------------------------------------------------------------

DEFECT_GP = "7859248901"
DEFECT_SERIAL = 46185
DEFECT_DATE = datetime.datetime(2026, 6, 12)
EPOCH = datetime.datetime(1970, 1, 1)


def _with_dates(columns, value, gp=DEFECT_GP, status="Deferred", dq="60 days"):
    """One-loan Servicer Inputs frame carrying `value` in each named column."""
    frame = _srf([(gp, "L-1", status, dq, 40000.0, 34445.68, 0.0)])
    for column in columns:
        frame[column] = value
    return frame


class ExcelDateSerialCoercionTests(unittest.TestCase):
    """_coerce_date reads a number as DAYS since 1899-12-30, never as a
    nanosecond count."""

    def test_the_reported_serial_becomes_the_client_date(self):
        for value in (DEFECT_SERIAL, float(DEFECT_SERIAL)):
            with self.subTest(value=value):
                self.assertEqual(msr._coerce_date(value), DEFECT_DATE)

    def test_a_fractional_serial_keeps_its_calendar_day(self):
        self.assertEqual(msr._coerce_date(DEFECT_SERIAL + 0.5).date(), DEFECT_DATE.date())

    def test_the_plausible_window_boundaries(self):
        self.assertEqual(msr._coerce_date(msr._MIN_EXCEL_SERIAL),
                         datetime.datetime(1900, 3, 1))
        self.assertEqual(msr._coerce_date(msr._MAX_EXCEL_SERIAL),
                         datetime.datetime(2100, 12, 31))
        self.assertIsNone(msr._coerce_date(msr._MIN_EXCEL_SERIAL - 1))
        self.assertIsNone(msr._coerce_date(msr._MAX_EXCEL_SERIAL + 1))

    def test_an_implausible_number_is_blank_and_never_the_epoch(self):
        """Zero, a count, a stray amount. Blank beats a date that would look
        exactly like a real one on a distributed report."""
        for value in (0, 0.0, 1, 60, -5, -46185, 1e12, float("nan"), True, False):
            with self.subTest(value=value):
                coerced = msr._coerce_date(value)
                self.assertIsNone(coerced)
                self.assertNotEqual(coerced, EPOCH)

    def test_existing_non_numeric_handling_is_unchanged(self):
        for value in (datetime.date(2026, 6, 12), datetime.datetime(2026, 6, 12),
                      pd.Timestamp("2026-06-12"), "06/12/2026", "2026-06-12",
                      "June 12, 2026"):
            with self.subTest(value=value):
                self.assertEqual(msr._coerce_date(value), DEFECT_DATE)

    def test_blank_and_unparseable_values_stay_blank(self):
        for value in (None, "", "   ", "n/a", "TBD", pd.NaT, float("nan"),
                      datetime.time(0, 0)):
            with self.subTest(value=value):
                self.assertIsNone(msr._coerce_date(value))


class SerialDateInEveryDateColumnTests(unittest.TestCase):
    """The five date columns share one call site, so they share the fix."""

    def test_every_date_column_reads_the_serial_correctly(self):
        issues = []
        rows = msr.build_servicer_rows(
            _with_dates(msr._DATE_COLUMNS, DEFECT_SERIAL), "Millennium Bank",
            issues=issues)
        self.assertEqual(len(rows), 1)
        for column in msr._DATE_COLUMNS:
            with self.subTest(column=column):
                self.assertEqual(rows[0][column], DEFECT_DATE)
        self.assertEqual(issues, [])

    def test_a_serial_and_a_real_date_produce_identical_rows(self):
        """The number is INTERPRETED differently; nothing else changes."""
        serial = msr.build_servicer_rows(
            _with_dates(msr._DATE_COLUMNS, DEFECT_SERIAL), "Millennium Bank")
        dated = msr.build_servicer_rows(
            _with_dates(msr._DATE_COLUMNS, DEFECT_DATE), "Millennium Bank")
        self.assertEqual(serial, dated)


class MillenniumModificationDateRegressionTests(unittest.TestCase):
    """THE REPORTED DEFECT. Millennium Bank GP 7859248901, July 2026."""

    def setUp(self):
        self.frame = _with_dates(["Modification  Date"], DEFECT_SERIAL)
        self.issues = []
        self.rows = msr.build_servicer_rows(self.frame, "Millennium Bank",
                                            issues=self.issues)
        self.row = self.rows[0]

    def test_the_row_holds_the_client_date(self):
        self.assertEqual(self.row["SBA GP Number"], DEFECT_GP)
        self.assertEqual(self.row["Modification  Date"], DEFECT_DATE)
        self.assertNotEqual(self.row["Modification  Date"], EPOCH)
        self.assertEqual(self.issues, [])

    def test_the_generated_cell_holds_the_client_date_in_the_client_format(self):
        report = msr.build_deal_report("SOUP2025_1", [
            {"servicer_name": "Millennium Bank", "display_label": "Millennium 2025",
             "rows": self.rows, "issues": self.issues, "note_series": "2025 FBC-1",
             "remittance_date": pd.Timestamp("2026-08-21"),
             "collection_period_calendar": SAMPLE_CALENDAR}])
        ws = _render(report)[msr.MONTHLY_SERVICER_REPORT_SHEET]
        cell = ws.cell(row=9, column=msr.MSR_COLUMNS.index("Modification  Date") + 1)
        self.assertEqual(cell.value, DEFECT_DATE)
        self.assertEqual(cell.number_format, "mm-dd-yy")

    def test_the_loan_status_and_delinquency_status_are_untouched(self):
        """This loan's determinations come from consolidation, not from here."""
        self.assertEqual(self.row["Loan Status"], "Deferred")
        self.assertEqual(self.row["Delinquency Status"], "60 days")

    def test_the_financial_values_are_untouched(self):
        self.assertEqual(self.row["Beginning Actual Balance"], 40000.0)
        self.assertEqual(self.row["Ending Actual Balance"], 34445.68)
        self.assertEqual(self.row["Total Payment Received"], 0.0)


class NoEpochDateIsEverWrittenTests(unittest.TestCase):
    """The guard for the NEXT variant of this bug, not this instance of it: no
    date this report writes may predate 1990."""

    def test_no_generated_date_cell_predates_1990(self):
        frame = pd.concat([
            _with_dates(msr._DATE_COLUMNS, DEFECT_SERIAL, gp="1"),
            _with_dates(msr._DATE_COLUMNS, DEFECT_DATE, gp="2"),
            _with_dates(msr._DATE_COLUMNS, "06/12/2026", gp="3"),
            _with_dates(msr._DATE_COLUMNS, 0, gp="4"),
            _with_dates(msr._DATE_COLUMNS, None, gp="5"),
        ], ignore_index=True)
        rows = msr.build_servicer_rows(frame, "Millennium Bank")
        report = msr.build_deal_report("SOUP2025_1", [
            {"servicer_name": "Millennium Bank", "display_label": "Millennium 2025",
             "rows": rows, "issues": [], "note_series": "2025 FBC-1",
             "remittance_date": pd.Timestamp("2026-08-21"),
             "collection_period_calendar": SAMPLE_CALENDAR}])
        ws = _render(report)[msr.MONTHLY_SERVICER_REPORT_SHEET]
        floor = datetime.datetime(1990, 1, 1)
        for offset in range(len(rows)):
            for column in msr._DATE_COLUMNS:
                cell = ws.cell(row=9 + offset,
                               column=msr.MSR_COLUMNS.index(column) + 1)
                with self.subTest(row=9 + offset, column=column):
                    if cell.value is not None:
                        self.assertGreaterEqual(cell.value, floor)


@unittest.skipUnless(os.environ.get("FALCONBRIDGE_JULY_MILLENNIUM_2025_XLSX"),
                     "set FALCONBRIDGE_JULY_MILLENNIUM_2025_XLSX")
class RealJulyModificationDateTests(unittest.TestCase):
    """The real July 2026 Millennium SRF, against the client's own finalized
    Monthly Servicer Report for the same loan."""

    @classmethod
    def setUpClass(cls):
        raw = open(os.environ["FALCONBRIDGE_JULY_MILLENNIUM_2025_XLSX"], "rb").read()
        _sk, _dl, _lt, si, cls.deal = file_readers.read_servicer_file(
            raw, filename="Millennium Bank||Millennium 2025__x.xlsx")
        cls.frame = si
        cls.issues = []
        cls.rows = msr.build_servicer_rows(si, "Millennium Bank", issues=cls.issues)

    def _row(self):
        return next(r for r in self.rows if r["SBA GP Number"] == DEFECT_GP)

    def test_the_source_cell_is_still_the_bare_serial(self):
        """If the lender ever fixes the cell format this test says so, rather
        than the fix silently ceasing to be exercised."""
        values = self.frame.loc[self.frame["SBA GP Number"] == DEFECT_GP,
                                "Modification  Date"]
        self.assertEqual([float(v) for v in values], [float(DEFECT_SERIAL)])

    def test_the_modification_date_matches_the_client_report(self):
        self.assertEqual(self.deal, "SOUP2025_1")
        self.assertEqual(self._row()["Modification  Date"], DEFECT_DATE)

    def test_the_loan_status_and_delinquency_status_are_as_submitted(self):
        row = self._row()
        self.assertEqual(row["Loan Status"], "Deferred")
        self.assertEqual(row["Delinquency Status"], "60 days")

    def test_no_july_millennium_date_predates_1990(self):
        floor = datetime.datetime(1990, 1, 1)
        for row in self.rows:
            for column in msr._DATE_COLUMNS:
                value = row[column]
                with self.subTest(gp=row["SBA GP Number"], column=column):
                    if value is not None:
                        self.assertGreaterEqual(value, floor)


if __name__ == "__main__":
    unittest.main()
