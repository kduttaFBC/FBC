"""Final Audited Portfolio Summary — Current Period Terminal Activity section.

WHAT THIS COVERS
----------------
CLIENT-CONFIRMED (2026-08-13). A fourth stacked section, FINAL AUDITED ONLY,
presenting the current period's terminal activity beside the population it came
out of:

    Servicing Population | Outstanding | PIF Count | PIF ($) | Liquidation Count | Liquidation ($)

Servicing Population and Outstanding are POPULATION measures. Paid In Full and
Liquidation are CURRENT PERIOD TERMINAL ACTIVITY and are never added into
either -- their dollars are BEGINNING balance (what left the portfolio this
cycle), not ending balance. Nothing about how PIF, Liquidation or Outstanding is
calculated changed; the only new figure is Servicing Population.

SERVICING POPULATION IS COMPUTED INDEPENDENTLY, never as
Outstanding + PIF + Liquidation. That arithmetic is right on the June 2026
accepted SOUP 2025 files only because every $0-Ending-Balance loan there happens
to be a Paid In Full, and it breaks in two ways that occur in real data: it
double-counts a terminal loan carrying a positive balance, and it misses an
ACTIVE loan at $0. The divergence test below pins exactly that.

THE ACCEPTED BASELINE (June 2026, SOUP 2025 deal) is the business acceptance
baseline and is asserted here as fixed numbers:

    Lendistry 2025            pop 1,234  outstanding 1,230  PIF 4 / $739,973.59    Liq 0 / $0.00
    Community Bank & Trust    pop   155  outstanding   152  PIF 3 / $679,563.79    Liq 0 / $0.00
    Millennium 2025           pop    47  outstanding    46  PIF 1 /  $45,998.74    Liq 0 / $0.00
    SOUP 2025 Total           pop 1,436  outstanding 1,428  PIF 8 / $1,465,536.12  Liq 0 / $0.00

Those files are not repo fixtures, so the baseline class skips unless the
accepted-file paths are supplied. Everything else here is synthetic and runs
unconditionally -- the invariants must not depend on files being present.

    FALCONBRIDGE_ACCEPTED_JUNE_DIR=<dir with the three accepted vF workbooks> \\
    FALCONBRIDGE_SOUP_XLSX=... python -m pytest test_final_summary_terminal_activity.py
"""

import glob
import io
import os
import unittest

import pandas as pd
from openpyxl import load_workbook

import checks
import file_readers
import report


ACCEPTED_DIR = os.environ.get("FALCONBRIDGE_ACCEPTED_JUNE_DIR")
SOUP_XLSX = os.environ.get("FALCONBRIDGE_SOUP_XLSX")
REAL = bool(ACCEPTED_DIR and SOUP_XLSX
            and os.path.isdir(ACCEPTED_DIR) and os.path.exists(SOUP_XLSX))

# servicer soup_key, display label, and the accepted figures
BASELINE = [
    ("Lendistry", "Lendistry 2025", "lendistry_soup 2025-1", 1234, 1230, 4, 739973.59, 0, 0.00),
    ("Community Bank & Trust", "Community Bank and Trust", "community bank and trust", 155, 152, 3, 679563.79, 0, 0.00),
    ("Millennium Bank", "Millennium 2025", "millennium_soup 2025-1", 47, 46, 1, 45998.74, 0, 0.00),
]
DEAL_TOTAL = dict(population=1436, outstanding=1428, pif=8, pif_amt=1465536.12, liq=0, liq_amt=0.00)
CLOSED_AT_ZERO_GP = "4338609104"      # CB&T, already closed in a prior period


def _row(**over):
    row = {
        "SBA GP Number": "1000000001",
        "Loan Status": "ACTIVE",
        "Delinquency Status": "CURRENT",
        "Payment Received": 250.0,
        "Principal": 100.0,
        "Scheduled Payment": 250.0,
        "Beginning Balance": 10000.0,
        "Ending Balance": 9900.0,
        "Loan Payment Instance": 1.0,
        "Last Payment Instance?": "Y",
        "_excel_row": 10,
    }
    row.update(over)
    return row


def _views(rows, overrides=None, excluded=None):
    return checks.verified_portfolio_summary_views(
        pd.DataFrame(rows), overrides or {}, "S", excluded_gps=excluded)


class ServicingPopulationTests(unittest.TestCase):
    """The one new figure."""

    def test_it_counts_the_eligible_population(self):
        v = _views([_row(**{"SBA GP Number": "1000000001"}),
                    _row(**{"SBA GP Number": "1000000002"})])
        self.assertEqual(v["counts"][checks.SERVICING_POPULATION], 2)

    def test_the_eligibility_gate_still_applies(self):
        # beginning == 0 AND ending == 0 is excluded before any categorisation,
        # so it is not part of the population being reported on either.
        v = _views([_row(**{"SBA GP Number": "1000000001"}),
                    _row(**{"SBA GP Number": "1000000002",
                            "Beginning Balance": 0.0, "Ending Balance": 0.0})])
        self.assertEqual(v["counts"][checks.SERVICING_POPULATION], 1)

    def test_closed_at_zero_exclusions_are_not_in_the_population(self):
        v = _views([_row(**{"SBA GP Number": "1000000001"}),
                    _row(**{"SBA GP Number": "1000000002"})],
                   excluded=["1000000002"])
        self.assertEqual(v["counts"][checks.SERVICING_POPULATION], 1)

    def test_a_paid_in_full_loan_is_in_the_population_but_not_outstanding(self):
        v = _views([_row(**{"SBA GP Number": "1000000001"}),
                    _row(**{"SBA GP Number": "1000000002",
                            "Loan Status": "PAID-IN-FULL", "Ending Balance": 0.0})])
        self.assertEqual(v["counts"][checks.SERVICING_POPULATION], 2)
        self.assertEqual(v["counts"]["Outstanding"], 1)
        self.assertEqual(v["counts"]["Paid In Full"], 1)

    def test_it_is_NOT_outstanding_plus_pif_plus_liquidation(self):
        # THE reason this is computed independently. A LIQUIDATION loan carrying
        # a POSITIVE balance is excluded from Outstanding but still counted in
        # Liquidation, so the sum double-counts it. Real occurrence: Mission
        # Valley GP 4778059109 at $276,702.10.
        v = _views([_row(**{"SBA GP Number": "1000000001"}),
                    _row(**{"SBA GP Number": "1000000002",
                            "Loan Status": "LIQUIDATION", "Ending Balance": 5000.0})])
        c = v["counts"]
        self.assertEqual(c[checks.SERVICING_POPULATION], 2)
        self.assertEqual(c["Outstanding"], 1)
        self.assertEqual(c["Liquidation"], 1)
        self.assertEqual(c["Outstanding"] + c["Paid In Full"] + c["Liquidation"], 2)
        # ...and the other direction: an ACTIVE loan at $0 is IN Outstanding, so
        # the sum can also understate. Both are why the shortcut is refused.
        v2 = _views([_row(**{"SBA GP Number": "1000000003", "Ending Balance": 0.0})])
        self.assertEqual(v2["counts"][checks.SERVICING_POPULATION], 1)
        self.assertEqual(v2["counts"]["Outstanding"], 1)
        self.assertEqual(v2["counts"]["Paid In Full"], 0)

    def test_an_empty_population_is_zero_not_an_error(self):
        self.assertEqual(checks._eligible_loan_count(None), 0)
        self.assertEqual(checks._eligible_loan_count(pd.DataFrame()), 0)

    def test_the_original_summary_exposes_the_same_key(self):
        v = checks.original_portfolio_summary_views(pd.DataFrame([_row()]), "S")
        self.assertEqual(v["counts"][checks.SERVICING_POPULATION], 1)

    def test_it_is_not_a_category(self):
        # Adding it to the category order would give it a column in all three
        # existing sections and a meaningless dollar figure.
        self.assertNotIn(checks.SERVICING_POPULATION, checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER)
        self.assertNotIn(checks.SERVICING_POPULATION, report.FINAL_SUMMARY_CATEGORIES)
        self.assertNotIn(f"{checks.SERVICING_POPULATION} ($)", _views([_row()])["dollars"])


class DealTotalTests(unittest.TestCase):
    def test_the_total_sums_servicing_population(self):
        a = _views([_row(**{"SBA GP Number": "1000000001"})])
        b = _views([_row(**{"SBA GP Number": "1000000002"}),
                    _row(**{"SBA GP Number": "1000000003"})])
        t = checks.portfolio_summary_totals([a["dollars"], b["dollars"]],
                                            [a["counts"], b["counts"]])
        self.assertEqual(t["counts"][checks.SERVICING_POPULATION], 3)

    def test_a_record_without_the_key_contributes_zero(self):
        a = _views([_row()])
        legacy = {k: v for k, v in a["counts"].items() if k != checks.SERVICING_POPULATION}
        t = checks.portfolio_summary_totals([a["dollars"], a["dollars"]],
                                            [a["counts"], legacy])
        self.assertEqual(t["counts"][checks.SERVICING_POPULATION], 1)


class ExistingRulesUnchangedTests(unittest.TestCase):
    """This change must not move Outstanding, PIF or Liquidation."""

    def test_outstanding_still_excludes_pif_and_liquidation(self):
        v = _views([_row(**{"SBA GP Number": "1000000001"}),
                    _row(**{"SBA GP Number": "1000000002",
                            "Loan Status": "PAID-IN-FULL", "Ending Balance": 0.0}),
                    _row(**{"SBA GP Number": "1000000003",
                            "Loan Status": "LIQUIDATION", "Ending Balance": 0.0})])
        self.assertEqual(v["counts"]["Outstanding"], 1)

    def test_pif_and_liquidation_dollars_still_use_beginning_balance(self):
        v = _views([_row(**{"Loan Status": "PAID-IN-FULL",
                            "Beginning Balance": 20000.0, "Ending Balance": 0.0})])
        self.assertEqual(v["dollars"]["Paid In Full ($)"], 20000.0)
        v2 = _views([_row(**{"Loan Status": "LIQUIDATION",
                             "Beginning Balance": 30000.0, "Ending Balance": 0.0})])
        self.assertEqual(v2["dollars"]["Liquidation ($)"], 30000.0)

    def test_liquidation_reflects_the_verified_state_not_the_submission(self):
        rows = [_row(**{"SBA GP Number": "1000000001", "Loan Status": "LIQUIDATION",
                        "Beginning Balance": 20088.08, "Ending Balance": 0.0})]
        self.assertEqual(_views(rows)["counts"]["Liquidation"], 1)
        approved = _views(rows, overrides={("1000000001", "Loan Status"): "ACTIVE"})
        self.assertEqual(approved["counts"]["Liquidation"], 0)
        self.assertEqual(approved["dollars"]["Liquidation ($)"], 0.0)
        # ...and the population is unchanged by the approval.
        self.assertEqual(approved["counts"][checks.SERVICING_POPULATION], 1)


class SectionRenderingTests(unittest.TestCase):
    def _sheet(self, title=report.FINAL_SUMMARY_DEFAULT_TITLE):
        v = _views([_row(**{"SBA GP Number": "1000000001"}),
                    _row(**{"SBA GP Number": "1000000002",
                            "Loan Status": "PAID-IN-FULL",
                            "Beginning Balance": 20000.0, "Ending Balance": 0.0})])
        blocks = checks.group_portfolio_summary_by_deal(
            [v["dollars"]], [v["counts"]], ["SOUP2025_1"])
        buf = io.BytesIO()
        report.generate_final_portfolio_summary_bytes(blocks, buf, title=title)
        buf.seek(0)
        ws = load_workbook(buf)[title]
        return ws, [str(ws.cell(row=r, column=1).value or "")
                    for r in range(1, ws.max_row + 1)]

    def test_the_section_renders_in_the_final_audited_workbook(self):
        _ws, col_a = self._sheet()
        self.assertIn(report.FINAL_SUMMARY_TERMINAL_SECTION, col_a)

    def test_it_renders_AFTER_the_three_existing_sections(self):
        _ws, col_a = self._sheet()
        order = [v for v in col_a
                 if v in tuple(report.FINAL_SUMMARY_SECTIONS) + (report.FINAL_SUMMARY_TERMINAL_SECTION,)]
        self.assertEqual(order, list(report.FINAL_SUMMARY_SECTIONS)
                         + [report.FINAL_SUMMARY_TERMINAL_SECTION])

    def test_the_original_summary_does_NOT_render_it(self):
        _ws, col_a = self._sheet(title="Original Portfolio Summary")
        self.assertNotIn(report.FINAL_SUMMARY_TERMINAL_SECTION, col_a)

    def test_the_section_is_not_in_the_shared_sections_tuple(self):
        # Guards test_original_summary_final_format.py's exact-equality assertion.
        self.assertNotIn(report.FINAL_SUMMARY_TERMINAL_SECTION, report.FINAL_SUMMARY_SECTIONS)

    def test_the_headers_are_the_confirmed_columns(self):
        ws, col_a = self._sheet()
        top = col_a.index(report.FINAL_SUMMARY_TERMINAL_SECTION) + 1
        headers = [str(ws.cell(row=top + 1, column=c).value or "")
                   for c in range(1, len(report.FINAL_SUMMARY_TERMINAL_COLUMNS) + 2)]
        self.assertEqual(headers, ["Servicer"] +
                         [h for h, _k, _t in report.FINAL_SUMMARY_TERMINAL_COLUMNS])

    def test_the_values_match_the_records(self):
        ws, col_a = self._sheet()
        top = col_a.index(report.FINAL_SUMMARY_TERMINAL_SECTION) + 1
        total = [ws.cell(row=top + 2, column=c).value
                 for c in range(2, len(report.FINAL_SUMMARY_TERMINAL_COLUMNS) + 2)]
        # population 2, outstanding 1, PIF 1 / $20,000.00, Liquidation 0 / $0.00
        self.assertEqual(total, [2, 1, 1, 20000.0, 0, 0.0])

    def test_counts_are_integer_formatted_and_dollars_currency_formatted(self):
        ws, col_a = self._sheet()
        top = col_a.index(report.FINAL_SUMMARY_TERMINAL_SECTION) + 1
        for i, (_h, _k, kind) in enumerate(report.FINAL_SUMMARY_TERMINAL_COLUMNS):
            fmt = ws.cell(row=top + 2, column=2 + i).number_format
            with self.subTest(column=_h):
                self.assertEqual(fmt, report._TERMINAL_INT_FORMAT if kind == "int"
                                 else report._TERMINAL_MONEY_FORMAT)

    def test_a_missing_key_renders_zero_rather_than_raising(self):
        vals = report._terminal_activity_values({}, {})
        self.assertEqual(vals, [0, 0, 0, 0.0, 0, 0.0])


@unittest.skipUnless(REAL, "set FALCONBRIDGE_ACCEPTED_JUNE_DIR and FALCONBRIDGE_SOUP_XLSX")
class AcceptedJune2026Baseline(unittest.TestCase):
    """THE BUSINESS ACCEPTANCE BASELINE — June 2026, SOUP 2025 deal, accepted
    Final files. These are the figures the client's management report carries."""

    @classmethod
    def setUpClass(cls):
        soup = file_readers.read_soup_file(open(SOUP_XLSX, "rb").read())
        cls.views = {}
        for soup_key, label, needle, *_rest in BASELINE:
            hits = [p for p in glob.glob(os.path.join(ACCEPTED_DIR, "*.xlsx"))
                    if all(t in os.path.basename(p).lower() for t in needle.split("_"))]
            if not hits:
                raise unittest.SkipTest(f"no accepted file matching {needle!r}")
            _k, lbl, _lt, si, _d = file_readers.read_servicer_file(
                open(hits[0], "rb").read(), os.path.basename(hits[0]))
            ss = file_readers.filter_soup_by_servicer(soup, soup_key)
            caz = checks.loans_closed_at_zero(ss)
            cls.views[label] = checks.verified_portfolio_summary_views(
                si, {}, label, excluded_gps=list(caz.keys() if isinstance(caz, dict) else caz or []))

    def test_per_lender_figures(self):
        for _sk, label, _n, pop, out, pif, pif_amt, liq, liq_amt in BASELINE:
            v = self.views[label]
            with self.subTest(servicer=label):
                self.assertEqual(v["counts"][checks.SERVICING_POPULATION], pop)
                self.assertEqual(v["counts"]["Outstanding"], out)
                self.assertEqual(v["counts"]["Paid In Full"], pif)
                self.assertAlmostEqual(v["dollars"]["Paid In Full ($)"], pif_amt, places=2)
                self.assertEqual(v["counts"]["Liquidation"], liq)
                self.assertAlmostEqual(v["dollars"]["Liquidation ($)"], liq_amt, places=2)

    def test_deal_total(self):
        t = checks.portfolio_summary_totals(
            [v["dollars"] for v in self.views.values()],
            [v["counts"] for v in self.views.values()], label="SOUP 2025 Total")
        self.assertEqual(t["counts"][checks.SERVICING_POPULATION], DEAL_TOTAL["population"])
        self.assertEqual(t["counts"]["Outstanding"], DEAL_TOTAL["outstanding"])
        self.assertEqual(t["counts"]["Paid In Full"], DEAL_TOTAL["pif"])
        self.assertAlmostEqual(t["dollars"]["Paid In Full ($)"], DEAL_TOTAL["pif_amt"], places=2)
        self.assertEqual(t["counts"]["Liquidation"], DEAL_TOTAL["liq"])
        self.assertAlmostEqual(t["dollars"]["Liquidation ($)"], DEAL_TOTAL["liq_amt"], places=2)

    def test_zero_liquidations_at_every_level(self):
        for label, v in self.views.items():
            with self.subTest(servicer=label):
                self.assertEqual(v["counts"]["Liquidation"], 0)

    def test_cbt_closed_at_zero_loan_is_not_current_period_activity(self):
        # GP 4338609104 is absent from the accepted CB&T file AND excluded by
        # closed-at-zero. It must contribute to neither Liquidation nor the
        # servicing population.
        v = self.views["Community Bank and Trust"]
        self.assertEqual(v["counts"]["Liquidation"], 0)
        self.assertEqual(v["counts"][checks.SERVICING_POPULATION], 155)


if __name__ == "__main__":
    unittest.main()
