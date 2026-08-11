"""Final Audited Portfolio Summary — COUNT / BALANCE % / BALANCE ($)
(client-confirmed 2026-08-10).

One worksheet, three stacked sections, over the SAME standardized categories
and the SAME verified loan population:

    original lender data + verified_overrides + closed-at-zero suppression

built ONCE by checks._verified_population. Three independently constructed
populations is the failure mode this design exists to prevent.

COUNT is derived by running the EXISTING _build_portfolio_summary_record over a
unit-weight copy of that population: every category is a membership test plus a
sum, so weighting each loan 1.0 turns each dollar total into a count. No second
classification engine, and the protected dollar function is not modified.

Run with:
    python -m pytest test_final_summary_three_sections.py
"""

import io
import unittest

import pandas as pd
from openpyxl import load_workbook

import checks
import report


CATEGORIES = ["Outstanding"] + checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER


def _srf(rows):
    return pd.DataFrame([{
        "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}", "Loan Status": ls,
        "Delinquency Status": dq, "Beginning Balance": bb, "Ending Balance": eb,
        "Payment Received": rec, "Scheduled Payment": 500.0, "Servicing Fee": 0.0,
    } for gp, ls, dq, bb, eb, rec in rows])


# 1 ordinary current loan; 2 deferred + not paying + 60 days (multi-category);
# 3 liquidation; 4 paid-in-full; 5 excluded by the (0, 0) eligibility gate.
SAMPLE = _srf([
    ("1", "ACTIVE", "CURRENT", 100000.0, 99000.0, 500.0),
    ("2", "DEFERRED", "60 DAYS", 50000.0, 50000.0, 0.0),
    ("3", "LIQUIDATION", "120 DAYS", 25000.0, 25000.0, 0.0),
    ("4", "PAID-IN-FULL", "CURRENT", 10000.0, 0.0, 10000.0),
    ("5", "ACTIVE", "CURRENT", 0.0, 0.0, 0.0),
])


def _views(df=SAMPLE, overrides=None, excluded=None, name="T"):
    return checks.verified_portfolio_summary_views(df, overrides or {}, name,
                                                   excluded_gps=excluded)


class CountsReuseTheDollarClassifierTests(unittest.TestCase):
    """Counts must follow every rule the dollar path follows."""

    def test_counts_are_integers(self):
        counts = _views()["counts"]
        for c in CATEGORIES:
            with self.subTest(category=c):
                self.assertIsInstance(counts[c], int)

    def test_the_eligibility_gate_is_preserved(self):
        # Loan 5 has beginning == ending == 0 and must be counted in NOTHING.
        counts = _views()["counts"]
        self.assertEqual(counts["Outstanding"], 2)      # loans 1 and 2 only
        # Loan 4 is PAID-IN-FULL: excluded from Outstanding, but still counted
        # in its aging bucket -- Loan Status is a reporting dimension, not a
        # universal filter. Same rule the dollar path follows.
        self.assertEqual(counts["Current"], 2)          # loans 1 and 4
        self.assertEqual(sum(counts[c] for c in ("Liquidation", "Paid In Full")), 2)

    def test_outstanding_excludes_paid_in_full_and_liquidation(self):
        counts = _views()["counts"]
        self.assertEqual(counts["Outstanding"], 2)
        self.assertEqual(counts["Liquidation"], 1)
        self.assertEqual(counts["Paid In Full"], 1)

    def test_categories_overlap_exactly_as_dollars_do(self):
        # Loan 2 is Deferred AND 60 Days AND Not Making Payment AND Outstanding.
        counts = _views()["counts"]
        self.assertEqual(counts["Deferred"], 1)
        self.assertEqual(counts["60 Days"], 1)
        self.assertEqual(counts["Not Making Payment"], 2)   # loans 2 and 3
        self.assertNotEqual(sum(counts[c] for c in CATEGORIES[1:]), counts["Outstanding"])

    def test_an_unclassifiable_delinquency_status_counts_in_outstanding_only(self):
        # The live Peoples July case: 'FORCLOSURE' in the DQ field.
        df = _srf([("1", "ACTIVE", "FORCLOSURE", 100.0, 100.0, 0.0)])
        counts = _views(df)["counts"]
        self.assertEqual(counts["Outstanding"], 1)
        self.assertEqual(sum(counts[c] for c in ("Current", "Delinquent", "30 Days", "60 Days",
                                                 "90 Days", "120 Days", "150+ Days")), 0)

    def test_counts_match_an_independent_manual_count(self):
        counts = _views()["counts"]
        grouped = checks.consolidate_cycle_postings(SAMPLE)
        manual = {c: 0 for c in CATEGORIES}
        for gp, row in grouped.iterrows():
            eb = float(row.get("ending_balance") or 0)
            bb = float(row.get("beginning_balance") or 0)
            if bb == 0 and eb == 0:
                continue
            ls = str(row.get("loan_status", "")).strip().upper()
            pif = ls in {"PAID-IN-FULL", "PAID IN FULL", "PIF"}
            liq = ls == "LIQUIDATION"
            if not pif and not liq:
                manual["Outstanding"] += 1
            bucket = checks._classify_delinquency_bucket(
                str(row.get("dq_status", "")).strip().upper())
            if bucket:
                manual[bucket] += 1
            rec = row.get("received")
            if rec is not None and pd.notna(rec) and float(rec) == 0:
                manual["Not Making Payment"] += 1
            if ls == "DEFERRED":
                manual["Deferred"] += 1
            elif "BANKRUPT" in ls:
                manual["Bankruptcy"] += 1
            elif "FORECLOS" in ls:
                manual["Foreclosure"] += 1
            if liq:
                manual["Liquidation"] += 1
            if pif:
                manual["Paid In Full"] += 1
        self.assertEqual({c: counts[c] for c in CATEGORIES}, manual)

    def test_the_unit_weight_frame_never_mutates_its_input(self):
        grouped = checks.consolidate_cycle_postings(SAMPLE)
        before = grouped["ending_balance"].tolist()
        checks._unit_weight_frame(grouped)
        self.assertEqual(grouped["ending_balance"].tolist(), before)


class OnePopulationDrivesAllViewsTests(unittest.TestCase):

    def test_the_dollar_view_is_identical_to_the_standalone_record(self):
        # verified_portfolio_summary_record now delegates here; identical output.
        self.assertEqual(_views()["dollars"],
                         checks.verified_portfolio_summary_record(SAMPLE, {}, "T"))

    def test_reviewer_decisions_move_counts_and_dollars_together(self):
        views = _views(overrides={("1", "Delinquency Status"): "90 DAYS"})
        self.assertEqual(views["counts"]["90 Days"], 1)
        self.assertEqual(views["dollars"]["90 Days ($)"], 99000.0)
        self.assertEqual(views["counts"]["Current"], 1)     # loan 4 (PIF) remains
        self.assertEqual(views["dollars"]["Current ($)"], 0.0)  # its ending balance is 0

    def test_a_loan_status_override_moves_counts_and_dollars_together(self):
        views = _views(overrides={("2", "Loan Status"): "ACTIVE"})
        self.assertEqual(views["counts"]["Deferred"], 0)
        self.assertEqual(views["dollars"]["Deferred ($)"], 0.0)

    def test_suppression_removes_a_loan_from_counts_and_dollars_together(self):
        views = _views(excluded=["3"])
        self.assertEqual(views["counts"]["Liquidation"], 0)
        self.assertEqual(views["dollars"]["Liquidation ($)"], 0.0)
        self.assertEqual(views["counts"]["Not Making Payment"], 1)
        self.assertEqual(views["dollars"]["Not Making Payment ($)"], 50000.0)


class PercentageTests(unittest.TestCase):

    def test_every_percentage_is_the_category_over_outstanding(self):
        dollars = _views()["dollars"]
        outstanding = dollars["Outstanding ($)"]
        self.assertEqual(dollars["Outstanding (%)"], 100.0)
        for c in checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER:
            with self.subTest(category=c):
                self.assertAlmostEqual(dollars[f"{c} (%)"],
                                       round(dollars[f"{c} ($)"] / outstanding * 100, 2), places=2)

    def test_zero_outstanding_yields_zero_percent_everywhere(self):
        # The EXISTING convention, preserved: Outstanding is 0.00%, not 100%.
        df = _srf([("1", "PAID-IN-FULL", "CURRENT", 10000.0, 0.0, 10000.0)])
        dollars = _views(df)["dollars"]
        self.assertEqual(dollars["Outstanding ($)"], 0.0)
        for c in CATEGORIES:
            with self.subTest(category=c):
                self.assertEqual(dollars[f"{c} (%)"], 0.0)

    def test_percentages_are_not_required_to_sum_to_100(self):
        dollars = _views()["dollars"]
        total = sum(dollars[f"{c} (%)"] for c in checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER)
        self.assertNotAlmostEqual(total, 100.0, places=1)


class TotalRowTests(unittest.TestCase):

    def setUp(self):
        self.a = checks.verified_portfolio_summary_views(SAMPLE, {}, "A")
        self.b = checks.verified_portfolio_summary_views(
            _srf([("9", "ACTIVE", "CURRENT", 1000.0, 1000.0, 100.0)]), {}, "B")
        self.total = checks.portfolio_summary_totals(
            [self.a["dollars"], self.b["dollars"]], [self.a["counts"], self.b["counts"]])

    def test_total_counts_are_the_sum_of_servicer_counts(self):
        for c in CATEGORIES:
            with self.subTest(category=c):
                self.assertEqual(self.total["counts"][c],
                                 self.a["counts"][c] + self.b["counts"][c])

    def test_total_dollars_are_the_sum_of_servicer_dollars(self):
        for c in CATEGORIES:
            with self.subTest(category=c):
                self.assertAlmostEqual(
                    self.total["dollars"][f"{c} ($)"],
                    round(self.a["dollars"][f"{c} ($)"] + self.b["dollars"][f"{c} ($)"], 2),
                    places=2)

    def test_total_percentages_are_recomputed_never_averaged(self):
        outstanding = self.total["dollars"]["Outstanding ($)"]
        self.assertEqual(self.total["dollars"]["Outstanding (%)"], 100.0)
        for c in checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER:
            recomputed = round(self.total["dollars"][f"{c} ($)"] / outstanding * 100, 2)
            averaged = round((self.a["dollars"][f"{c} (%)"] + self.b["dollars"][f"{c} (%)"]) / 2, 2)
            with self.subTest(category=c):
                self.assertAlmostEqual(self.total["dollars"][f"{c} (%)"], recomputed, places=2)
                if abs(recomputed - averaged) > 0.01:
                    self.assertNotAlmostEqual(self.total["dollars"][f"{c} (%)"], averaged, places=2)

    def test_zero_total_outstanding_yields_zero_percent(self):
        pif = checks.verified_portfolio_summary_views(
            _srf([("1", "PAID-IN-FULL", "CURRENT", 10.0, 0.0, 10.0)]), {}, "P")
        total = checks.portfolio_summary_totals([pif["dollars"]], [pif["counts"]])
        for c in CATEGORIES:
            with self.subTest(category=c):
                self.assertEqual(total["dollars"][f"{c} (%)"], 0.0)


class WorkbookLayoutTests(unittest.TestCase):

    def setUp(self):
        # Two servicers in ONE deal, so the existing layout assertions below
        # keep describing a single block. Multi-deal layout is covered by
        # DealGroupedTotalsTests.
        self.a = checks.verified_portfolio_summary_views(SAMPLE, {}, "Bank A")
        self.b = checks.verified_portfolio_summary_views(
            _srf([("9", "ACTIVE", "CURRENT", 1000.0, 1000.0, 100.0)]), {}, "Bank B")
        self.blocks = checks.group_portfolio_summary_by_deal(
            [self.a["dollars"], self.b["dollars"]],
            [self.a["counts"], self.b["counts"]],
            ["SOUP2024_1", "SOUP2024_1"])
        self.total = self.blocks[0]["totals"]
        buf = io.BytesIO()
        report.generate_final_portfolio_summary_bytes(self.blocks, buf)
        self.wb = load_workbook(io.BytesIO(buf.getvalue()))
        self.ws = self.wb.active
        self.col_a = [self.ws.cell(row=r, column=1).value for r in range(1, self.ws.max_row + 1)]

    def test_a_single_worksheet(self):
        self.assertEqual(len(self.wb.sheetnames), 1)
        self.assertEqual(self.ws.title, "Final Audited Portfolio Summary")

    def test_the_three_sections_appear_in_order(self):
        found = [v for v in self.col_a if v in report.FINAL_SUMMARY_SECTIONS]
        self.assertEqual(found, ["COUNT", "BALANCE %", "BALANCE ($)"])

    def test_total_is_the_first_row_of_every_section(self):
        for section in report.FINAL_SUMMARY_SECTIONS:
            i = self.col_a.index(section)
            with self.subTest(section=section):
                self.assertEqual(self.col_a[i + 1], "Servicer")   # header row
                # Totals are per SOUP deal and carry the deal's label, not a
                # bare "Total" -- these two servicers share SOUP2024_1.
                self.assertEqual(self.col_a[i + 2], "2024 SOUP Deal Total")
                self.assertEqual(self.col_a[i + 3], "Bank A")     # then upload order
                self.assertEqual(self.col_a[i + 4], "Bank B")

    def test_the_category_columns_are_in_the_confirmed_order(self):
        i = self.col_a.index("COUNT")
        headers = [self.ws.cell(row=i + 2, column=c).value
                   for c in range(1, len(CATEGORIES) + 2)]
        self.assertEqual(headers, ["Servicer"] + CATEGORIES)
        self.assertEqual(CATEGORIES[0], "Outstanding")
        self.assertIn("Foreclosure", CATEGORIES)

    def test_substitution_in_is_not_a_category(self):
        self.assertNotIn("SUBSTITUTION-IN", CATEGORIES)
        self.assertNotIn("SUBSTITUTION-IN", [str(v) for v in self.col_a])

    def test_the_written_values_match_the_records_exactly(self):
        # Formatting must never alter a value.
        for section, key, source in (("COUNT", None, self.total["counts"]),
                                     ("BALANCE %", " (%)", self.total["dollars"]),
                                     ("BALANCE ($)", " ($)", self.total["dollars"])):
            i = self.col_a.index(section)
            for col, category in enumerate(CATEGORIES, start=2):
                written = self.ws.cell(row=i + 3, column=col).value   # the Total row
                expected = source[category if key is None else f"{category}{key}"]
                with self.subTest(section=section, category=category):
                    self.assertAlmostEqual(float(written), float(expected), places=2)

    def test_number_formats(self):
        for section, fmt in (("COUNT", "#,##0"), ("BALANCE %", '#,##0.00"%"'),
                             ("BALANCE ($)", '"$"#,##0.00')):
            i = self.col_a.index(section)
            with self.subTest(section=section):
                self.assertEqual(self.ws.cell(row=i + 3, column=2).number_format, fmt)

    def test_the_header_band_is_navy_with_white_text(self):
        i = self.col_a.index("COUNT")
        header = self.ws.cell(row=i + 1, column=1)
        self.assertIn(report.NAVY, str(header.fill.fgColor.rgb))
        self.assertTrue(header.font.bold)


class DealGroupedTotalsTests(unittest.TestCase):
    """CLIENT-CONFIRMED (2026-08-10): Totals are reported PER SOUP DEAL. A
    2024-deal servicer and a 2025-deal servicer belong to different
    securitizations, so there is deliberately NO combined total."""

    def _blocks(self, deals):
        views = [checks.verified_portfolio_summary_views(
            _srf([(str(i), "ACTIVE", "CURRENT", 1000.0 * (i + 1), 1000.0 * (i + 1), 5.0)]),
            {}, f"Bank {i}") for i in range(len(deals))]
        return checks.group_portfolio_summary_by_deal(
            [v["dollars"] for v in views], [v["counts"] for v in views], deals), views

    def test_the_deal_pattern_parses_the_real_identifiers(self):
        self.assertEqual(checks._deal_year_and_series("SOUP2024_1"), (2024, 1))
        self.assertEqual(checks._deal_year_and_series("SOUP2025_1"), (2025, 1))

    def test_malformed_or_blank_deals_do_not_parse(self):
        for bad in ("", None, "SOUP2024", "SOUP20241", "2024 FBC-1", "GARBAGE"):
            with self.subTest(deal=bad):
                self.assertIsNone(checks._deal_year_and_series(bad))

    def test_labels_are_the_confirmed_display_strings(self):
        deals = ["SOUP2024_1", "SOUP2025_1"]
        self.assertEqual(checks._deal_total_label("SOUP2024_1", deals), "2024 SOUP Deal Total")
        self.assertEqual(checks._deal_total_label("SOUP2025_1", deals), "2025 SOUP Deal Total")

    def test_2024_sorts_before_2025_and_unassigned_sorts_last(self):
        blocks, _ = self._blocks(["SOUP2025_1", None, "SOUP2024_1"])
        self.assertEqual([b["label"] for b in blocks],
                         ["2024 SOUP Deal Total", "2025 SOUP Deal Total",
                          checks.UNASSIGNED_DEAL_LABEL])

    def test_servicers_land_in_their_declared_deal_in_upload_order(self):
        blocks, _ = self._blocks(["SOUP2025_1", "SOUP2024_1", "SOUP2024_1", "SOUP2025_1"])
        by_label = {b["label"]: [r["Servicer"] for r in b["counts"]] for b in blocks}
        self.assertEqual(by_label["2024 SOUP Deal Total"], ["Bank 1", "Bank 2"])
        self.assertEqual(by_label["2025 SOUP Deal Total"], ["Bank 0", "Bank 3"])

    def test_a_missing_deal_is_surfaced_not_dropped(self):
        blocks, _ = self._blocks(["SOUP2024_1", None])
        unassigned = next(b for b in blocks if b["label"] == checks.UNASSIGNED_DEAL_LABEL)
        self.assertEqual([r["Servicer"] for r in unassigned["counts"]], ["Bank 1"])
        self.assertEqual(sum(len(b["counts"]) for b in blocks), 2)   # nothing lost

    def test_each_deal_outstanding_is_independently_100_percent(self):
        blocks, _ = self._blocks(["SOUP2024_1", "SOUP2025_1", "SOUP2024_1"])
        for block in blocks:
            with self.subTest(deal=block["deal"]):
                self.assertEqual(block["totals"]["dollars"]["Outstanding (%)"], 100.0)

    def test_each_deal_total_sums_only_its_own_servicers(self):
        blocks, views = self._blocks(["SOUP2024_1", "SOUP2025_1", "SOUP2024_1"])
        by_deal = {b["deal"]: b for b in blocks}
        expected_2024 = round(views[0]["dollars"]["Outstanding ($)"]
                              + views[2]["dollars"]["Outstanding ($)"], 2)
        self.assertAlmostEqual(by_deal["SOUP2024_1"]["totals"]["dollars"]["Outstanding ($)"],
                               expected_2024, places=2)
        self.assertAlmostEqual(by_deal["SOUP2025_1"]["totals"]["dollars"]["Outstanding ($)"],
                               views[1]["dollars"]["Outstanding ($)"], places=2)
        self.assertEqual(by_deal["SOUP2024_1"]["totals"]["counts"]["Outstanding"], 2)
        self.assertEqual(by_deal["SOUP2025_1"]["totals"]["counts"]["Outstanding"], 1)

    def test_deal_percentages_are_recomputed_not_averaged(self):
        blocks, views = self._blocks(["SOUP2024_1", "SOUP2024_1"])
        totals = blocks[0]["totals"]["dollars"]
        outstanding = totals["Outstanding ($)"]
        for category in checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER:
            with self.subTest(category=category):
                self.assertAlmostEqual(
                    totals[f"{category} (%)"],
                    round(totals[f"{category} ($)"] / outstanding * 100, 2), places=2)

    def test_no_combined_total_across_deals_is_produced(self):
        blocks, _ = self._blocks(["SOUP2024_1", "SOUP2025_1"])
        labels = [b["label"] for b in blocks]
        self.assertNotIn("Total", labels)
        self.assertEqual(len(blocks), 2)

    def test_a_shared_year_falls_back_to_the_full_identifier(self):
        # Ambiguity surfaces rather than two deals reading as one label.
        deals = ["SOUP2024_1", "SOUP2024_2"]
        self.assertEqual(checks._deal_total_label("SOUP2024_1", deals), "SOUP2024_1 Total")
        self.assertEqual(checks._deal_total_label("SOUP2024_2", deals), "SOUP2024_2 Total")

    def test_per_servicer_records_are_untouched_by_grouping(self):
        blocks, views = self._blocks(["SOUP2024_1", "SOUP2025_1"])
        emitted = [r for b in blocks for r in b["dollars"]]
        self.assertEqual(sorted(r["Servicer"] for r in emitted), ["Bank 0", "Bank 1"])
        for view in views:
            self.assertIn(view["dollars"], emitted)


class DealGroupedWorkbookTests(unittest.TestCase):
    """The workbook renders every deal block in every section."""

    def setUp(self):
        views = [checks.verified_portfolio_summary_views(
            _srf([(str(i), "ACTIVE", "CURRENT", 1000.0, 1000.0, 5.0)]), {}, f"Bank {i}")
            for i in range(3)]
        blocks = checks.group_portfolio_summary_by_deal(
            [v["dollars"] for v in views], [v["counts"] for v in views],
            ["SOUP2025_1", "SOUP2024_1", "SOUP2024_1"])
        buf = io.BytesIO()
        report.generate_final_portfolio_summary_bytes(blocks, buf)
        self.ws = load_workbook(io.BytesIO(buf.getvalue())).active
        self.col_a = [self.ws.cell(row=r, column=1).value
                      for r in range(1, self.ws.max_row + 1)]

    def test_every_section_shows_2024_block_then_2025_block(self):
        for section in report.FINAL_SUMMARY_SECTIONS:
            i = self.col_a.index(section)
            with self.subTest(section=section):
                self.assertEqual(self.col_a[i + 2], "2024 SOUP Deal Total")
                self.assertEqual(self.col_a[i + 3], "Bank 1")
                self.assertEqual(self.col_a[i + 4], "Bank 2")
                self.assertEqual(self.col_a[i + 5], "2025 SOUP Deal Total")
                self.assertEqual(self.col_a[i + 6], "Bank 0")

    def test_each_deal_reads_100_percent_outstanding_in_the_workbook(self):
        i = self.col_a.index("BALANCE %")
        for offset in (2, 5):                       # the two deal total rows
            with self.subTest(row=i + offset):
                self.assertEqual(self.ws.cell(row=i + offset + 1, column=2).value, 100.0)

    def test_no_row_is_labelled_just_total(self):
        self.assertNotIn("Total", self.col_a)


class ProtectedLogicUnchangedTests(unittest.TestCase):
    """The dollar path must be provably untouched."""

    def test_port01_is_unaffected(self):
        record = checks.portfolio_summary_check(SAMPLE, "T")["records"][0]
        self.assertEqual(record["Outstanding ($)"], 149000.0)
        self.assertEqual(record["Liquidation ($)"], 25000.0)
        self.assertEqual(record["Paid In Full ($)"], 10000.0)

    def test_the_audited_record_is_unaffected(self):
        self.assertEqual(
            checks.audited_portfolio_summary_record(SAMPLE, {}, "T"),
            checks.portfolio_summary_check(SAMPLE, "T")["records"][0])

    def test_building_counts_does_not_disturb_the_dollar_record(self):
        before = checks.verified_portfolio_summary_record(SAMPLE, {}, "T")
        _views()
        self.assertEqual(checks.verified_portfolio_summary_record(SAMPLE, {}, "T"), before)

    def test_the_two_sheet_audited_writer_still_exists_unchanged(self):
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes(
            [checks.portfolio_summary_check(SAMPLE, "T")["records"][0]], buf)
        wb = load_workbook(io.BytesIO(buf.getvalue()))
        self.assertEqual(wb.sheetnames,
                         ["Portfolio Summary Dollars", "Portfolio Summary Percentages"])


if __name__ == "__main__":
    unittest.main()
