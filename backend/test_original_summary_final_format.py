"""Original Portfolio Summary in the Final Audited format —
CLIENT-CONFIRMED (2026-08-11).

THE CHANGE
----------
The Original Portfolio Summary adopts the Final Audited summary's
presentation: one worksheet, three stacked sections (COUNT / BALANCE % /
BALANCE ($)), a Total row per SOUP deal, loan counts beside dollars, and the
same category order, labels and formatting.

PRESENTATION AND AGGREGATION ONLY. The Original must keep reporting exactly
what the lender submitted, and must stay invariant to audit calculations,
reviewer decisions, corrections, verified overrides and Final Audited values.

Two things carry that guarantee, and both are asserted here:

  * checks.original_portfolio_summary_views takes the lender's servicer_inputs
    frame and NOTHING else. It has no parameter for a verified override, an
    excluded-GP list, an audited status map, a reviewer decision or a
    correction, so an audited value has no route in. That is the same
    structural bar portfolio_summary_check already meets, and the reason this
    function exists at all rather than calling verified_portfolio_summary_views
    with empty arguments -- which is provably equivalent today but would
    downgrade the guarantee to "correct because one call site passes {}".

  * group_portfolio_summary_by_deal and generate_final_portfolio_summary_bytes
    consume already-computed records and never see a data source, so reusing
    them for a second report cannot mix the two layers.

Seven protections, in order:

  1  real-file value equivalence   dollars == PORT-01's own record, 8 servicers
  2  structural invariance         the signature admits nothing else
  3  counts from one population    same builder, same eligibility gate
  4  SOUP 2024 / 2025 grouping     real deals, real servicer split
  5  workbook shape                sections, categories, per-deal totals, title
  6  review invariance             a reviewer decision moves nothing here
  7  cross-layer separation        Original holds while Final Audited moves

The real-file tests skip unless the fixtures are present. Run with:
    FALCONBRIDGE_SOUP_XLSX=... FALCONBRIDGE_SRF_DIR=... \\
    python -m pytest test_original_summary_final_format.py
"""

import base64
import glob
import inspect
import io
import os
import unittest

import pandas as pd
from openpyxl import load_workbook

import checks
import report


SOUP_PATH = os.environ.get("FALCONBRIDGE_SOUP_XLSX")
SRF_DIR = os.environ.get("FALCONBRIDGE_SRF_DIR")
REAL = all([SOUP_PATH, SRF_DIR])

# (filename substring, SOUP servicer key, display label, expected SOUP deal).
# Identity comes from the upload filename convention, never an in-sheet label.
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

# The reviewer approval measured earlier on the real June files: Mission
# Valley's GP 4778059109 moves LIQUIDATION -> ACTIVE, re-entering Outstanding
# with an ending balance of 276,702.10. Used by protection 7.
MVB_OVERRIDE_GP = "4778059109"
MVB_OVERRIDE_DELTA = 276702.10


def _srf(rows):
    return pd.DataFrame([{
        "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}", "Loan Status": ls,
        "Delinquency Status": dq, "Beginning Balance": bb, "Ending Balance": eb,
        "Payment Received": rec, "Scheduled Payment": 500.0, "Servicing Fee": 0.0,
    } for gp, ls, dq, bb, eb, rec in rows])


SAMPLE = _srf([
    ("1", "ACTIVE", "CURRENT", 100000.0, 99000.0, 500.0),
    ("2", "ACTIVE", "30 DAYS", 50000.0, 50000.0, 0.0),
    ("3", "LIQUIDATION", "120 DAYS", 25000.0, 25000.0, 0.0),
    ("4", "PAID-IN-FULL", "CURRENT", 10000.0, 0.0, 10000.0),
    ("5", "ACTIVE", "CURRENT", 0.0, 0.0, 0.0),          # closed at zero: ineligible
])


def _find(substring):
    hits = [p for p in glob.glob(os.path.join(SRF_DIR, "*.xlsx"))
            if substring.lower() in os.path.basename(p).lower()]
    return hits[0] if hits else None


def _load_june():
    """(dollars, counts, deals, labels) for all eight June servicers, in a
    fixed order. Uses the real reader, so the deals are the real Note Series
    values rather than anything this test asserts into existence."""
    import file_readers
    dollars, counts, deals, labels = [], [], [], []
    for substring, soup_key, display, _expected in JUNE_SERVICERS:
        path = _find(substring)
        assert path, f"no June SRF found for {display}"
        _, _, _, servicer_inputs_df, deal = file_readers.read_servicer_file(
            open(path, "rb").read(), filename=f"{soup_key}||{display}__x.xlsx")
        views = checks.original_portfolio_summary_views(servicer_inputs_df, display)
        dollars.append(views["dollars"])
        counts.append(views["counts"])
        deals.append(deal)
        labels.append(display)
    return dollars, counts, deals, labels


# ---------------------------------------------------------------------------
# 1. Real-file value equivalence
# ---------------------------------------------------------------------------

@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _SRF_DIR")
class ValueEquivalenceWithPort01Tests(unittest.TestCase):
    """The format changed; not one figure did. This is the headline check:
    every dollar the new function reports equals PORT-01's own record for the
    same lender file, field for field, on real data."""

    @classmethod
    def setUpClass(cls):
        import file_readers
        cls.pairs = []
        for substring, soup_key, display, _expected in JUNE_SERVICERS:
            path = _find(substring)
            assert path, f"no June SRF found for {display}"
            _, _, _, si, _ = file_readers.read_servicer_file(
                open(path, "rb").read(), filename=f"{soup_key}||{display}__x.xlsx")
            cls.pairs.append((
                display,
                checks.portfolio_summary_check(si, display)["records"][0],
                checks.original_portfolio_summary_views(si, display)["dollars"],
            ))

    def test_every_servicer_record_is_identical_to_port01(self):
        for display, port01, views in self.pairs:
            with self.subTest(servicer=display):
                self.assertEqual(port01, views)

    def test_every_field_matches_not_just_outstanding(self):
        for display, port01, views in self.pairs:
            for category in ["Outstanding"] + checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER:
                for suffix in ("($)", "(%)"):
                    key = f"{category} {suffix}"
                    with self.subTest(servicer=display, field=key):
                        self.assertEqual(port01[key], views[key])

    def test_all_eight_servicers_were_exercised(self):
        self.assertEqual(len(self.pairs), 8)

    def test_the_figures_are_not_trivially_zero(self):
        """Equality would be worthless if both sides were empty."""
        for display, port01, _ in self.pairs:
            with self.subTest(servicer=display):
                self.assertGreater(port01["Outstanding ($)"], 0.0)


# ---------------------------------------------------------------------------
# 2. Structural invariance
# ---------------------------------------------------------------------------

class StructuralInvarianceTests(unittest.TestCase):
    """The guarantee is the signature. An audited value, a reviewer decision, a
    correction or a suppression set has no parameter to arrive through."""

    def test_the_signature_is_exactly_the_lender_frame_and_a_name(self):
        params = list(inspect.signature(checks.original_portfolio_summary_views).parameters)
        self.assertEqual(params, ["servicer_inputs_df", "servicer_name"])

    def test_no_audit_or_review_parameter_exists(self):
        params = inspect.signature(checks.original_portfolio_summary_views).parameters
        for forbidden in ("verified_overrides", "excluded_gps", "audited_status_by_loan",
                          "decisions", "corrections", "correction_data", "soup_df"):
            with self.subTest(parameter=forbidden):
                self.assertNotIn(forbidden, params)

    def test_the_body_reads_no_override_or_suppression_helper(self):
        body = inspect.getsource(checks.original_portfolio_summary_views)
        code = "\n".join(line for line in body.splitlines()
                         if not line.strip().startswith("#"))
        code = code.split('"""')[-1]        # drop the docstring, which names them
        for forbidden in ("_apply_verified_overrides", "_verified_population",
                          "verified_portfolio_summary_views", "audited_status_by_loan",
                          "excluded_gps"):
            with self.subTest(symbol=forbidden):
                self.assertNotIn(forbidden, code)

    def test_it_uses_the_same_aging_source_as_port01(self):
        """Both read the row's own servicer-reported dq_status. Neither reads a
        calculated status. This is what makes protection 1 hold by
        construction rather than by coincidence."""
        original = inspect.getsource(checks.original_portfolio_summary_views)
        port01 = inspect.getsource(checks.portfolio_summary_check)
        self.assertIn('row.get("dq_status", "")', original)
        self.assertIn('row.get("dq_status", "")', port01)

    def test_portfolio_summary_check_is_untouched(self):
        """The pre-existing Original entry point keeps its own signature, so
        the two invariance tests guarding it still mean what they meant."""
        params = list(inspect.signature(checks.portfolio_summary_check).parameters)
        self.assertEqual(params, ["servicer_inputs_df", "servicer_name"])


# ---------------------------------------------------------------------------
# 3. Counts come from the same population as the dollars
# ---------------------------------------------------------------------------

class CountsShareThePopulationTests(unittest.TestCase):
    """Counts are a unit-weighted pass over the SAME frame through the SAME
    builder -- not a second classifier that could disagree about membership."""

    def setUp(self):
        self.views = checks.original_portfolio_summary_views(SAMPLE, "Test Bank")

    def test_counts_are_integers(self):
        for category in ["Outstanding"] + checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER:
            with self.subTest(category=category):
                self.assertIsInstance(self.views["counts"][category], int)

    def test_a_category_with_dollars_has_a_nonzero_count(self):
        for category in ["Outstanding"] + checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER:
            with self.subTest(category=category):
                if self.views["dollars"][f"{category} ($)"] > 0:
                    self.assertGreater(self.views["counts"][category], 0)

    def test_outstanding_counts_the_loans_outstanding_excludes_terminal_ones(self):
        # Loans 1 and 2 are outstanding; 3 is LIQUIDATION and 4 PAID-IN-FULL
        # (both excluded from Outstanding); 5 is closed at zero (ineligible).
        self.assertEqual(self.views["counts"]["Outstanding"], 2)

    def test_the_closed_at_zero_loan_is_in_neither_view(self):
        """The eligibility gate must survive unit weighting -- weighting every
        row unconditionally would silently re-admit loan 5, which is 0/0."""
        weighted = checks._unit_weight_frame(checks.consolidate_cycle_postings(SAMPLE))
        self.assertEqual(weighted.loc["5", "beginning_balance"], 0.0)
        self.assertEqual(weighted.loc["5", "ending_balance"], 0.0)
        self.assertEqual(self.views["dollars"]["Current ($)"], 99000.0)

    def test_a_zero_balance_loan_can_count_without_contributing_dollars(self):
        """Loan 4 is PAID-IN-FULL/CURRENT with beginning 10,000 and ending 0.
        It passes the eligibility gate on its beginning balance, so it is one
        loan in Current, but contributes 0 to Current dollars. Counts and
        dollars therefore answer different questions and need not move
        together -- pinned here because it looks like a discrepancy and is
        not. Identical in the Final Audited summary, which uses the same
        builder."""
        self.assertEqual(self.views["counts"]["Current"], 2)          # loans 1 and 4
        self.assertEqual(self.views["dollars"]["Current ($)"], 99000.0)  # loan 1 only
        final = checks.verified_portfolio_summary_views(SAMPLE, {}, "Test Bank")
        self.assertEqual(self.views["counts"], final["counts"])

    def test_liquidation_and_paid_in_full_are_counted_though_not_outstanding(self):
        self.assertEqual(self.views["counts"]["Liquidation"], 1)
        self.assertEqual(self.views["counts"]["Paid In Full"], 1)

    def test_building_counts_does_not_disturb_the_dollar_record(self):
        again = checks.original_portfolio_summary_views(SAMPLE, "Test Bank")
        self.assertEqual(again["dollars"], self.views["dollars"])
        self.assertEqual(
            again["dollars"], checks.portfolio_summary_check(SAMPLE, "Test Bank")["records"][0])

    def test_the_source_frame_is_never_mutated(self):
        before = SAMPLE.copy(deep=True)
        checks.original_portfolio_summary_views(SAMPLE, "Test Bank")
        pd.testing.assert_frame_equal(SAMPLE, before)


# ---------------------------------------------------------------------------
# 4. SOUP 2024 / SOUP 2025 grouping
# ---------------------------------------------------------------------------

@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _SRF_DIR")
class DealGroupingTests(unittest.TestCase):
    """Real Note Series values, real servicer split: five 2024 servicers and
    three 2025 ones, totalled separately and never combined."""

    @classmethod
    def setUpClass(cls):
        cls.dollars, cls.counts, cls.deals, cls.labels = _load_june()
        cls.blocks = checks.group_portfolio_summary_by_deal(
            cls.dollars, cls.counts, cls.deals)

    def test_each_servicer_declares_the_expected_deal(self):
        for (substring, _key, display, expected), actual in zip(JUNE_SERVICERS, self.deals):
            with self.subTest(servicer=display):
                self.assertEqual(actual, expected)

    def test_exactly_two_blocks_2024_then_2025(self):
        self.assertEqual([b["deal"] for b in self.blocks], ["SOUP2024_1", "SOUP2025_1"])

    def test_the_labels_read_as_the_client_specified(self):
        self.assertEqual([b["label"] for b in self.blocks],
                         ["2024 SOUP Deal Total", "2025 SOUP Deal Total"])

    def test_five_servicers_in_2024_and_three_in_2025(self):
        self.assertEqual(len(self.blocks[0]["dollars"]), 5)
        self.assertEqual(len(self.blocks[1]["dollars"]), 3)

    def test_no_unassigned_block_for_the_june_set(self):
        self.assertNotIn(checks.UNASSIGNED_DEAL_LABEL, [b["label"] for b in self.blocks])

    def test_each_deal_total_is_the_sum_of_its_own_servicers(self):
        for block in self.blocks:
            with self.subTest(deal=block["deal"]):
                expected = round(sum(r["Outstanding ($)"] for r in block["dollars"]), 2)
                self.assertAlmostEqual(
                    block["totals"]["dollars"]["Outstanding ($)"], expected, places=2)

    def test_each_deal_counts_its_own_loans(self):
        for block in self.blocks:
            with self.subTest(deal=block["deal"]):
                expected = sum(r["Outstanding"] for r in block["counts"])
                self.assertEqual(block["totals"]["counts"]["Outstanding"], expected)

    def test_every_deal_reads_one_hundred_percent_outstanding(self):
        for block in self.blocks:
            with self.subTest(deal=block["deal"]):
                self.assertEqual(block["totals"]["dollars"]["Outstanding (%)"], 100.0)

    def test_there_is_no_combined_cross_deal_total(self):
        labels = [b["label"] for b in self.blocks]
        self.assertNotIn("Total", labels)
        self.assertEqual(len(self.blocks), 2)

    def test_grouping_changes_no_servicer_figure(self):
        regrouped = [r for b in self.blocks for r in b["dollars"]]
        self.assertEqual(sorted(r["Servicer"] for r in regrouped),
                         sorted(r["Servicer"] for r in self.dollars))
        by_name = {r["Servicer"]: r for r in regrouped}
        for record in self.dollars:
            with self.subTest(servicer=record["Servicer"]):
                self.assertEqual(by_name[record["Servicer"]], record)


# ---------------------------------------------------------------------------
# 5. Workbook shape
# ---------------------------------------------------------------------------

def _sample_blocks():
    a = checks.original_portfolio_summary_views(SAMPLE, "Bank A")
    b = checks.original_portfolio_summary_views(SAMPLE, "Bank B")
    return checks.group_portfolio_summary_by_deal(
        [a["dollars"], b["dollars"]], [a["counts"], b["counts"]],
        ["SOUP2024_1", "SOUP2025_1"])


def _render(title=None, cycle_month=None):
    buf = io.BytesIO()
    kwargs = {"cycle_month": cycle_month}
    if title is not None:
        kwargs["title"] = title
    report.generate_final_portfolio_summary_bytes(_sample_blocks(), buf, **kwargs)
    return load_workbook(io.BytesIO(buf.getvalue()))


class WorkbookShapeTests(unittest.TestCase):
    """The Original workbook must be the Final Audited layout with a different
    name on it -- and naming it must not disturb the Final one."""

    def setUp(self):
        self.wb = _render(title="Original Portfolio Summary")
        self.ws = self.wb["Original Portfolio Summary"]

    def test_one_sheet_named_for_the_original_summary(self):
        self.assertEqual(self.wb.sheetnames, ["Original Portfolio Summary"])

    def test_the_heading_names_the_original_summary(self):
        self.assertEqual(self.ws.cell(row=1, column=1).value, "Original Portfolio Summary")

    def test_it_does_not_claim_to_be_audited_or_final_anywhere(self):
        text = " ".join(str(c.value) for row in self.ws.iter_rows() for c in row
                        if isinstance(c.value, str))
        self.assertNotIn("Final Audited", text)
        self.assertNotIn("Audited", text)

    def test_three_stacked_sections_in_the_confirmed_order(self):
        titles = [self.ws.cell(row=r, column=1).value for r in range(1, self.ws.max_row + 1)]
        found = [t for t in titles if t in report.FINAL_SUMMARY_SECTIONS]
        self.assertEqual(found, list(report.FINAL_SUMMARY_SECTIONS))

    def test_the_category_columns_match_the_final_summary_exactly(self):
        for r in range(1, self.ws.max_row + 1):
            if self.ws.cell(row=r, column=1).value == "COUNT":
                header = [self.ws.cell(row=r + 1, column=c).value
                          for c in range(1, len(report.FINAL_SUMMARY_CATEGORIES) + 2)]
                self.assertEqual(header, ["Servicer"] + report.FINAL_SUMMARY_CATEGORIES)
                return
        self.fail("no COUNT section found")

    def test_outstanding_appears_exactly_once_in_the_header(self):
        self.assertEqual(report.FINAL_SUMMARY_CATEGORIES.count("Outstanding"), 1)

    def test_a_total_row_per_deal_in_every_section(self):
        labels = [self.ws.cell(row=r, column=1).value for r in range(1, self.ws.max_row + 1)]
        for deal_label in ("2024 SOUP Deal Total", "2025 SOUP Deal Total"):
            with self.subTest(label=deal_label):
                self.assertEqual(labels.count(deal_label), len(report.FINAL_SUMMARY_SECTIONS))

    def test_the_reporting_period_line_is_written_when_known(self):
        ws = _render(title="Original Portfolio Summary",
                     cycle_month=pd.Timestamp("2026-06-01"))["Original Portfolio Summary"]
        self.assertEqual(ws.cell(row=2, column=1).value, "Reporting Period: June 2026")

    def test_the_final_audited_workbook_is_unchanged_by_the_new_parameter(self):
        wb = _render()      # no title argument at all -- the pre-existing call shape
        self.assertEqual(wb.sheetnames, ["Final Audited Portfolio Summary"])
        ws = wb["Final Audited Portfolio Summary"]
        self.assertEqual(ws.cell(row=1, column=1).value, "Final Audited Portfolio Summary")

    def test_the_two_layouts_are_identical_apart_from_the_title(self):
        """AMENDED 2026-08-13. Still the same guarantee -- ONE writer, so the
        Original summary cannot drift from the Final Audited layout across
        everything the Original carries -- narrowed to exactly that.

        It previously asserted whole-sheet equality including max_row. The
        client then confirmed a FINAL-AUDITED-ONLY fourth section (CURRENT
        PERIOD TERMINAL ACTIVITY), so the Final Audited sheet is deliberately
        longer. Freezing max_row would have pinned "the Original summary must
        also carry the audited terminal-activity section", which is the opposite
        of what was confirmed: that section reports the audited outcome, and the
        Original summary reports the lender's submission.

        The three shared sections are still compared cell for cell, and the
        next test asserts the ONLY difference is the appended section."""
        original = _render(title="Original Portfolio Summary")["Original Portfolio Summary"]
        final = _render()["Final Audited Portfolio Summary"]
        self.assertEqual(original.max_column, final.max_column)
        for r in range(3, original.max_row + 1):      # row 1 is the title
            for c in range(1, original.max_column + 1):
                with self.subTest(row=r, col=c):
                    self.assertEqual(original.cell(row=r, column=c).value,
                                     final.cell(row=r, column=c).value)

    def test_the_only_extra_content_in_the_final_audited_sheet_is_that_section(self):
        original = _render(title="Original Portfolio Summary")["Original Portfolio Summary"]
        final = _render()["Final Audited Portfolio Summary"]
        self.assertGreater(final.max_row, original.max_row)
        extra = [str(final.cell(row=r, column=1).value or "")
                 for r in range(original.max_row + 1, final.max_row + 1)]
        self.assertIn(report.FINAL_SUMMARY_TERMINAL_SECTION, extra)
        # ...and none of the three shared section titles is repeated down there.
        for section in report.FINAL_SUMMARY_SECTIONS:
            with self.subTest(section=section):
                self.assertNotIn(section, extra)

    def test_the_two_sheet_writer_still_exists_for_the_audited_workbook(self):
        buf = io.BytesIO()
        report.generate_portfolio_summary_bytes(
            [checks.portfolio_summary_check(SAMPLE, "T")["records"][0]], buf)
        wb = load_workbook(io.BytesIO(buf.getvalue()))
        self.assertEqual(wb.sheetnames,
                         ["Portfolio Summary Dollars", "Portfolio Summary Percentages"])


# ---------------------------------------------------------------------------
# 6. Review invariance
# ---------------------------------------------------------------------------

class ReviewInvarianceTests(unittest.TestCase):
    """A reviewer decision moves the Final Audited summary. It must move
    nothing here."""

    OVERRIDES = {("3", "Loan Status"): "ACTIVE", ("1", "Delinquency Status"): "120 DAYS"}

    def test_the_original_views_are_unchanged_by_verified_overrides(self):
        before = checks.original_portfolio_summary_views(SAMPLE, "T")
        checks.verified_portfolio_summary_views(SAMPLE, self.OVERRIDES, "T")
        after = checks.original_portfolio_summary_views(SAMPLE, "T")
        self.assertEqual(before, after)

    def test_the_final_layer_DOES_move_on_the_same_overrides(self):
        """Without this the test above could pass because the overrides are
        inert rather than because the Original is insulated."""
        plain = checks.verified_portfolio_summary_views(SAMPLE, {}, "T")
        overridden = checks.verified_portfolio_summary_views(SAMPLE, self.OVERRIDES, "T")
        self.assertNotEqual(plain["dollars"], overridden["dollars"])

    def test_the_original_differs_from_the_reviewed_final(self):
        original = checks.original_portfolio_summary_views(SAMPLE, "T")["dollars"]
        final = checks.verified_portfolio_summary_views(SAMPLE, self.OVERRIDES, "T")["dollars"]
        self.assertNotEqual(original, final)

    def test_the_original_ignores_closed_at_zero_suppression(self):
        """Suppression is an audit judgement. The lender's submission still
        contains the loan, so the Original still reports it."""
        suppressed = checks.verified_portfolio_summary_views(
            SAMPLE, {}, "T", excluded_gps=["2"])
        original = checks.original_portfolio_summary_views(SAMPLE, "T")
        self.assertEqual(original["dollars"]["30 Days ($)"], 50000.0)
        self.assertEqual(suppressed["dollars"]["30 Days ($)"], 0.0)
        self.assertEqual(original["counts"]["30 Days"], 1)
        self.assertEqual(suppressed["counts"]["30 Days"], 0)

    def test_repeated_calls_are_stable(self):
        first = checks.original_portfolio_summary_views(SAMPLE, "T")
        for _ in range(3):
            self.assertEqual(checks.original_portfolio_summary_views(SAMPLE, "T"), first)


# ---------------------------------------------------------------------------
# 7. Cross-layer separation on real files
# ---------------------------------------------------------------------------

@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _SRF_DIR")
class CrossLayerSeparationTests(unittest.TestCase):
    """The measured proof. Mission Valley's GP 4778059109 moves LIQUIDATION ->
    ACTIVE on reviewer approval, re-entering Outstanding with 276,702.10. The
    Final Audited 2024 total moves by exactly that. The Original 2024 total
    must not move at all."""

    @classmethod
    def setUpClass(cls):
        import file_readers
        cls.frames = {}
        for substring, soup_key, display, deal in JUNE_SERVICERS:
            path = _find(substring)
            assert path, f"no June SRF found for {display}"
            _, _, _, si, actual_deal = file_readers.read_servicer_file(
                open(path, "rb").read(), filename=f"{soup_key}||{display}__x.xlsx")
            cls.frames[display] = (si, actual_deal)

    def _totals(self, views_fn):
        dollars, counts, deals = [], [], []
        for _s, _k, display, _d in JUNE_SERVICERS:
            si, deal = self.frames[display]
            views = views_fn(display, si)
            dollars.append(views["dollars"])
            counts.append(views["counts"])
            deals.append(deal)
        blocks = checks.group_portfolio_summary_by_deal(dollars, counts, deals)
        return {b["deal"]: b["totals"] for b in blocks}

    def test_the_original_deal_totals_are_the_lender_figures(self):
        totals = self._totals(
            lambda display, si: checks.original_portfolio_summary_views(si, display))
        self.assertAlmostEqual(totals["SOUP2024_1"]["dollars"]["Outstanding ($)"],
                               61228003.82, places=2)
        self.assertAlmostEqual(totals["SOUP2025_1"]["dollars"]["Outstanding ($)"],
                               99128983.40, places=2)

    def test_the_final_layer_moves_by_exactly_the_approved_amount(self):
        overrides = {"Mission Valley Bank": {(MVB_OVERRIDE_GP, "Loan Status"): "ACTIVE"}}
        final = self._totals(lambda display, si: checks.verified_portfolio_summary_views(
            si, overrides.get(display, {}), display))
        self.assertAlmostEqual(
            final["SOUP2024_1"]["dollars"]["Outstanding ($)"],
            61228003.82 + MVB_OVERRIDE_DELTA, places=2)

    def test_the_original_holds_while_the_final_moves(self):
        overrides = {"Mission Valley Bank": {(MVB_OVERRIDE_GP, "Loan Status"): "ACTIVE"}}
        original = self._totals(
            lambda display, si: checks.original_portfolio_summary_views(si, display))
        final = self._totals(lambda display, si: checks.verified_portfolio_summary_views(
            si, overrides.get(display, {}), display))
        self.assertNotAlmostEqual(
            original["SOUP2024_1"]["dollars"]["Outstanding ($)"],
            final["SOUP2024_1"]["dollars"]["Outstanding ($)"], places=2)
        self.assertAlmostEqual(
            original["SOUP2024_1"]["dollars"]["Outstanding ($)"], 61228003.82, places=2)

    def test_the_2025_deal_is_untouched_by_a_2024_servicers_review(self):
        overrides = {"Mission Valley Bank": {(MVB_OVERRIDE_GP, "Loan Status"): "ACTIVE"}}
        original = self._totals(
            lambda display, si: checks.original_portfolio_summary_views(si, display))
        final = self._totals(lambda display, si: checks.verified_portfolio_summary_views(
            si, overrides.get(display, {}), display))
        self.assertAlmostEqual(original["SOUP2025_1"]["dollars"]["Outstanding ($)"],
                               final["SOUP2025_1"]["dollars"]["Outstanding ($)"], places=2)

    def test_the_deals_are_never_summed_together(self):
        totals = self._totals(
            lambda display, si: checks.original_portfolio_summary_views(si, display))
        combined = (totals["SOUP2024_1"]["dollars"]["Outstanding ($)"]
                    + totals["SOUP2025_1"]["dollars"]["Outstanding ($)"])
        self.assertAlmostEqual(combined, 160356987.22, places=2)
        for block_totals in totals.values():
            self.assertNotAlmostEqual(
                block_totals["dollars"]["Outstanding ($)"], combined, places=2)


if __name__ == "__main__":
    unittest.main()
