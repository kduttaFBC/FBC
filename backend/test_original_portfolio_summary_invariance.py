"""Original Portfolio Summary invariance (client-confirmed 2026-08-11).

The Original Portfolio Summary was removed from the client-facing output on
2026-08-10 and RESTORED on 2026-08-11 as the PROMINENT top-level download, with
the Audited workbook retained as a secondary download. PORT-01 itself never
stopped running and is unchanged by either decision.

The client-facing contract these tests exist to protect:

    Original Portfolio Summary          lender submitted data ONLY
    Audit Findings and Recommendations  machine calculated audit results
    Final Audited Portfolio Summary     reviewer approved corrections

"Lender submitted data ONLY" is the whole point, and it is the property most
likely to rot silently -- nothing about a wrong Original summary looks wrong on
screen. So every route by which audit-derived, reviewer-derived, correction-
derived or suppression-derived data could reach it is asserted closed here,
one test per route:

  * STAT-01's calculated status          test_original_uses_lender_dq_*
  * closed-at-zero suppression           test_original_ignores_closed_at_zero_*
  * reviewer decisions / overrides       test_original_invariant_to_*
  * the whole /api/audit + review flow   OriginalSummaryEndpointTests

These are deliberately BEHAVIOURAL rather than structural: they assert the
values that come out, not which function produced them, so a future refactor
that keeps the contract passes and one that quietly re-points PORT-01 at an
audited source fails.

No fixtures required. Run with:
    python -m pytest test_original_portfolio_summary_invariance.py
"""

import base64
import io
import unittest

import pandas as pd
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

import checks
import file_readers
import main


def _srf(rows):
    """Same SRF shape the other portfolio-summary suites use."""
    return pd.DataFrame([{
        "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}", "Loan Status": ls,
        "Delinquency Status": dq, "Beginning Balance": bb, "Ending Balance": eb,
        "Payment Received": rec, "Scheduled Payment": 500.0, "Servicing Fee": 0.0,
    } for gp, ls, dq, bb, eb, rec in rows])


# Loan "1" is the divergence carrier: the lender reports CURRENT, the audit
# calculates 60 DAYS, and the reviewer approves 90 DAYS. One loan therefore
# lands in a DIFFERENT bucket in each of the three layers, which is what makes
# a three-way comparison meaningful rather than vacuous.
SAMPLE = _srf([
    ("1", "ACTIVE", "CURRENT", 100000.0, 99000.0, 500.0),
    ("2", "ACTIVE", "30 DAYS", 50000.0, 50000.0, 0.0),
])

AUDITED_STATUS = {"1": "60 DAYS"}
REVIEWER_OVERRIDES = {("1", "Delinquency Status"): "90 DAYS"}


class OriginalUsesLenderStatusTests(unittest.TestCase):
    """PORT-01's aging buckets read the LENDER's Delinquency Status, never
    STAT-01's calculated_dq. This is the single most important guard in the
    file: it is the regression that would silently put audit-derived data into
    the summary labelled "Original"."""

    def test_original_buckets_by_lender_reported_status(self):
        record = checks.portfolio_summary_check(SAMPLE, servicer_name="Test Bank")["records"][0]
        # The lender said CURRENT for loan 1, so its ending balance is Current.
        self.assertEqual(record["Current ($)"], 99000.0)
        self.assertEqual(record["60 Days ($)"], 0.0)
        self.assertEqual(record["30 Days ($)"], 50000.0)

    def test_audited_buckets_by_calculated_status(self):
        record = checks.audited_portfolio_summary_record(
            SAMPLE, AUDITED_STATUS, servicer_name="Test Bank",
        )
        # The audit calculated 60 DAYS for loan 1, so it moves out of Current.
        self.assertEqual(record["Current ($)"], 0.0)
        self.assertEqual(record["60 Days ($)"], 99000.0)
        self.assertEqual(record["30 Days ($)"], 50000.0)

    def test_original_and_audited_actually_diverge_on_this_fixture(self):
        # Guards the guard: if the fixture ever stopped producing a divergence,
        # every other assertion here would still pass while testing nothing.
        original = checks.portfolio_summary_check(SAMPLE, servicer_name="Test Bank")["records"][0]
        audited = checks.audited_portfolio_summary_record(
            SAMPLE, AUDITED_STATUS, servicer_name="Test Bank",
        )
        self.assertNotEqual(original["Current ($)"], audited["Current ($)"])
        self.assertNotEqual(original["60 Days ($)"], audited["60 Days ($)"])

    def test_non_aging_categories_are_identical_across_layers(self):
        # Outstanding, Not Making Payment, Deferred/Bankruptcy/Foreclosure,
        # Liquidation and Paid In Full never consult the status function, so
        # they MUST agree. If one of these ever diverges, an audited value has
        # leaked into a column that is supposed to be lender-sourced in both.
        original = checks.portfolio_summary_check(SAMPLE, servicer_name="Test Bank")["records"][0]
        audited = checks.audited_portfolio_summary_record(
            SAMPLE, AUDITED_STATUS, servicer_name="Test Bank",
        )
        for category in ["Outstanding", "Not Making Payment", "Deferred",
                         "Bankruptcy", "Foreclosure", "Liquidation", "Paid In Full"]:
            self.assertEqual(
                original[f"{category} ($)"], audited[f"{category} ($)"],
                f"{category} must not differ between Original and Audited",
            )

    def test_original_takes_no_audited_status_argument_at_all(self):
        # Structural backstop for the behavioural tests above: the audited
        # status simply has no parameter to arrive through.
        import inspect
        params = inspect.signature(checks.portfolio_summary_check).parameters
        self.assertEqual(list(params), ["servicer_inputs_df", "servicer_name"])


class OriginalIgnoresClosedAtZeroTests(unittest.TestCase):
    """Closed-at-zero suppression is an AUDIT judgement -- the audit deciding
    the client's own SOUP history establishes a loan as already closed. It
    therefore applies to the Audited summary and NOT to the Original one, whose
    job is to report what the lender submitted, including loans the audit went
    on to reject."""

    def test_suppressed_loan_still_appears_in_original(self):
        record = checks.portfolio_summary_check(SAMPLE, servicer_name="Test Bank")["records"][0]
        # Loan 2 is suppressed below; Original must still carry its balance.
        self.assertEqual(record["30 Days ($)"], 50000.0)
        self.assertEqual(record["Outstanding ($)"], 149000.0)

    def test_suppressed_loan_is_removed_from_audited(self):
        record = checks.audited_portfolio_summary_record(
            SAMPLE, AUDITED_STATUS, servicer_name="Test Bank", excluded_gps=["2"],
        )
        self.assertEqual(record["30 Days ($)"], 0.0)
        self.assertEqual(record["Outstanding ($)"], 99000.0)

    def test_original_has_no_exclusion_parameter(self):
        import inspect
        self.assertNotIn(
            "excluded_gps", inspect.signature(checks.portfolio_summary_check).parameters,
        )


class OriginalInvariantToReviewerDecisionsTests(unittest.TestCase):
    """A reviewer decision must never move an Original figure. The Final
    Audited summary is where approved corrections land."""

    def test_original_unchanged_by_verified_overrides(self):
        before = checks.portfolio_summary_check(SAMPLE, servicer_name="Test Bank")["records"][0]
        # Run the Final layer, which consumes the same source frame.
        checks.verified_portfolio_summary_views(
            SAMPLE, REVIEWER_OVERRIDES, servicer_name="Test Bank",
        )
        after = checks.portfolio_summary_check(SAMPLE, servicer_name="Test Bank")["records"][0]
        self.assertEqual(before, after)

    def test_verified_overrides_do_not_mutate_the_source_frame(self):
        # _apply_verified_overrides copies rather than mutating; if that ever
        # regressed, the Original summary computed AFTER a review would silently
        # inherit the reviewer's values.
        snapshot = SAMPLE.copy(deep=True)
        checks.verified_portfolio_summary_views(
            SAMPLE, REVIEWER_OVERRIDES, servicer_name="Test Bank",
        )
        pd.testing.assert_frame_equal(SAMPLE, snapshot)

    def test_final_layer_does_reflect_the_override(self):
        # Again guarding the guard: the invariance assertions above are only
        # meaningful if the override actually does something somewhere.
        views = checks.verified_portfolio_summary_views(
            SAMPLE, REVIEWER_OVERRIDES, servicer_name="Test Bank",
        )
        self.assertEqual(views["dollars"]["90 Days ($)"], 99000.0)
        self.assertEqual(views["dollars"]["Current ($)"], 0.0)


class ThreeLayersDivergeTests(unittest.TestCase):
    """Original != Audited != Final on one fixture, in the one bucket each
    layer disagrees about. If any two collapse onto each other, a layer has
    started reading the wrong source."""

    def test_each_layer_puts_loan_one_in_a_different_bucket(self):
        original = checks.portfolio_summary_check(SAMPLE, servicer_name="Test Bank")["records"][0]
        audited = checks.audited_portfolio_summary_record(
            SAMPLE, AUDITED_STATUS, servicer_name="Test Bank",
        )
        final = checks.verified_portfolio_summary_views(
            SAMPLE, REVIEWER_OVERRIDES, servicer_name="Test Bank",
        )["dollars"]

        self.assertEqual(original["Current ($)"], 99000.0)   # lender said CURRENT
        self.assertEqual(audited["60 Days ($)"], 99000.0)    # audit said 60 DAYS
        self.assertEqual(final["90 Days ($)"], 99000.0)      # reviewer said 90 DAYS

        # And the three records are genuinely distinct objects, not aliases.
        self.assertNotEqual(original, audited)
        self.assertNotEqual(audited, final)
        self.assertNotEqual(original, final)

    def test_outstanding_agrees_across_all_three_layers(self):
        # No layer changes a balance -- only which bucket a balance falls in.
        # Outstanding is the invariant that proves no money was invented.
        original = checks.portfolio_summary_check(SAMPLE, servicer_name="Test Bank")["records"][0]
        audited = checks.audited_portfolio_summary_record(
            SAMPLE, AUDITED_STATUS, servicer_name="Test Bank",
        )
        final = checks.verified_portfolio_summary_views(
            SAMPLE, REVIEWER_OVERRIDES, servicer_name="Test Bank",
        )["dollars"]
        self.assertEqual(original["Outstanding ($)"], 149000.0)
        self.assertEqual(audited["Outstanding ($)"], 149000.0)
        self.assertEqual(final["Outstanding ($)"], 149000.0)


# ---------------------------------------------------------------------------
# End-to-end: the /api/audit response and the review round trip
# ---------------------------------------------------------------------------

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
    for i, r in enumerate(rows, start=HEADER_ROW + 1):
        for col, value in enumerate(r, start=1):
            ws.cell(row=i, column=col, value=value)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class OriginalSummaryEndpointTests(unittest.TestCase):
    """POST /api/audit must return BOTH workbooks -- Original as the prominent
    top-level download, Audited retained as secondary -- and the Original bytes
    must be unchanged by a subsequent review."""

    def setUp(self):
        self.file_bytes = _srf_workbook_bytes([
            ("1", "L-1", "ACTIVE", "CURRENT", 500.0, 99000.0, 100000.0, 500.0),
            ("2", "L-2", "ACTIVE", "30 DAYS", 0.0, 50000.0, 50000.0, 500.0),
        ])
        self.servicer_inputs_df = _srf([
            ("1", "ACTIVE", "CURRENT", 100000.0, 99000.0, 500.0),
            ("2", "ACTIVE", "30 DAYS", 50000.0, 50000.0, 0.0),
        ])
        # SOUP disagrees with the lender on loan 1, so STAT-01 has something
        # real to calculate and the Audited summary can genuinely differ.
        self.soup_df = pd.DataFrame([
            {"SBA GP Number": "1", "Servicer": "Test Bank", "Payment Month": "2026-05-01",
             "Loan Status": "ACTIVE", "Delinquency Status": "60 DAYS",
             "Beginning Actual Balance": 100000.0, "Ending Actual Balance": 100000.0},
            {"SBA GP Number": "2", "Servicer": "Test Bank", "Payment Month": "2026-05-01",
             "Loan Status": "ACTIVE", "Delinquency Status": "30 DAYS",
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

    def test_audit_returns_the_original_portfolio_summary(self):
        body = self._run_audit().json()
        self.assertIn("portfolio_summary_original_b64", body)
        self.assertIsNotNone(body["portfolio_summary_original_b64"])
        self.assertTrue(body["portfolio_summary_original_filename"].startswith("Portfolio_Summary_Original_"))

    def test_audit_still_returns_the_audited_portfolio_summary(self):
        # Option B: the Audited workbook is DEMOTED, not removed.
        body = self._run_audit().json()
        self.assertIn("portfolio_summary_audited_b64", body)
        self.assertIsNotNone(body["portfolio_summary_audited_b64"])
        self.assertTrue(body["portfolio_summary_audited_filename"].startswith("Portfolio_Summary_Audited_"))

    def test_both_workbooks_are_named_for_the_same_period(self):
        body = self._run_audit().json()
        original = body["portfolio_summary_original_filename"].replace("Portfolio_Summary_Original_", "")
        audited = body["portfolio_summary_audited_filename"].replace("Portfolio_Summary_Audited_", "")
        self.assertEqual(original, audited)

    def test_the_two_workbooks_report_different_figures(self):
        # They describe the same portfolio through different status sources, so
        # identical FIGURES would mean one is being generated from the other's
        # records. Compared as values, not bytes -- openpyxl's creation
        # timestamp would make a byte comparison pass without proving anything.
        body = self._run_audit().json()
        self.assertNotEqual(
            self._workbook_values(body["portfolio_summary_original_b64"]),
            self._workbook_values(body["portfolio_summary_audited_b64"]),
        )

    @staticmethod
    def _workbook_values(b64):
        """Every cell of every sheet, as plain values.

        Raw .xlsx bytes are NOT comparable across runs: openpyxl stamps a
        creation timestamp into docProps/core.xml, so two workbooks holding
        identical figures differ in bytes. The invariance claim is about the
        reported figures, so compare those."""
        wb = load_workbook(io.BytesIO(base64.b64decode(b64)))
        return {
            name: [[c.value for c in row] for row in wb[name].iter_rows()]
            for name in wb.sheetnames
        }

    def _original_current_dollars(self, b64):
        ws = load_workbook(io.BytesIO(base64.b64decode(b64)))["Portfolio Summary Dollars"]
        import report
        col = report.PORTFOLIO_SUMMARY_CATEGORIES.index("Current") + 2
        return ws.cell(row=2, column=col).value

    def test_original_workbook_carries_the_lender_reported_bucket(self):
        body = self._run_audit().json()
        # Lender said CURRENT for loan 1; SOUP said 60 DAYS. Original must
        # report the lender's figure.
        self.assertEqual(self._original_current_dollars(body["portfolio_summary_original_b64"]), 99000.0)

    def test_original_figures_are_unchanged_by_a_review(self):
        """The end-to-end invariance claim: run the audit, put every
        recommendation through Manual Review, re-run the audit on the SAME
        inputs, and every figure in the Original workbook must be identical.
        The Final Audited summary produced by the review is deliberately not
        compared -- it is allowed, and expected, to differ."""
        first = self._run_audit().json()
        decisions = [
            {"id": rec["id"], "decision": "approve", "reviewer_note": "confirmed"}
            for rec in first["recommendations"]
        ]
        if decisions:
            review = self.client.post("/api/corrections/review", json={
                "session_id": first["session_id"],
                "reviewed_by": "K. Dutta",
                "decisions": decisions,
            })
            self.assertEqual(review.status_code, 200)

        second = self._run_audit().json()
        self.assertEqual(
            self._workbook_values(first["portfolio_summary_original_b64"]),
            self._workbook_values(second["portfolio_summary_original_b64"]),
            "the Original Portfolio Summary must not move after Manual Review",
        )


if __name__ == "__main__":
    unittest.main()
