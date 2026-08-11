"""PORT-01 is hidden from the on-screen Check results table but preserved
internally (client-confirmed 2026-08-11).

THE CHANGE
----------
PORT-01 (Original Portfolio Summary) now has its own downloadable workbook, so
a duplicate on-screen row reading "INFO / 0 issues" per servicer -- eight rows
on the June set -- adds nothing a reader can act on. It is hidden from the
Check results TABLE only.

WHAT MUST NOT CHANGE, and is asserted here:

  * PORT-01 stays in /api/audit's `results` payload. main.py builds the
    Original Portfolio Summary workbook FROM that record, and this app's
    stated contract (main.py _filter_report_results) is that every check's
    complete result is always in the JSON regardless of what any client-facing
    view chooses to show. A hidden row must never be read as a check that did
    not run.
  * The Original Portfolio Summary workbook is still produced.
  * PORT-01's calculation is untouched.
  * PORT-01 is still absent from the reconciliation workbook tabs -- it has
    been in REPORT_ALWAYS_EXCLUDED_CHECK_IDS since long before this change,
    so the workbook half of "remove PORT-01 from the client-facing report"
    required no code change at all. Pinned here so that stays true.

KNOWN, DELIBERATE CONSEQUENCE
-----------------------------
PORT-01 declares severity="INFO", but checks._result sets
status = "PASS" if issues == 0 else severity, and PORT-01 always reports 0
issues -- so it reaches the client as PASS and is counted by summarize()'s
`checks_passed`, one of the four scorecards on the results screen.

Hiding the row therefore leaves the scorecard counting a check that no longer
has a visible row (eight of them on the June set). That is a display-only
inconsistency, not a wrong figure: every summary number is exactly what it was
before this change, which test_hiding_the_row_did_not_change_the_scorecards
asserts. Excluding PORT-01 from `checks_passed` would change a client-facing
count and was NOT approved as part of this change; it is pinned here so the
decision is explicit rather than accidental.

FRONTEND COVERAGE CAVEAT
------------------------
This repository has no JavaScript test runner (no vitest/jest, no test script,
no test files anywhere under frontend/). The hiding rule is therefore asserted
at SOURCE level below: ResultsScreen.jsx must route the table through
`visibleResults`, and `visibleResults` must exclude PORT-01. That is weaker
than rendering the component and inspecting the DOM. The rule was written as a
pure exported function (`visibleResults`) specifically so it can be unit-tested
directly the moment a runner is added.

Run with:
    python -m pytest test_port01_hidden_from_results_table.py
"""

import io
import os
import re
import unittest

import pandas as pd
from fastapi.testclient import TestClient
from openpyxl import Workbook

import file_readers
import main


HEADER_ROW = 9


def _srf_workbook_bytes(rows):
    """A real workbook shaped like the SRF. Required because preflight.py runs
    on RAW BYTES before Calamine and fails a non-workbook upload (FORM-04),
    so a placeholder byte string never reaches the audit at all."""
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


RESULTS_SCREEN = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "frontend", "src", "components", "ResultsScreen.jsx",
)


def _srf(rows):
    return pd.DataFrame([{
        "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}", "Loan Status": ls,
        "Delinquency Status": dq, "Beginning Balance": bb, "Ending Balance": eb,
        "Payment Received": rec, "Scheduled Payment": 500.0, "Servicing Fee": 0.0,
    } for gp, ls, dq, bb, eb, rec in rows])


class Port01StillInTheApiResponseTests(unittest.TestCase):
    """The backend contract is UNCHANGED. This is the half of the change that
    could silently break the Original Portfolio Summary, so it is asserted
    behaviourally against a real /api/audit call."""

    def setUp(self):
        self.file_bytes = _srf_workbook_bytes([
            ("1", "L-1", "ACTIVE", "CURRENT", 500.0, 99000.0, 100000.0, 500.0),
            ("2", "L-2", "ACTIVE", "30 DAYS", 0.0, 50000.0, 50000.0, 500.0),
        ])
        self.servicer_inputs_df = _srf([
            ("1", "ACTIVE", "CURRENT", 100000.0, 99000.0, 500.0),
            ("2", "ACTIVE", "30 DAYS", 50000.0, 50000.0, 0.0),
        ])
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

    def _audit(self):
        return self.client.post(
            "/api/audit",
            files={
                "soup_file": ("soup.xlsx", b"dummy soup bytes", "application/octet-stream"),
                "servicer_files": ("servicer.xlsx", self.file_bytes, "application/octet-stream"),
            },
        ).json()

    def test_port01_is_present_in_the_api_results(self):
        body = self._audit()
        ids = [r["id"].split("::")[0] for r in body["results"]]
        self.assertIn("PORT-01", ids, "PORT-01 must remain in the API response for traceability")

    def test_port01_result_still_carries_its_full_record(self):
        body = self._audit()
        port01 = next(r for r in body["results"] if r["id"].split("::")[0] == "PORT-01")
        self.assertEqual(port01["name"], "Original Portfolio Summary")
        self.assertEqual(port01["issues_found"], 0)
        self.assertEqual(len(port01["records"]), 1)
        # The lender said CURRENT for loan 1 -- unchanged calculation.
        self.assertEqual(port01["records"][0]["Current ($)"], 99000.0)

    def test_port01_surfaces_as_pass_not_info(self):
        """PORT-01 declares severity="INFO", but checks._result derives
        status = "PASS" if issues == 0 else severity, and PORT-01 always
        reports 0 issues. So it reaches the API as PASS.

        This is pinned because it has a CLIENT-VISIBLE consequence that
        outlives this change: PORT-01 is counted by summarize()'s
        `checks_passed`, which IS one of the four scorecards on the results
        screen -- while its row is no longer in the table below. See this
        module's docstring."""
        body = self._audit()
        port01 = next(r for r in body["results"] if r["id"].split("::")[0] == "PORT-01")
        self.assertEqual(port01["severity"], "INFO")
        self.assertEqual(port01["status"], "PASS")

    def test_hiding_the_row_did_not_change_the_scorecards(self):
        """The change is display-only, so every summary figure must be exactly
        what it was before. Asserted against the results list itself rather
        than hardcoded totals, so it stays true as the fixture evolves."""
        body = self._audit()
        results, summary = body["results"], body["summary"]
        self.assertEqual(summary["total_checks"], len(results))
        self.assertEqual(
            summary["checks_passed"], sum(1 for r in results if r["status"] == "PASS")
        )
        self.assertEqual(
            summary["total_issues"], sum(r["issues_found"] for r in results)
        )

    def test_the_original_portfolio_summary_workbook_is_still_produced(self):
        # The row and the workbook are produced from the SAME record, so a
        # change that removed PORT-01 from the backend would take the workbook
        # with it. This is the guard against that.
        body = self._audit()
        self.assertIsNotNone(body["portfolio_summary_original_b64"])
        self.assertTrue(body["portfolio_summary_original_filename"].startswith("Portfolio_Summary_Original_"))

    def test_audited_workbook_is_unaffected(self):
        body = self._audit()
        self.assertIsNotNone(body["portfolio_summary_audited_b64"])


class Port01StillExcludedFromWorkbookTabsTests(unittest.TestCase):
    """The reconciliation workbook already excluded PORT-01 before this change
    and must continue to. Pinned so the two exclusions cannot drift apart."""

    def test_port01_is_in_the_always_excluded_tab_list(self):
        self.assertIn("PORT-01", main.REPORT_ALWAYS_EXCLUDED_CHECK_IDS)

    def test_port01_gets_no_worksheet_even_with_findings(self):
        results = [{"id": "PORT-01::Test Bank", "issues_found": 5, "status": "INFO"}]
        self.assertEqual(main._filter_report_results(results), [])


class ResultsTableHidesPort01Tests(unittest.TestCase):
    """Source-level assertion of the on-screen hiding rule -- see this module's
    FRONTEND COVERAGE CAVEAT. Deliberately asserts STRUCTURE (the table routes
    through visibleResults; visibleResults names PORT-01) rather than matching
    a formatted source line, so reformatting does not break it."""

    @classmethod
    def setUpClass(cls):
        with open(RESULTS_SCREEN, encoding="utf-8") as fh:
            cls.src = fh.read()

    def test_the_component_file_exists_where_expected(self):
        self.assertTrue(os.path.isfile(RESULTS_SCREEN), RESULTS_SCREEN)

    def test_a_hidden_checks_list_names_port01(self):
        m = re.search(
            r"CHECKS_HIDDEN_FROM_RESULTS_TABLE\s*=\s*\[(?P<body>[^\]]*)\]", self.src, re.S
        )
        self.assertIsNotNone(m, "expected a CHECKS_HIDDEN_FROM_RESULTS_TABLE list")
        self.assertIn("PORT-01", m.group("body"))

    def test_visible_results_is_defined_and_filters_that_list(self):
        m = re.search(
            r"export\s+function\s+visibleResults\s*\([^)]*\)\s*\{(?P<body>.*?)\n\}", self.src, re.S
        )
        self.assertIsNotNone(m, "expected an exported visibleResults() function")
        body = m.group("body")
        self.assertIn("CHECKS_HIDDEN_FROM_RESULTS_TABLE", body)
        self.assertIn("filter", body)

    def test_the_check_results_table_maps_over_visible_results(self):
        # The table must not map over the raw results array.
        self.assertIn("visibleResults(results).map(", self.src)
        self.assertNotIn("{results.map(", self.src)

    def test_the_hiding_rule_is_documented_as_display_only(self):
        # The comment is the thing that stops a future reader concluding the
        # check stopped running. Asserted so it cannot be dropped silently.
        self.assertIn("DISPLAY ONLY", self.src)


if __name__ == "__main__":
    unittest.main()
