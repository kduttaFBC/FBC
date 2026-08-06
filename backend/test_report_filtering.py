"""
Regression tests for main._filter_report_results -- the curated-workbook
decision logic for which checks get their own tab in the detailed,
client-facing reconciliation workbook.

CLIENT-CONFIRMED (2026-08-06) decision table:
  POP-01   include only when issues exist
  PAY-01   exclude always
  PAY-02   exclude always (until validated)
  DUP-01   exclude always
  NEG-01   include only when issues exist
  PORT-01  exclude always (has its own separate workbooks)
  Every other check (BAL-01, BAL-02, STAT-01, STAT-02, DQ-01, ...): always
  included, unaffected by this filter at all.

Run with: python -m unittest test_report_filtering.py
"""

import unittest

import main


def _result(check_id, issues_found):
    return {"id": f"{check_id}::Test Bank", "issues_found": issues_found}


class AlwaysExcludedChecksTests(unittest.TestCase):
    """PAY-01, PAY-02, DUP-01, PORT-01 must never appear, regardless of
    whether they have findings."""

    def test_pay01_excluded_with_zero_issues(self):
        results = [_result("PAY-01", 0)]
        self.assertEqual(main._filter_report_results(results), [])

    def test_pay01_excluded_even_with_issues(self):
        results = [_result("PAY-01", 5)]
        self.assertEqual(main._filter_report_results(results), [])

    def test_pay02_excluded_even_with_issues(self):
        results = [_result("PAY-02", 3)]
        self.assertEqual(main._filter_report_results(results), [])

    def test_dup01_excluded_even_with_issues(self):
        # Explicitly the case the client called out: DUP-01 is
        # informational and can create noise even when it finds something
        # -- it must stay excluded from this workbook regardless.
        results = [_result("DUP-01", 10)]
        self.assertEqual(main._filter_report_results(results), [])

    def test_port01_excluded_even_with_issues(self):
        results = [_result("PORT-01", 1)]
        self.assertEqual(main._filter_report_results(results), [])


class ConditionalChecksTests(unittest.TestCase):
    """POP-01 and NEG-01 must be excluded when they have nothing to
    report, and included the moment they have real findings."""

    def test_pop01_excluded_when_zero_issues(self):
        results = [_result("POP-01", 0)]
        self.assertEqual(main._filter_report_results(results), [])

    def test_pop01_included_when_issues_exist(self):
        results = [_result("POP-01", 2)]
        filtered = main._filter_report_results(results)
        self.assertEqual(len(filtered), 1)
        self.assertEqual(filtered[0]["id"], "POP-01::Test Bank")

    def test_neg01_excluded_when_zero_issues(self):
        results = [_result("NEG-01", 0)]
        self.assertEqual(main._filter_report_results(results), [])

    def test_neg01_included_when_issues_exist(self):
        results = [_result("NEG-01", 4)]
        filtered = main._filter_report_results(results)
        self.assertEqual(len(filtered), 1)
        self.assertEqual(filtered[0]["id"], "NEG-01::Test Bank")

    def test_missing_issues_found_key_treated_as_zero(self):
        # Defensive: a result dict without an issues_found key at all
        # (shouldn't happen in practice, but must not crash or be
        # silently included) is treated as zero findings.
        results = [{"id": "POP-01::Test Bank"}]
        self.assertEqual(main._filter_report_results(results), [])


class AlwaysIncludedChecksTests(unittest.TestCase):
    """Every other check is completely unaffected by this filter, whether
    or not it has findings."""

    def test_bal01_always_included_regardless_of_issues(self):
        for issues in (0, 1, 10):
            with self.subTest(issues=issues):
                results = [_result("BAL-01", issues)]
                filtered = main._filter_report_results(results)
                self.assertEqual(len(filtered), 1)

    def test_stat01_stat02_dq01_always_included(self):
        results = [_result("STAT-01", 0), _result("STAT-02", 0), _result("DQ-01", 0)]
        filtered = main._filter_report_results(results)
        self.assertEqual(len(filtered), 3)


class FullDecisionTableTests(unittest.TestCase):
    """One combined test exercising the client's exact decision table in a
    single realistic multi-check run, both with and without findings."""

    def test_full_run_no_findings_on_conditional_checks(self):
        results = [
            _result("POP-01", 0), _result("BAL-01", 0), _result("BAL-02", 0),
            _result("STAT-01", 2), _result("STAT-02", 0), _result("PAY-01", 3),
            _result("PAY-02", 1), _result("DUP-01", 5), _result("NEG-01", 0),
            _result("PORT-01", 0), _result("DQ-01", 1),
        ]
        filtered_ids = {r["id"].split("::")[0] for r in main._filter_report_results(results)}
        self.assertEqual(filtered_ids, {"BAL-01", "BAL-02", "STAT-01", "STAT-02", "DQ-01"})

    def test_full_run_with_findings_on_conditional_checks(self):
        results = [
            _result("POP-01", 1),  # now has findings -> included
            _result("BAL-01", 0), _result("BAL-02", 0),
            _result("STAT-01", 2), _result("STAT-02", 0), _result("PAY-01", 3),
            _result("PAY-02", 1), _result("DUP-01", 5),
            _result("NEG-01", 2),  # now has findings -> included
            _result("PORT-01", 0), _result("DQ-01", 1),
        ]
        filtered_ids = {r["id"].split("::")[0] for r in main._filter_report_results(results)}
        self.assertEqual(filtered_ids, {"POP-01", "BAL-01", "BAL-02", "STAT-01", "STAT-02", "NEG-01", "DQ-01"})

    def test_missing_worksheet_is_never_evidence_a_check_failed_to_run(self):
        # A check excluded from the workbook must still be present,
        # untouched, in the ORIGINAL results list -- this filter only
        # builds a derived subset, it never mutates or removes anything
        # from the source.
        results = [_result("PAY-01", 3), _result("BAL-01", 0)]
        main._filter_report_results(results)
        self.assertEqual(len(results), 2)  # original list untouched
        self.assertEqual(results[0]["id"], "PAY-01::Test Bank")


if __name__ == "__main__":
    unittest.main()
