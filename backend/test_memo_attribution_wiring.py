"""/api/audit must actually invoke the memo attribution resolver.

The resolver is enforced by exactly ONE line in main.py:

    memo_df = file_readers.resolve_memo_attribution(
        memo_df, file_readers.build_memo_loan_index(soup_df))

Both consumers read its output CONDITIONALLY -- `if "resolved_gp" in
df.columns` in memo_audit.classify_memo_rows, and the same guard in
checks._lookup_deferral. That conditional design is deliberate: it keeps the
change additive and lets unit tests exercise the degraded path without
fixtures. The cost is that deleting the call site degrades the audit SILENTLY.
Measured: no exception, no warning, no log line -- MEMO-01 simply reverts from
198 joined / 2 lookup-failed to 192 / 8, and reports that as the truth.

Before this file, every one of the 542 existing tests passed with the call site
removed. Nothing bound the pipeline to the resolver.

WHAT THIS FILE DOES. It pins the wiring three ways -- through the real
endpoint, through a deliberate bypass that must be detected, and at the source
level -- so the call site cannot be removed or reordered without a failure.

WHAT IT DOES NOT DO. It does not implement the attribution_issue surfacing
requirement. `attribution_method` and `attribution_issue` are produced by the
resolver and read by NO production code, so a conflict or ambiguity currently
reaches no human. That is a separate requirement decision -- where such a row
should be presented to the reviewer -- and is deliberately out of scope here.
See the test at the end, which records the gap without asserting a fix.

Real-file tests skip unless the fixtures are present:
    FALCONBRIDGE_SOUP_XLSX=... FALCONBRIDGE_MEMO_XLSX=... \\
    FALCONBRIDGE_SRF_DIR=... python -m pytest test_memo_attribution_wiring.py
"""

import asyncio
import inspect
import json
import os
import re
import unittest

import file_readers
import main
import memo_audit


SOUP_PATH = os.environ.get("FALCONBRIDGE_SOUP_XLSX")
MEMO_PATH = os.environ.get("FALCONBRIDGE_MEMO_XLSX")
SRF_DIR = os.environ.get("FALCONBRIDGE_SRF_DIR")
REAL = all([SOUP_PATH, MEMO_PATH, SRF_DIR])

# The six rows the SBA Loan ID -> Lender Loan Number fallback resolves, and the
# two that match nothing under either identifier in either population.
FALLBACK_ROWS = [113, 114, 142, 143, 144, 145]
UNRESOLVABLE_ROWS = [96, 112]

# Community Bank and Trust is used as the servicer file simply because it
# passes preflight on the June drop. MEMO-01 coverage is computed from the
# whole memo against the whole SOUP population, so the choice of servicer file
# does not affect the counts under test -- only which rows are scoped into the
# audit-facing view.
SERVICER_MATCH = "community bank"
SERVICER_FILENAME = "Community Bank & Trust||Community Bank and Trust__june.xlsx"


class _Upload:
    """Minimal stand-in for FastAPI's UploadFile."""

    def __init__(self, path, filename=None):
        self._path = path
        self.filename = filename or os.path.basename(path)

    async def read(self):
        with open(self._path, "rb") as fh:
            return fh.read()


def _servicer_path():
    import glob
    return [p for p in glob.glob(os.path.join(SRF_DIR, "*.xlsx"))
            if SERVICER_MATCH in os.path.basename(p).lower()][0]


def _run_audit():
    response = asyncio.new_event_loop().run_until_complete(main.audit(
        soup_file=_Upload(SOUP_PATH),
        servicer_files=[_Upload(_servicer_path(), SERVICER_FILENAME)],
        memo_file=_Upload(MEMO_PATH),
    ))
    return response.status_code, json.loads(bytes(response.body).decode())


def _lookup_failed_rows(body):
    memo_result = next(r for r in body["results"]
                       if r["id"].split("::")[0] == "MEMO-01")
    return sorted(rec["Memo Sheet Row"] for rec in memo_result["records"]
                  if rec.get("Memo Row Class") == memo_audit.ROW_CLASS_LOOKUP_FAILED)


class SourceLevelWiringTests(unittest.TestCase):
    """Fixture-free. Catches deletion or reordering of the call site itself."""

    def test_the_audit_handler_calls_the_resolver(self):
        source = inspect.getsource(main.audit)
        self.assertIn("resolve_memo_attribution", source,
                      "/api/audit must invoke file_readers.resolve_memo_attribution")
        self.assertIn("build_memo_loan_index", source,
                      "the resolver must be given a loan index built from SOUP")

    @staticmethod
    def _executable_lines(func):
        """Source with comment lines removed -- main.audit's own comments name
        run_servicer_reconciliation before the resolver is called, so a raw
        text search finds the comment rather than the call."""
        return "\n".join(line for line in inspect.getsource(func).splitlines()
                         if not line.strip().startswith("#"))

    def test_the_resolver_runs_before_any_consumer(self):
        # Order matters: STAT-01 and MEMO-01 must both see the resolved frame,
        # otherwise they would disagree about which loan a memo row belongs to.
        source = self._executable_lines(main.audit)
        resolve_at = source.index("resolve_memo_attribution")
        reconcile_at = source.index("run_servicer_reconciliation(")
        coverage_at = source.index("build_memo_coverage_result(")
        self.assertLess(resolve_at, reconcile_at,
                        "attribution must be resolved before the checks run")
        self.assertLess(resolve_at, coverage_at,
                        "attribution must be resolved before MEMO-01 classifies rows")

    def test_the_index_is_built_from_the_soup_dataframe(self):
        source = inspect.getsource(main.audit)
        self.assertRegex(source, r"build_memo_loan_index\(\s*soup_df\s*\)")

    def test_the_consumers_still_read_the_resolved_columns(self):
        # If a consumer stops reading resolved_gp the resolver becomes dead
        # weight and coverage silently regresses.
        self.assertIn("resolved_gp", inspect.getsource(memo_audit.classify_memo_rows))
        import checks
        self.assertIn("resolved_gp", inspect.getsource(checks._lookup_deferral))
        self.assertIn("resolved_servicer", inspect.getsource(checks._lookup_deferral))


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class EndpointAttributionTests(unittest.TestCase):
    """The behaviour that must survive: 198 matched, 2 unresolved."""

    @classmethod
    def setUpClass(cls):
        cls.status, cls.body = _run_audit()

    def test_the_audit_succeeds(self):
        self.assertEqual(self.status, 200)

    def test_only_the_two_genuinely_unresolvable_rows_remain_lookup_failed(self):
        self.assertEqual(_lookup_failed_rows(self.body), UNRESOLVABLE_ROWS)

    def test_none_of_the_six_fallback_rows_is_reported_unattributable(self):
        failed = _lookup_failed_rows(self.body)
        for row in FALLBACK_ROWS:
            with self.subTest(memo_row=row):
                self.assertNotIn(row, failed)

    def test_the_resolver_produces_198_matched_and_2_unresolved(self):
        # The same population MEMO-01 is computed against, asserted directly on
        # the resolver so the headline numbers are pinned as numbers.
        soup = file_readers.read_soup_file(open(SOUP_PATH, "rb").read())
        memo = file_readers.read_service_memo(open(MEMO_PATH, "rb").read())
        resolved = file_readers.resolve_memo_attribution(
            memo, file_readers.build_memo_loan_index(soup))
        counts = memo_audit.memo_partition_counts(
            memo_audit.classify_memo_rows(
                resolved, soup["SBA GP Number"].dropna().tolist()))
        self.assertEqual((counts["joined"], counts["lookup_failed"]), (198, 2))

    def test_every_memo_data_quality_count_is_unaffected_by_attribution(self):
        soup = file_readers.read_soup_file(open(SOUP_PATH, "rb").read())
        memo = file_readers.read_service_memo(open(MEMO_PATH, "rb").read())
        before = memo_audit.memo_partition_counts(memo_audit.classify_memo_rows(
            memo, soup["SBA GP Number"].dropna().tolist()))
        after = memo_audit.memo_partition_counts(memo_audit.classify_memo_rows(
            file_readers.resolve_memo_attribution(
                memo, file_readers.build_memo_loan_index(soup)),
            soup["SBA GP Number"].dropna().tolist()))
        for flag in ("deferral_dates_missing", "date_anomaly_violations",
                     "date_anomaly_rows", "label_conflict", "lender_unmapped",
                     "total_rows", "content", "phantom"):
            with self.subTest(flag=flag):
                self.assertEqual(after[flag], before[flag])


@unittest.skipUnless(REAL, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class BypassIsDetectedTests(unittest.TestCase):
    """Proves the assertions above are actually sensitive to the bypass.

    A regression test that cannot fail is worth nothing. This replaces the
    resolver with a pass-through -- exactly what deleting the call site in
    main.py would achieve -- and asserts the endpoint output visibly degrades.
    Measured: lookup-failed goes from [96, 112] to all eight rows.
    """

    def test_removing_the_resolver_degrades_the_audit_and_is_observable(self):
        original = file_readers.resolve_memo_attribution
        file_readers.resolve_memo_attribution = lambda memo_df, loan_index: memo_df
        try:
            _status, body = _run_audit()
            degraded = _lookup_failed_rows(body)
        finally:
            file_readers.resolve_memo_attribution = original

        self.assertEqual(degraded, sorted(UNRESOLVABLE_ROWS + FALLBACK_ROWS),
                         "a bypass must lose exactly the six fallback-resolved rows")
        self.assertNotEqual(degraded, UNRESOLVABLE_ROWS)

        # ...and the wiring is intact again afterwards.
        _status, body = _run_audit()
        self.assertEqual(_lookup_failed_rows(body), UNRESOLVABLE_ROWS)

    def test_a_bypass_raises_nothing_and_warns_nothing(self):
        # Records WHY this file exists: the degradation is completely silent, so
        # only an explicit assertion can catch it.
        soup = file_readers.read_soup_file(open(SOUP_PATH, "rb").read())
        raw_memo = file_readers.read_service_memo(open(MEMO_PATH, "rb").read())
        for column in file_readers.MEMO_ATTRIBUTION_COLUMNS:
            self.assertNotIn(column, raw_memo.columns)
        classified = memo_audit.classify_memo_rows(
            raw_memo, soup["SBA GP Number"].dropna().tolist())      # no exception
        self.assertEqual(
            memo_audit.memo_partition_counts(classified)["lookup_failed"], 8)


class AttributionIssueSurfacingIsNotImplementedTests(unittest.TestCase):
    """OPEN CLIENT BUSINESS REQUIREMENT -- not an application defect.

    The resolver produces `attribution_method` and `attribution_issue`. Neither
    is read by any production module -- checks.py, memo_audit.py, main.py,
    corrections.py and report.py contain zero read sites. A row flagged
    "conflict" or "ambiguous" is therefore computed correctly and then
    discarded: it reaches no check, no report tab, no API field and no screen.

    This is NOT a defect. The resolver behaves exactly as specified: probe 3
    was scoped and approved as attribution only, and no requirement has ever
    stated where an ambiguous or conflicting memo row should be presented. The
    delivered June memo contains zero such rows, so no reviewer is currently
    being deprived of information.

    What is genuinely open is a CLIENT DECISION: when one memo row could belong
    to more than one loan (ambiguous), or when an SBA Loan ID matches one loan
    by GP Number and a different loan by Lender Loan Number (conflict), who
    should see it and where -- a MEMO-01 record field beside "Resolved Servicer
    Key", a Data Quality Flags entry, an API field, or the correction review
    page. That choice determines the implementation, so it must be decided
    before any code is written rather than guessed at here.

    These tests pin the CURRENT state so that implementing the surfacing later
    is a visible, deliberate change rather than a silent one."""

    def test_the_resolver_still_produces_both_fields(self):
        self.assertIn("attribution_method", file_readers.MEMO_ATTRIBUTION_COLUMNS)
        self.assertIn("attribution_issue", file_readers.MEMO_ATTRIBUTION_COLUMNS)
        self.assertEqual(file_readers.ATTRIBUTION_ISSUE_CONFLICT, "conflict")
        self.assertEqual(file_readers.ATTRIBUTION_ISSUE_AMBIGUOUS, "ambiguous")

    def test_no_production_module_reads_them_yet(self):
        import checks
        import corrections
        import report
        sources = {
            "checks": inspect.getsource(checks),
            "memo_audit": inspect.getsource(memo_audit),
            "main": inspect.getsource(main),
            "corrections": inspect.getsource(corrections),
            "report": inspect.getsource(report),
        }
        for module, text in sources.items():
            body = "\n".join(line for line in text.splitlines()
                             if not line.strip().startswith("#"))
            for field in ("attribution_issue", "attribution_method"):
                with self.subTest(module=module, field=field):
                    self.assertNotIn(
                        field, body,
                        f"{module} now reads {field} -- the surfacing requirement has "
                        f"been implemented; update this test and remove the gap note")

    def test_memo_01_records_carry_no_attribution_provenance(self):
        # Asserted on the EMITTED FIELD NAMES, not the source text: the
        # function's docstring legitimately uses the word "attribution" to
        # describe Lender-identity scoping, which is a different concept.
        # MEMO-01 is the natural home for this if the decision lands there.
        import pandas as pd
        memo = pd.DataFrame([{
            "Borrower Name": "T", "SBA Loan ID": "9999999999", "Lender Loan ID": "",
            "Securitization Action": "Deferred", "Servicer Action": "3 month deferment",
            "Deferral Start Date": pd.Timestamp("2026-05-01"),
            "Deferral End Date": pd.Timestamp("2026-03-01"),      # forces a record
            "Regular Payment Resume Date": pd.NaT, "Lender": "",
        }])
        memo["gp_sba"] = memo["SBA Loan ID"].apply(file_readers.normalize_memo_loan_id)
        memo["gp_lender"] = memo["Lender Loan ID"].apply(file_readers.normalize_memo_loan_id)
        resolved = file_readers.resolve_memo_attribution(
            memo, file_readers.build_memo_loan_index(
                pd.DataFrame([{"SBA GP Number": "1", "Lender Loan Number": "2",
                               "Servicer": "Bank A"}])))
        records = memo_audit.memo_coverage_records(
            memo_audit.classify_memo_rows(resolved, ["1"]))
        self.assertTrue(records, "expected at least one MEMO-01 record")
        emitted = set(records[0])
        for field in file_readers.MEMO_ATTRIBUTION_COLUMNS:
            with self.subTest(field=field):
                self.assertNotIn(field, emitted)
        self.assertFalse([k for k in emitted if "attribution" in k.lower()])


if __name__ == "__main__":
    unittest.main()
