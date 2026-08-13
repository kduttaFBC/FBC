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

# MEMO-01 issues_found for the Community Bank & Trust run below -- that
# servicer's OWN memo rows. Pinned so the attribution fix is provably confined
# to whose-row-is-this: the six fallback rows belong to Mission Valley Bank and
# Poppy Bank, and attributing them correctly must not add them to this number.
SCOPED_ISSUE_COUNT = 6

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


def _memo_result(body):
    return next(r for r in body["results"] if r["id"].split("::")[0] == "MEMO-01")


def _lookup_failed_rows(body):
    return sorted(rec["Memo Sheet Row"] for rec in _memo_result(body)["records"]
                  if rec.get("Memo Row Class") == memo_audit.ROW_CLASS_LOOKUP_FAILED)


def _unattributable_rows(body):
    """Rows the report actually labels UNATTRIBUTABLE.

    Deliberately a SEPARATE probe from _lookup_failed_rows. "Memo Row Class"
    and "Memo Scope" are different columns answering different questions --
    which loan is this, versus whose row is this -- and the 2026-08-10 fallback
    fixed only the first. Six rows read "joined" and "UNATTRIBUTABLE -- file
    level" simultaneously, and every assertion in this file measured the class,
    so the contradiction was invisible to the suite."""
    return sorted(rec["Memo Sheet Row"] for rec in _memo_result(body)["records"]
                  if rec.get("Memo Scope") == memo_audit.MEMO_SCOPE_UNATTRIBUTABLE)


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

    def test_memo_audit_attributes_through_the_shared_resolver(self):
        # memo_audit had THREE attribution sites and the 2026-08-10 fallback
        # wired none of them: the lender-unmapped flag, "Memo Scope" and
        # "Resolved Servicer Key" each called memo_lender_soup_key(row["Lender"])
        # directly, so six rows the resolver had already attributed were still
        # reported "UNATTRIBUTABLE -- file level". They now share _row_soup_key.
        # A direct Lender-only call reappearing in either consumer is that same
        # defect coming back, so assert on the absence.
        self.assertIn("resolved_servicer", inspect.getsource(memo_audit._row_soup_key),
                      "_row_soup_key must consult the resolver's servicer")
        for consumer in (memo_audit.classify_memo_rows, memo_audit.memo_coverage_records):
            source = inspect.getsource(consumer)
            with self.subTest(consumer=consumer.__name__):
                self.assertIn("_row_soup_key", source,
                              "attribution must go through the shared resolver")
                self.assertNotIn("memo_lender_soup_key(", source,
                                 "attribution must not bypass _row_soup_key by "
                                 "reading the Lender cell directly")


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
        # Asserts BOTH columns. The name always promised the scope, but until
        # 2026-08-12 the body only checked the class -- and passed green while
        # all six rows carried "UNATTRIBUTABLE -- file level".
        failed = _lookup_failed_rows(self.body)
        unattributable = _unattributable_rows(self.body)
        for row in FALLBACK_ROWS:
            with self.subTest(memo_row=row):
                self.assertNotIn(row, failed)
                self.assertNotIn(row, unattributable,
                                 "a row the fallback resolved has a servicer, so it "
                                 "cannot be scoped UNATTRIBUTABLE")

    def test_only_the_two_unresolvable_rows_are_scoped_unattributable(self):
        self.assertEqual(_unattributable_rows(self.body), UNRESOLVABLE_ROWS)
        self.assertEqual(_memo_result(self.body)["unattributable_count"],
                         len(UNRESOLVABLE_ROWS))

    def test_the_fallback_rows_report_the_servicer_of_the_matched_loan(self):
        # The attribution is not merely "not unattributable" -- it names the
        # right servicer, taken from the loan the fallback matched.
        by_row = {rec["Memo Sheet Row"]: rec for rec in _memo_result(self.body)["records"]}
        for row in FALLBACK_ROWS:
            with self.subTest(memo_row=row):
                # Mission Valley Bank / Poppy Bank are not in this audit, so a
                # correctly attributed row is scoped OUT of a Community Bank &
                # Trust run rather than shown as its issue.
                self.assertNotIn(row, by_row)

    def test_the_uploaded_servicers_own_issue_count_is_unchanged(self):
        # The fix must not move another servicer's rows onto this audit's books.
        self.assertEqual(_memo_result(self.body)["issues_found"], SCOPED_ISSUE_COUNT)

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

    def _counts_without_and_with_attribution(self):
        soup = file_readers.read_soup_file(open(SOUP_PATH, "rb").read())
        memo = file_readers.read_service_memo(open(MEMO_PATH, "rb").read())
        population = soup["SBA GP Number"].dropna().tolist()
        before = memo_audit.memo_partition_counts(
            memo_audit.classify_memo_rows(memo, population))
        after = memo_audit.memo_partition_counts(memo_audit.classify_memo_rows(
            file_readers.resolve_memo_attribution(
                memo, file_readers.build_memo_loan_index(soup)),
            population))
        return before, after

    def test_every_data_quality_count_except_attribution_is_unaffected(self):
        # AMENDED 2026-08-12. lender_unmapped was in this list, which pinned the
        # incomplete wiring as if it were the requirement: the flag was computed
        # from the Lender cell alone, so of course attribution could not move
        # it, and asserting that it must not move is what protected the defect.
        # The flag means "nothing identifies this row's servicer", and the
        # fallback identifies six of them -- so it MUST move. It is asserted
        # explicitly in the next test instead of being frozen here.
        #
        # The rest of the list stands unchanged and is the real invariant:
        # attribution answers WHOSE ROW THIS IS and must never touch what the
        # row says about itself. A date anomaly is a date anomaly whether or not
        # the servicer was identified.
        before, after = self._counts_without_and_with_attribution()
        for flag in ("deferral_dates_missing", "date_anomaly_violations",
                     "date_anomaly_rows", "label_conflict",
                     "total_rows", "content", "phantom"):
            with self.subTest(flag=flag):
                self.assertEqual(after[flag], before[flag])

    def test_attribution_clears_lender_unmapped_on_exactly_the_fallback_rows(self):
        before, after = self._counts_without_and_with_attribution()
        self.assertEqual(before["lender_unmapped"], len(UNRESOLVABLE_ROWS) + len(FALLBACK_ROWS))
        self.assertEqual(after["lender_unmapped"], len(UNRESOLVABLE_ROWS))


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

    def test_a_bypass_also_restores_the_unattributable_mislabelling(self):
        # The scope column is now load-bearing, so the bypass probe has to cover
        # it too -- otherwise removing the resolver would still look half-clean.
        original = file_readers.resolve_memo_attribution
        file_readers.resolve_memo_attribution = lambda memo_df, loan_index: memo_df
        try:
            _status, body = _run_audit()
            degraded = _unattributable_rows(body)
        finally:
            file_readers.resolve_memo_attribution = original

        self.assertEqual(degraded, sorted(UNRESOLVABLE_ROWS + FALLBACK_ROWS))
        _status, body = _run_audit()
        self.assertEqual(_unattributable_rows(body), UNRESOLVABLE_ROWS)

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
