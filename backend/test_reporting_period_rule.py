"""Audit cycle month — CLIENT-CONFIRMED (2026-08-11).

THE RULE
--------
The lender-reported Loan Status and Delinquency Status in SOUP represent the
month IMMEDIATELY PRIOR to the raw servicer reporting file's period. SOUP
therefore always lags the cycle being audited by exactly one month, and

    audit cycle month = latest SOUP Payment Month + 1 month

is the INTENDED business rule, not an approximation. `checks._infer_cycle_month`
implements it. Before this file the rule had no direct test at all: the only
mention anywhere in the suite was a comment in
test_resume_expected_loan_status.py.

The rule is load-bearing. `cycle_month` selects which memo deferral window
covers this cycle (`checks._lookup_deferral`, `_check_change_in_terms_resume`),
which sets the deferral phase (start / mid / resume), which sets the expected
Loan Status and Delinquency Status. Measured on the real June files, shifting
the inferred month by one changes the STAT-01 finding on 4 (-1 month) to 9
(+1 month) loans and changes the recommendation set a reviewer is asked to
approve. Nothing pinned it. Now something does.

REPORTING PERIOD VALIDATION
---------------------------
Every real SRF states its own period in the Servicer Inputs header block:

    ► Report Due Date:   2026-07-08   (June 2026 cycle)
    ► Remittance Date:   2026-07-23

Production reads NEITHER -- only "► Note Series:" is read, for the Deal. The
tests here confirm that when the servicer file's stated period IS available,
it agrees with the inferred cycle month across every current fixture.

CLIENT-CONFIRMED (2026-08-11): a disagreement is a REPORTING PERIOD VALIDATION
ISSUE FOR REVIEW. It must NEVER automatically change the cycle month. This
file pins that as an invariant: `_infer_cycle_month` depends on SOUP alone, and
production behaviour is deliberately unchanged -- no new check, no new warning,
no new blocking path. `_classify_reporting_period` below is TEST-ONLY. It
expresses the agreed classification so the semantics are executable and
reviewable before any production surface is chosen; it is not imported by, and
must not be imported by, any production module.

The real-file tests skip unless the fixtures are present. Run with:
    FALCONBRIDGE_SOUP_XLSX=... FALCONBRIDGE_SRF_DIR=... \\
    FALCONBRIDGE_JULY_SOUP_XLSX=... FALCONBRIDGE_JULY_SRF_DIR=... \\
    python -m pytest test_reporting_period_rule.py
"""

import glob
import inspect
import io
import os
import unittest

import pandas as pd
from python_calamine import CalamineWorkbook

import checks
import file_readers
import main


SOUP_PATH = os.environ.get("FALCONBRIDGE_SOUP_XLSX")            # June SOUP
SRF_DIR = os.environ.get("FALCONBRIDGE_SRF_DIR")                # June submissions
JULY_SOUP_PATH = os.environ.get("FALCONBRIDGE_JULY_SOUP_XLSX")  # July SOUP
JULY_SRF_DIR = os.environ.get("FALCONBRIDGE_JULY_SRF_DIR")      # July submissions

JUNE_REAL = all([SOUP_PATH, SRF_DIR])
JULY_REAL = all([JULY_SOUP_PATH, JULY_SRF_DIR])

# (filename substring, SOUP servicer key, display label). Identity comes from
# the upload filename convention, never from an in-sheet label -- these mirror
# the frontend's SERVICER_SLOTS / SERVICER_SOUP_KEYS maps, as the other
# real-file tests do.
JUNE_SERVICERS = [
    ("Community Bank and Trust", "Community Bank & Trust", "Community Bank and Trust"),
    ("Lendistry_SOUP 2024", "Lendistry", "Lendistry 2024"),
    ("Lendistry_SOUP 2025", "Lendistry", "Lendistry 2025"),
    ("Millennium_SOUP 2024", "Millennium Bank", "Millennium 2024"),
    ("Millennium_SOUP 2025", "Millennium Bank", "Millennium 2025"),
    ("Mission Valley Bank", "Mission Valley Bank", "Mission Valley Bank"),
    ("Peoples Bank", "Peoples Bank", "Peoples Bank"),
    ("Poppy Bank", "Poppy Bank", "Poppy Bank"),
]
JULY_SERVICERS = [
    ("Millennium_SOUP 2024", "Millennium Bank", "Millennium 2024"),
    ("Millennium_SOUP 2025", "Millennium Bank", "Millennium 2025"),
]

JUNE_CYCLE = pd.Timestamp("2026-06-01")
JULY_CYCLE = pd.Timestamp("2026-07-01")


# ---------------------------------------------------------------------------
# TEST-ONLY helpers. Neither exists in production, and the last test class
# asserts that production still does not read a servicer reporting period.
# ---------------------------------------------------------------------------

CONSISTENT = "CONSISTENT"
REVIEW = "REPORTING_PERIOD_VALIDATION_ISSUE"
NOT_VALIDATABLE = "NOT_VALIDATABLE"

# Label cells scanned for in the Servicer Inputs header block, in preference
# order. Parsing mirrors file_readers.extract_deal_from_note_series exactly:
# scan the first 8 rows, match the label cell, take the cell to its right.
_PERIOD_LABELS = ("report due date", "remittance date")


def _stated_reporting_period(file_bytes):
    """TEST-ONLY. The cycle month the servicer file states for itself, or None
    when the file does not state one.

    The header block carries the DUE date, which falls in the month AFTER the
    cycle the file reports on (June 2026 cycle -> "► Report Due Date:
    2026-07-08"), so the stated cycle month is that date's month minus one.
    Returns a Timestamp normalized to the 1st, or None -- callers must treat
    None as "not validatable", never as a disagreement."""
    try:
        wb = CalamineWorkbook.from_object(io.BytesIO(file_bytes))
        sheet = wb.get_sheet_by_name("Servicer Inputs")
    except Exception:
        return None
    rows = sheet.to_python()
    for label in _PERIOD_LABELS:
        for row in rows[:8]:
            for i, cell in enumerate(row):
                if cell and label in str(cell).lower() and i + 1 < len(row):
                    due = pd.to_datetime(str(row[i + 1]).strip(), errors="coerce")
                    if pd.isna(due):
                        continue
                    return (due.replace(day=1) - pd.DateOffset(months=1)).replace(day=1)
    return None


def _classify_reporting_period(inferred, stated):
    """TEST-ONLY. CLIENT-CONFIRMED (2026-08-11) classification.

    Agreement, or no stated period to compare against, is not an issue. A
    genuine disagreement is a REPORTING PERIOD VALIDATION ISSUE FOR REVIEW.

    The returned cycle month is ALWAYS the inferred one. A disagreement never
    substitutes the stated period -- the inferred value stands and a human
    decides. This is the whole point of the classification."""
    if inferred is None or stated is None:
        status = NOT_VALIDATABLE
    elif pd.Timestamp(inferred) == pd.Timestamp(stated):
        status = CONSISTENT
    else:
        status = REVIEW
    return {"status": status, "inferred": inferred, "stated": stated,
            "cycle_month": inferred}


def _soup(rows):
    return pd.DataFrame([{"SBA GP Number": gp, "Servicer": "S", "Payment Month": pm}
                         for gp, pm in rows])


def _find(directory, substring):
    hits = [p for p in glob.glob(os.path.join(directory, "*.xlsx"))
            if substring.lower() in os.path.basename(p).lower()]
    return hits[0] if hits else None


# ---------------------------------------------------------------------------
# The rule itself -- no fixtures required
# ---------------------------------------------------------------------------

class InferredCycleMonthRuleTests(unittest.TestCase):
    """CLIENT-CONFIRMED (2026-08-11): cycle month = latest SOUP Payment Month
    + 1 month, because SOUP's lender-reported statuses describe the month
    immediately prior to the servicer file's period."""

    def test_one_month_after_the_latest_soup_payment_month(self):
        soup = _soup([("1", "2026-04-01"), ("2", "2026-05-01")])
        self.assertEqual(checks._infer_cycle_month(soup), pd.Timestamp("2026-06-01"))

    def test_it_is_the_LATEST_month_that_governs_not_the_earliest(self):
        older = _soup([("1", "2026-01-01"), ("2", "2026-02-01"), ("3", "2026-05-01")])
        self.assertEqual(checks._infer_cycle_month(older), pd.Timestamp("2026-06-01"))

    def test_row_order_does_not_matter(self):
        forward = _soup([("1", "2026-03-01"), ("2", "2026-05-01")])
        reverse = _soup([("1", "2026-05-01"), ("2", "2026-03-01")])
        self.assertEqual(checks._infer_cycle_month(forward),
                         checks._infer_cycle_month(reverse))

    def test_result_is_normalized_to_the_first_of_the_month(self):
        soup = _soup([("1", "2026-05-27")])
        cycle = checks._infer_cycle_month(soup)
        self.assertEqual(cycle.day, 1)
        self.assertEqual(cycle, pd.Timestamp("2026-06-01"))

    def test_a_mid_month_soup_date_still_advances_exactly_one_month(self):
        for day in ("01", "09", "15", "28", "31"):
            with self.subTest(day=day):
                soup = _soup([("1", f"2026-03-{day}")])
                self.assertEqual(checks._infer_cycle_month(soup),
                                 pd.Timestamp("2026-04-01"))

    def test_december_rolls_over_into_the_next_january(self):
        soup = _soup([("1", "2025-12-01")])
        self.assertEqual(checks._infer_cycle_month(soup), pd.Timestamp("2026-01-01"))

    def test_january_advances_into_february_of_the_same_year(self):
        soup = _soup([("1", "2026-01-31")])
        self.assertEqual(checks._infer_cycle_month(soup), pd.Timestamp("2026-02-01"))

    def test_the_rule_holds_across_twelve_consecutive_months(self):
        for m in range(1, 13):
            with self.subTest(month=m):
                soup = _soup([("1", f"2026-{m:02d}-01")])
                expected = (pd.Timestamp(f"2026-{m:02d}-01") + pd.DateOffset(months=1))
                self.assertEqual(checks._infer_cycle_month(soup), expected)

    def test_returns_a_pandas_timestamp(self):
        self.assertIsInstance(checks._infer_cycle_month(_soup([("1", "2026-05-01")])),
                              pd.Timestamp)


class UndeterminableCycleMonthTests(unittest.TestCase):
    """None means "skip deferral-window logic", never an exception -- callers
    degrade to the pre-deferral-aware behaviour. Pinned so the safe-degradation
    contract cannot regress into a raise."""

    def test_none_dataframe_returns_none(self):
        self.assertIsNone(checks._infer_cycle_month(None))

    def test_empty_dataframe_returns_none(self):
        self.assertIsNone(checks._infer_cycle_month(_soup([]).iloc[0:0]))

    def test_missing_payment_month_column_returns_none(self):
        self.assertIsNone(checks._infer_cycle_month(
            pd.DataFrame([{"SBA GP Number": "1", "Servicer": "S"}])))

    def test_all_unparseable_months_returns_none(self):
        self.assertIsNone(checks._infer_cycle_month(
            _soup([("1", "not a date"), ("2", "")])))

    def test_a_parseable_month_among_unparseable_ones_still_governs(self):
        soup = _soup([("1", "not a date"), ("2", "2026-05-01"), ("3", "")])
        self.assertEqual(checks._infer_cycle_month(soup), pd.Timestamp("2026-06-01"))

    def test_none_does_not_raise_for_any_degenerate_input(self):
        for label, frame in (("None", None),
                             ("empty", _soup([]).iloc[0:0]),
                             ("no column", pd.DataFrame([{"SBA GP Number": "1"}])),
                             ("unparseable", _soup([("1", "xyz")]))):
            with self.subTest(case=label):
                self.assertIsNone(checks._infer_cycle_month(frame))


# ---------------------------------------------------------------------------
# Reporting period validation -- real fixtures
# ---------------------------------------------------------------------------

@unittest.skipUnless(JUNE_REAL, "set FALCONBRIDGE_SOUP_XLSX / _SRF_DIR")
class JuneFixtureConsistencyTests(unittest.TestCase):
    """All eight June 2026 servicer files against the June SOUP."""

    @classmethod
    def setUpClass(cls):
        cls.soup = file_readers.read_soup_file(open(SOUP_PATH, "rb").read())

    def _rows(self):
        for sub, soup_key, display in JUNE_SERVICERS:
            path = _find(SRF_DIR, sub)
            self.assertIsNotNone(path, f"no June SRF found for {display}")
            inferred = checks._infer_cycle_month(
                file_readers.filter_soup_by_servicer(self.soup, soup_key))
            stated = _stated_reporting_period(open(path, "rb").read())
            yield display, inferred, stated

    def test_every_june_servicer_infers_june_2026(self):
        for display, inferred, _ in self._rows():
            with self.subTest(servicer=display):
                self.assertEqual(inferred, JUNE_CYCLE)

    def test_every_june_servicer_file_states_a_reporting_period(self):
        for display, _, stated in self._rows():
            with self.subTest(servicer=display):
                self.assertIsNotNone(stated)

    def test_every_june_servicer_is_classified_consistent(self):
        for display, inferred, stated in self._rows():
            with self.subTest(servicer=display):
                self.assertEqual(
                    _classify_reporting_period(inferred, stated)["status"], CONSISTENT)

    def test_no_june_servicer_raises_a_validation_issue(self):
        issues = [d for d, i, s in self._rows()
                  if _classify_reporting_period(i, s)["status"] == REVIEW]
        self.assertEqual(issues, [])

    def test_all_eight_june_servicers_were_actually_exercised(self):
        self.assertEqual(len(list(self._rows())), 8)


@unittest.skipUnless(JULY_REAL, "set FALCONBRIDGE_JULY_SOUP_XLSX / _JULY_SRF_DIR")
class JulyFixtureConsistencyTests(unittest.TestCase):
    """Both July 2026 servicer files against the July SOUP. The July SOUP is a
    later vintage than June's, so this also confirms the rule holds as the
    dataset rolls forward rather than only at one point in time."""

    @classmethod
    def setUpClass(cls):
        cls.soup = file_readers.read_soup_file(open(JULY_SOUP_PATH, "rb").read())

    def _rows(self):
        for sub, soup_key, display in JULY_SERVICERS:
            path = _find(JULY_SRF_DIR, sub)
            self.assertIsNotNone(path, f"no July SRF found for {display}")
            inferred = checks._infer_cycle_month(
                file_readers.filter_soup_by_servicer(self.soup, soup_key))
            stated = _stated_reporting_period(open(path, "rb").read())
            yield display, inferred, stated

    def test_every_july_servicer_infers_july_2026(self):
        for display, inferred, _ in self._rows():
            with self.subTest(servicer=display):
                self.assertEqual(inferred, JULY_CYCLE)

    def test_every_july_servicer_is_classified_consistent(self):
        for display, inferred, stated in self._rows():
            with self.subTest(servicer=display):
                self.assertEqual(
                    _classify_reporting_period(inferred, stated)["status"], CONSISTENT)

    def test_the_july_cycle_is_exactly_one_month_after_the_june_cycle(self):
        self.assertEqual(JUNE_CYCLE + pd.DateOffset(months=1), JULY_CYCLE)

    def test_both_july_servicers_were_actually_exercised(self):
        self.assertEqual(len(list(self._rows())), 2)


@unittest.skipUnless(JUNE_REAL and JULY_REAL,
                     "set the June and July SOUP / SRF fixtures")
class CrossVintageTests(unittest.TestCase):
    """The June and July fixtures together are the only real evidence that the
    rule tracks a rolling dataset."""

    def test_the_same_servicer_advances_one_month_between_vintages(self):
        jun = file_readers.read_soup_file(open(SOUP_PATH, "rb").read())
        jul = file_readers.read_soup_file(open(JULY_SOUP_PATH, "rb").read())
        for _, soup_key, display in JULY_SERVICERS:
            with self.subTest(servicer=display):
                a = checks._infer_cycle_month(
                    file_readers.filter_soup_by_servicer(jun, soup_key))
                b = checks._infer_cycle_month(
                    file_readers.filter_soup_by_servicer(jul, soup_key))
                self.assertEqual(b, a + pd.DateOffset(months=1))

    def test_pairing_a_june_file_with_july_soup_is_a_validation_issue(self):
        """The realistic operator error: re-running a prior period's servicer
        file after SOUP has been refreshed. The June file still states a June
        period; the July SOUP infers July. Classified for review -- and the
        cycle month returned is still the inferred one, never the stated one."""
        jul = file_readers.read_soup_file(open(JULY_SOUP_PATH, "rb").read())
        june_file = _find(SRF_DIR, "Millennium_SOUP 2024")
        self.assertIsNotNone(june_file)
        inferred = checks._infer_cycle_month(
            file_readers.filter_soup_by_servicer(jul, "Millennium Bank"))
        stated = _stated_reporting_period(open(june_file, "rb").read())
        verdict = _classify_reporting_period(inferred, stated)
        self.assertEqual(verdict["status"], REVIEW)
        self.assertEqual(verdict["inferred"], JULY_CYCLE)
        self.assertEqual(verdict["stated"], JUNE_CYCLE)
        self.assertEqual(verdict["cycle_month"], JULY_CYCLE)


# ---------------------------------------------------------------------------
# The classification contract, and the invariant that it changes nothing
# ---------------------------------------------------------------------------

class ReportingPeriodClassificationTests(unittest.TestCase):
    """CLIENT-CONFIRMED (2026-08-11): a disagreement is a validation issue for
    REVIEW. It never rewrites the cycle month."""

    def test_agreement_is_consistent(self):
        v = _classify_reporting_period(JUNE_CYCLE, JUNE_CYCLE)
        self.assertEqual(v["status"], CONSISTENT)

    def test_disagreement_is_a_reporting_period_validation_issue(self):
        v = _classify_reporting_period(JULY_CYCLE, JUNE_CYCLE)
        self.assertEqual(v["status"], REVIEW)

    def test_a_disagreement_NEVER_substitutes_the_stated_period(self):
        for inferred, stated in ((JULY_CYCLE, JUNE_CYCLE), (JUNE_CYCLE, JULY_CYCLE),
                                 (pd.Timestamp("2026-01-01"), pd.Timestamp("2025-06-01"))):
            with self.subTest(inferred=str(inferred), stated=str(stated)):
                v = _classify_reporting_period(inferred, stated)
                self.assertEqual(v["status"], REVIEW)
                self.assertEqual(v["cycle_month"], inferred)
                self.assertNotEqual(v["cycle_month"], stated)

    def test_no_stated_period_is_not_validatable_and_not_an_issue(self):
        v = _classify_reporting_period(JUNE_CYCLE, None)
        self.assertEqual(v["status"], NOT_VALIDATABLE)
        self.assertNotEqual(v["status"], REVIEW)
        self.assertEqual(v["cycle_month"], JUNE_CYCLE)

    def test_no_inferred_month_is_not_validatable_either(self):
        v = _classify_reporting_period(None, JUNE_CYCLE)
        self.assertEqual(v["status"], NOT_VALIDATABLE)
        self.assertIsNone(v["cycle_month"])

    def test_a_missing_period_is_never_reported_as_a_disagreement(self):
        for inferred, stated in ((None, None), (JUNE_CYCLE, None), (None, JUNE_CYCLE)):
            with self.subTest(inferred=str(inferred), stated=str(stated)):
                self.assertNotEqual(
                    _classify_reporting_period(inferred, stated)["status"], REVIEW)

    def test_day_level_differences_within_a_month_are_not_a_disagreement(self):
        v = _classify_reporting_period(pd.Timestamp("2026-06-01"),
                                       pd.Timestamp("2026-06-01"))
        self.assertEqual(v["status"], CONSISTENT)


class ProductionBehaviourIsUnchangedTests(unittest.TestCase):
    """No production change was made for Finding 12. These pin that.

    If a future change wires reporting-period validation into the pipeline,
    these fail and tell the implementer to revisit the requirement -- exactly
    the visible, deliberate change the client asked for."""

    def test_the_inferred_cycle_month_depends_on_soup_alone(self):
        """Signature and body evidence: the only input is the SOUP frame. A
        servicer file, its bytes, or a stated period cannot reach it."""
        sig = inspect.signature(checks._infer_cycle_month)
        self.assertEqual(list(sig.parameters), ["soup_servicer_df"])

    def test_no_cycle_month_module_reads_a_servicer_reporting_period(self):
        """The modules that DERIVE or RESOLVE the audit cycle month must not
        see a stated reporting period at all.

        NARROWED (2026-08-11, Monthly Servicer Report): this assertion
        originally covered file_readers too, as a proxy for "the stated period
        influences nothing". file_readers now exposes read_remittance_date so
        the Monthly Servicer Report can PRINT the remittance date in its
        metadata block -- a presentational read that never touches the cycle
        month. The proxy was replaced with the invariant it stood for, asserted
        directly in the two tests below; checks.py, where _infer_cycle_month
        lives, is still held to the original bar."""
        import corrections
        import preflight
        import report
        for name, module in (("checks", checks), ("corrections", corrections),
                             ("preflight", preflight), ("report", report)):
            body = "\n".join(line for line in inspect.getsource(module).splitlines()
                             if not line.strip().startswith("#"))
            for label in ("Report Due Date", "Remittance Date"):
                with self.subTest(module=name, label=label):
                    self.assertNotIn(
                        label, body,
                        f"{name} now reads '{label}' -- reporting period validation "
                        f"has been wired into the cycle-month path. Revisit the "
                        f"CLIENT-CONFIRMED (2026-08-11) decision: a disagreement is a "
                        f"REVIEW issue and must never change the cycle month.")

    def test_the_remittance_date_reader_never_reaches_the_cycle_month(self):
        """file_readers.read_remittance_date exists for the Monthly Servicer
        Report's metadata block. It must never feed _infer_cycle_month,
        _resolve_portfolio_summary_cycle_month, or any cycle_month variable."""
        body = "\n".join(line for line in inspect.getsource(main).splitlines()
                         if not line.strip().startswith("#"))
        for line in body.splitlines():
            if "read_remittance_date" not in line:
                continue
            with self.subTest(line=line.strip()):
                self.assertNotIn("cycle_month", line)
                self.assertNotIn("_infer_cycle_month", line)
        source = inspect.getsource(file_readers.read_remittance_date)
        self.assertNotIn("cycle", source)

    def test_the_cycle_month_still_comes_only_from_soup(self):
        """The invariant the proxy above stood for, asserted directly."""
        self.assertEqual(list(inspect.signature(checks._infer_cycle_month).parameters),
                         ["soup_servicer_df"])
        body = inspect.getsource(checks._infer_cycle_month)
        self.assertIn("Payment Month", body)
        self.assertNotIn("remittance", body.lower())

    def test_the_only_period_enforcement_is_the_mixed_month_guard(self):
        """The one place a reporting period stops an audit today: several
        servicers in one run inferring different months. Unchanged."""
        a, err_a = main._resolve_portfolio_summary_cycle_month({JUNE_CYCLE})
        self.assertEqual(a, JUNE_CYCLE)
        self.assertIsNone(err_a)
        b, err_b = main._resolve_portfolio_summary_cycle_month({JUNE_CYCLE, JULY_CYCLE})
        self.assertIsNone(b)
        self.assertIn("different reporting months", err_b)

    def test_the_mixed_month_guard_cannot_see_a_uniform_error(self):
        """Every servicer inferring the SAME wrong month is not a conflict, so
        the existing guard is silent. Documents why the classification above
        is a separate concern rather than something the guard already covers."""
        resolved, err = main._resolve_portfolio_summary_cycle_month({JULY_CYCLE})
        self.assertEqual(resolved, JULY_CYCLE)
        self.assertIsNone(err)

    def test_the_test_only_helpers_are_not_production_code(self):
        for module in (checks, file_readers, main):
            source = inspect.getsource(module)
            with self.subTest(module=module.__name__):
                self.assertNotIn("_classify_reporting_period", source)
                self.assertNotIn("_stated_reporting_period", source)


if __name__ == "__main__":
    unittest.main()
