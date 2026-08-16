"""SOUP 2025 is the validated business scope — CLIENT-CONFIRMED (2026-08-13).

WHY THIS FILE EXISTS
--------------------
The Monthly Servicer Report is DEAL-PARTITIONED, not deal-restricted: it emits
one workbook per SOUP deal present in the upload. Read the code alone and it
looks like a universal cross-deal report. It is not. Only SOUP 2025 has been
reconciled against a client-supplied workbook and given an acceptance baseline.

So there are two claims to protect, and they pull in opposite directions:

  1. SOUP 2025 is the VALIDATED scope -- 1,436 loans, $2,805,017.89, three
     lenders. A future change must not silently move those figures.
  2. SOUP 2024 capability is RETAINED. The scope is a statement about what has
     been validated, NOT a licence to delete the per-deal architecture. A
     future "simplification" that hardcodes 2025 would break a client-confirmed
     (2026-08-11) design decision.

And one rule that guards both: deal membership is resolved from each SRF's own
"Note Series:" cell by file_readers.extract_deal_from_note_series. No lender
name participates in that routing. The three names below are the OUTCOME for
June 2026, never an input -- so onboarding a fourth 2025 lender requires no
code change at all.

FIXTURES -- THE TWO POPULATIONS ARE NOT INTERCHANGEABLE
-------------------------------------------------------
FALCONBRIDGE_ACCEPTED_SRF_DIR  the client's FINAL corrected submissions
                               (2026.06/Final). THE acceptance baseline: 1,436
                               loans, $2,805,017.89, 3 recommendations.
FALCONBRIDGE_SRF_DIR           the RAW lender submissions. A stress population
                               only -- 95 recommendations, and it does NOT
                               reconcile (156 CB&T loans, $2,805,017.90).

Deliberately separate variables. They differ by one loan and one cent, which is
small enough to look like noise and large enough to invalidate an acceptance
claim, so the two must never be silently swapped.

    FALCONBRIDGE_SOUP_XLSX=... FALCONBRIDGE_MEMO_XLSX=... \\
    FALCONBRIDGE_ACCEPTED_SRF_DIR=... FALCONBRIDGE_SRF_DIR=... \\
    python3 -m pytest test_soup2025_scope.py
"""

import glob
import inspect
import os
import unittest

import checks
import corrections
import file_readers
import main
import monthly_servicer_report as msr


SOUP_PATH = os.environ.get("FALCONBRIDGE_SOUP_XLSX")
MEMO_PATH = os.environ.get("FALCONBRIDGE_MEMO_XLSX")
ACCEPTED_DIR = os.environ.get("FALCONBRIDGE_ACCEPTED_SRF_DIR")
RAW_DIR = os.environ.get("FALCONBRIDGE_SRF_DIR")

ACCEPTED = all([SOUP_PATH, MEMO_PATH, ACCEPTED_DIR])
RAW = all([SOUP_PATH, MEMO_PATH, RAW_DIR])

DEAL_2025 = "SOUP2025_1"
DEAL_2024 = "SOUP2024_1"

# (filename substring, soup_key, display_label) — the substring is a file
# locator for the fixture directory, never an input to deal routing.
LENDERS_2025 = [
    ("Community Bank and Trust", "Community Bank & Trust", "Community Bank and Trust"),
    ("Lendistry_SOUP 2025", "Lendistry", "Lendistry 2025"),
    ("Millennium_SOUP 2025", "Millennium Bank", "Millennium 2025"),
]
LENDERS_2024 = [
    ("Lendistry_SOUP 2024", "Lendistry", "Lendistry 2024"),
    ("Millennium_SOUP 2024", "Millennium Bank", "Millennium 2024"),
    ("Mission Valley Bank", "Mission Valley Bank", "Mission Valley Bank"),
    ("Peoples Bank", "Peoples Bank", "Peoples Bank"),
    ("Poppy Bank", "Poppy Bank", "Poppy Bank"),
]

# THE VALIDATED SOUP 2025 SCOPE. Measured against the client's own supplied
# workbook and recorded in JUNE_2025_ACCEPTANCE_BASELINE.md.
ACCEPTED_LOAN_COUNTS = {
    "Community Bank & Trust": 155,
    "Lendistry": 1234,
    "Millennium Bank": 47,
}
ACCEPTED_TOTAL_LOANS = 1436
ACCEPTED_TOTAL_PAYMENT_RECEIVED = 2805017.89

# The D4 acceptance denominator. Three recommendations, one review-flagged.
# CB&T and Millennium are the two the June acceptance baseline recorded;
# Lendistry 6251789108 is the third, surfaced by PR #15's net-negative reversal
# routing (previously dropped before Human Review) and the only flagged one.
ACCEPTED_RECOMMENDATIONS = {
    ("Community Bank and Trust", "4271437005"): False,
    ("Lendistry 2025", "6251789108"): True,
    ("Millennium 2025", "7859248901"): False,
}
ACCEPTED_RECOMMENDATION_COUNT = 3
ACCEPTED_FLAGGED_COUNT = 1

# The raw population. Recorded so the difference is visible and testable, NOT
# because it is an acceptance target.
RAW_CBT_LOAN_COUNT = 156
RAW_TOTAL_PAYMENT_RECEIVED = 2805017.90
# AMENDED 2026-08-15: 95 -> 97 and 34 -> 36. VERIFIED as exactly the 2 new
# required balance candidates raised on the raw SOUP 2025 population -- Lendistry
# 2025 GP 5886179101 and GP 5991379104, both lender-reported LIQUIDATION loans
# whose reversal to ACTIVE is offered and whose cycle opened with a positive
# balance. Both are review-flagged, which is why RAW_FLAGGED_COUNT moves by the
# same 2. Nothing else in this population changed:
#   97 - 2 balance candidates = 95, the prior baseline, intact.
# Re-derived on every run by test_the_increase_is_exactly_the_balance_candidates.
RAW_RECOMMENDATION_COUNT = 97
RAW_FLAGGED_COUNT = 36
RAW_BALANCE_CANDIDATES = 2
PRE_BALANCE_RAW_RECOMMENDATION_COUNT = 95
PRE_BALANCE_RAW_FLAGGED_COUNT = 34

# Every lender token that must NOT appear in deal-routing logic.
LENDER_NAME_TOKENS = (
    "Community Bank", "CB&T", "Lendistry", "Millennium",
    "Mission Valley", "Peoples", "Poppy",
)


def _executable_only(source):
    """Source with comments and docstrings removed.

    The scope decision is documented IN comments that necessarily name the
    three lenders, so a raw substring scan would flag the very documentation
    this file exists to require. The rule is about routing LOGIC, not prose."""
    out, in_doc, delim = [], False, None
    for line in source.splitlines():
        stripped = line.strip()
        if in_doc:
            if delim in stripped:
                in_doc = False
            continue
        if stripped.startswith("#"):
            continue
        for d in ('"""', "'''"):
            if stripped.startswith(d) or (" = " + d) in stripped or ("(" + d) in stripped:
                if stripped.count(d) == 1:
                    in_doc, delim = True, d
                break
        if in_doc:
            continue
        out.append(line.split("#")[0] if "#" in line and '"' not in line else line)
    return "\n".join(out)


def _find(directory, substring):
    hits = [p for p in sorted(glob.glob(os.path.join(directory, "*.xlsx")))
            if substring.lower() in os.path.basename(p).lower()]
    return hits[0] if hits else None


_CACHE = {}


def _load(directory, lenders, key):
    """Reconcile each lender and build its MSR rows. Cached per directory."""
    if key in _CACHE:
        return _CACHE[key]
    soup = file_readers.read_soup_file(open(SOUP_PATH, "rb").read())
    memo = file_readers.resolve_memo_attribution(
        file_readers.read_service_memo(open(MEMO_PATH, "rb").read()),
        file_readers.build_memo_loan_index(soup))
    loaded = {}
    for substring, soup_key, display in lenders:
        path = _find(directory, substring)
        assert path, f"no fixture for {display} in {directory}"
        raw = open(path, "rb").read()
        sk, dl, lt, si, deal = file_readers.read_servicer_file(
            raw, filename=f"{soup_key}||{display}__{os.path.basename(path)}")
        results = checks.run_servicer_reconciliation(
            dl, lt, si, file_readers.filter_soup_by_servicer(soup, sk),
            memo_df=memo, memo_lender_key=sk, deal=deal)
        stat01 = next(r for r in results if r["id"].split("::")[0] == "STAT-01")
        excluded = [t["SBA GP Number"]
                    for t in (stat01.get("suppressed_closed_loans") or [])]
        recs, cands, _skipped = corrections.build_recommendations(stat01, display)
        issues = []
        loaded[display] = {
            "soup_key": sk, "deal": deal, "si": si, "excluded": excluded,
            "recs": recs, "candidates": cands,
            "note_series": file_readers.read_note_series_label(raw),
            "remittance_date": file_readers.read_remittance_date(raw),
            "rows": msr.build_servicer_rows(si, sk, verified_overrides={},
                                            excluded_gps=excluded, issues=issues),
            "issues": issues,
        }
    _CACHE[key] = loaded
    return loaded


def _report(loaded, displays):
    by_deal = {}
    for d in displays:
        v = loaded[d]
        by_deal.setdefault(v["deal"], []).append({
            "servicer_name": v["soup_key"], "display_label": d, "rows": v["rows"],
            "issues": v["issues"], "note_series": v["note_series"],
            "remittance_date": v["remittance_date"],
        })
    return {deal: msr.build_deal_report(deal, sv) for deal, sv in by_deal.items()}


def _accepted():
    return _load(ACCEPTED_DIR, LENDERS_2025 + LENDERS_2024, "accepted")


# ---------------------------------------------------------------------------
# The routing rule — no fixtures required, so this always runs
# ---------------------------------------------------------------------------

class DealRoutingIsDataDrivenTests(unittest.TestCase):
    """A lender name must never decide which deal a file belongs to. This is
    what makes onboarding a fourth 2025 lender a zero-code change, and what
    stops the validated scope calcifying into a hardcoded allow-list."""

    def test_deal_extraction_names_no_lender(self):
        code = _executable_only(inspect.getsource(
            file_readers.extract_deal_from_note_series))
        for token in LENDER_NAME_TOKENS:
            with self.subTest(token=token):
                self.assertNotIn(token, code)

    def test_deal_ordering_names_no_lender(self):
        code = _executable_only(inspect.getsource(checks._deal_sort_key))
        for token in LENDER_NAME_TOKENS:
            with self.subTest(token=token):
                self.assertNotIn(token, code)

    def test_the_msr_report_builder_names_no_lender(self):
        code = _executable_only(inspect.getsource(msr.build_deal_report))
        for token in LENDER_NAME_TOKENS:
            with self.subTest(token=token):
                self.assertNotIn(token, code)

    def test_the_review_endpoint_routes_without_naming_a_lender(self):
        """main.corrections_review builds msr_servicers_by_deal. Its comments
        legitimately name the three lenders to document the scope; its CODE
        must not."""
        code = _executable_only(inspect.getsource(main.corrections_review))
        for token in LENDER_NAME_TOKENS:
            with self.subTest(token=token):
                self.assertNotIn(token, code)

    def test_the_scope_is_documented_where_a_developer_will_see_it(self):
        """The counterpart to the tests above: the figures must be stated in
        prose, or the next developer reasonably reads this as a universal
        cross-deal report."""
        for module in (msr, main):
            with self.subTest(module=module.__name__):
                source = inspect.getsource(module)
                self.assertIn("SOUP 2025", source)
                self.assertIn("2,805,017.89", source)


# ---------------------------------------------------------------------------
# The validated SOUP 2025 population
# ---------------------------------------------------------------------------

@unittest.skipUnless(ACCEPTED, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _ACCEPTED_SRF_DIR")
class ValidatedSoup2025ScopeTests(unittest.TestCase):

    def test_exactly_three_lenders_resolve_to_soup_2025(self):
        loaded = _accepted()
        in_2025 = sorted(d for d, v in loaded.items() if v["deal"] == DEAL_2025)
        self.assertEqual(in_2025, ["Community Bank and Trust", "Lendistry 2025",
                                   "Millennium 2025"])

    def test_membership_comes_from_the_note_series_cell(self):
        loaded = _accepted()
        for display, v in loaded.items():
            with self.subTest(servicer=display):
                expected = DEAL_2025 if "2025" in str(v["note_series"]) else DEAL_2024
                self.assertEqual(v["deal"], expected,
                                 f"{display} Note Series {v['note_series']!r}")

    def test_per_lender_loan_counts(self):
        report = _report(_accepted(), [d for d, _, _ in
                                       [(x[2], x[1], x[0]) for x in LENDERS_2025]])
        counts = {}
        for row in report[DEAL_2025]["rows"]:
            counts[row["Lender/Servicer Name"]] = counts.get(row["Lender/Servicer Name"], 0) + 1
        self.assertEqual(counts, ACCEPTED_LOAN_COUNTS)

    def test_the_deal_totals_1436_unique_loans(self):
        report = _report(_accepted(), [x[2] for x in LENDERS_2025])
        rows = report[DEAL_2025]["rows"]
        gps = [str(r["SBA GP Number"]).strip() for r in rows]
        self.assertEqual(len(rows), ACCEPTED_TOTAL_LOANS)
        self.assertEqual(len(set(gps)), ACCEPTED_TOTAL_LOANS,
                         "every loan must appear exactly once")

    def test_total_payment_received_reconciles(self):
        report = _report(_accepted(), [x[2] for x in LENDERS_2025])
        total = sum(r["Total Payment Received"] or 0
                    for r in report[DEAL_2025]["rows"])
        self.assertAlmostEqual(total, ACCEPTED_TOTAL_PAYMENT_RECEIVED, places=2)

    def test_the_totals_reconcile_to_the_three_lender_inputs(self):
        """Not just 'the total is 1,436' but 'it is the SUM OF THE INPUTS' --
        the assertion that catches a lender silently contributing twice or not
        at all while the grand total still happens to look right."""
        loaded = _accepted()
        report = _report(loaded, [x[2] for x in LENDERS_2025])
        expected = 0
        for _sub, _sk, display in LENDERS_2025:
            v = loaded[display]
            consolidated = checks.consolidate_cycle_postings(v["si"])
            expected += len(consolidated) - len(v["excluded"])
        self.assertEqual(expected, len(report[DEAL_2025]["rows"]))
        self.assertEqual(expected, ACCEPTED_TOTAL_LOANS)

    def test_no_lender_outside_soup_2025_appears(self):
        report = _report(_accepted(), [x[2] for x in LENDERS_2025])
        permitted = set(ACCEPTED_LOAN_COUNTS)
        present = {r["Lender/Servicer Name"] for r in report[DEAL_2025]["rows"]}
        self.assertEqual(present, permitted)
        self.assertEqual(set(report[DEAL_2025]["remittance_servicers"]), permitted)


# ---------------------------------------------------------------------------
# SOUP 2024 capability is retained, and the deals do not contaminate each other
# ---------------------------------------------------------------------------

@unittest.skipUnless(ACCEPTED, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _ACCEPTED_SRF_DIR")
class DealPartitioningIsRetainedTests(unittest.TestCase):
    """SOUP 2025 being the validated scope must never be implemented as
    'SOUP 2025 is the only deal'. CLIENT-CONFIRMED (2026-08-11): one workbook
    per SOUP deal."""

    def test_soup_2024_still_produces_its_own_report(self):
        report = _report(_accepted(), [x[2] for x in LENDERS_2024])
        self.assertIn(DEAL_2024, report)
        self.assertGreater(len(report[DEAL_2024]["rows"]), 0)

    def test_uploading_both_deals_produces_two_separate_workbooks(self):
        loaded = _accepted()
        report = _report(loaded, list(loaded))
        self.assertEqual(sorted(report), [DEAL_2024, DEAL_2025])

    def test_the_2025_report_is_identical_whether_or_not_2024_is_uploaded(self):
        """No cross-deal leakage: the validated figures must not depend on what
        else the operator happened to upload."""
        loaded = _accepted()
        alone = _report(loaded, [x[2] for x in LENDERS_2025])[DEAL_2025]
        together = _report(loaded, list(loaded))[DEAL_2025]
        self.assertEqual(len(alone["rows"]), ACCEPTED_TOTAL_LOANS)
        self.assertEqual(len(together["rows"]), len(alone["rows"]))
        self.assertEqual([r["SBA GP Number"] for r in together["rows"]],
                         [r["SBA GP Number"] for r in alone["rows"]])

    def test_no_2024_loan_reaches_the_2025_report(self):
        loaded = _accepted()
        report = _report(loaded, list(loaded))
        gps_2024 = {str(r["SBA GP Number"]).strip() for r in report[DEAL_2024]["rows"]}
        gps_2025 = {str(r["SBA GP Number"]).strip() for r in report[DEAL_2025]["rows"]}
        self.assertEqual(gps_2024 & gps_2025, set())


# ---------------------------------------------------------------------------
# The D4 acceptance denominator
# ---------------------------------------------------------------------------

@unittest.skipUnless(ACCEPTED, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _ACCEPTED_SRF_DIR")
class AcceptedRecommendationBaselineTests(unittest.TestCase):
    """The denominator a D4 completeness gate must satisfy. Named by loan, so a
    change in WHICH recommendation appears fails even if the count holds."""

    def _recs_2025(self):
        loaded = _accepted()
        return [(d, r) for _s, _k, d in LENDERS_2025 for r in loaded[d]["recs"]]

    def test_the_denominator_is_three(self):
        self.assertEqual(len(self._recs_2025()), ACCEPTED_RECOMMENDATION_COUNT)

    def test_exactly_one_is_review_flagged(self):
        flagged = [r for _d, r in self._recs_2025() if r.get("reviewFlag")]
        self.assertEqual(len(flagged), ACCEPTED_FLAGGED_COUNT)
        self.assertEqual(flagged[0]["gp"], "6251789108")

    def test_each_recommendation_is_the_expected_loan(self):
        actual = {(d, r["gp"]): bool(r.get("reviewFlag")) for d, r in self._recs_2025()}
        self.assertEqual(actual, ACCEPTED_RECOMMENDATIONS)

    def test_candidates_and_recommendations_stay_parallel(self):
        """The D4 denominator is the CANDIDATE set; the review page shows
        recommendations. If these ever diverge the gate becomes unsatisfiable."""
        loaded = _accepted()
        for _s, _k, display in LENDERS_2025:
            with self.subTest(servicer=display):
                self.assertEqual(len(loaded[display]["recs"]),
                                 len(loaded[display]["candidates"]))


# ---------------------------------------------------------------------------
# The raw directory is NOT the acceptance fixture
# ---------------------------------------------------------------------------

@unittest.skipUnless(RAW, "set FALCONBRIDGE_SOUP_XLSX / _MEMO_XLSX / _SRF_DIR")
class RawSubmissionsAreAStressPopulationTests(unittest.TestCase):
    """Documented as a test rather than a comment so the difference is measured
    on every run. These figures are NOT targets -- they are the evidence that
    the raw directory must not be substituted for the accepted one."""

    def _raw(self):
        return _load(RAW_DIR, LENDERS_2025, "raw")

    def test_the_raw_population_does_not_reconcile_to_the_accepted_scope(self):
        report = _report(self._raw(), [x[2] for x in LENDERS_2025])
        total = sum(r["Total Payment Received"] or 0
                    for r in report[DEAL_2025]["rows"])
        self.assertAlmostEqual(total, RAW_TOTAL_PAYMENT_RECEIVED, places=2)
        self.assertNotAlmostEqual(total, ACCEPTED_TOTAL_PAYMENT_RECEIVED, places=2)

    def test_the_raw_cbt_file_carries_one_more_loan(self):
        loaded = self._raw()
        consolidated = checks.consolidate_cycle_postings(
            loaded["Community Bank and Trust"]["si"])
        self.assertEqual(len(consolidated), RAW_CBT_LOAN_COUNT)
        self.assertNotEqual(len(consolidated),
                            ACCEPTED_LOAN_COUNTS["Community Bank & Trust"])

    def test_the_raw_population_is_the_stress_denominator(self):
        loaded = self._raw()
        recs = [r for _s, _k, d in LENDERS_2025 for r in loaded[d]["recs"]]
        flagged = [r for r in recs if r.get("reviewFlag")]
        self.assertEqual(len(recs), RAW_RECOMMENDATION_COUNT)
        self.assertEqual(len(flagged), RAW_FLAGGED_COUNT)
        self.assertNotEqual(len(recs), ACCEPTED_RECOMMENDATION_COUNT)

    def test_the_increase_is_exactly_the_balance_candidates(self):
        """The justification for moving this recorded baseline, re-derived on every
        run so it cannot drift silently."""
        import corrections
        loaded = self._raw()
        recs = [r for _s, _k, d in LENDERS_2025 for r in loaded[d]["recs"]]
        balance = [r for r in recs
                   if r["field"] == corrections.BALANCE_DECISION_FIELD]
        self.assertEqual(len(balance), RAW_BALANCE_CANDIDATES)
        self.assertEqual(len(recs) - len(balance),
                         PRE_BALANCE_RAW_RECOMMENDATION_COUNT)
        flagged = [r for r in recs if r.get("reviewFlag")]
        self.assertEqual(len(flagged) - len(balance),
                         PRE_BALANCE_RAW_FLAGGED_COUNT)
        for rec in balance:
            self.assertTrue(rec["requiresBalanceDecision"])
            self.assertGreater(rec["balanceRestoreValue"], 0)


if __name__ == "__main__":
    unittest.main()
