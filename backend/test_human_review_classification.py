"""Human Review organization — classification, bulk approval and loan grouping.

WHAT THIS COVERS
----------------
CLIENT-CONFIRMED (2026-08-13). Human Review recommendations are organised into
review categories, and a bulk control approves ONLY the deterministic group.

    STANDARD CORRECTIONS      Delinquency Status, no review flag, no reversal
                              condition of any kind. The existing business rule
                              already determines the value.
    REVIEW REQUIRED           Loan Status (always) / Reversal-Transaction /
                              Data Domain-Exception / Terminal Status Conflict /
                              Other-Flagged. Individually judged, never bulked.

THIS IS PRESENTATION AND WORKFLOW METADATA, NOT A SECOND AUDIT ENGINE. The whole
point of these tests is that classification cannot reach anything that matters:
it cannot change a recommended value, cannot widen eligibility, cannot unlock a
write, and cannot satisfy D4. Every assertion below exists to pin one of those.

MEASURED JUNE POPULATION (both June submitted-file generations, identical):
    127 total = 77 Standard + 50 Review Required
    Review Required: 39 Reversal / 5 Loan Status / 4 Data Domain / 2 Terminal

THE THIRD REVERSAL DISCRIMINATOR. `_reversal` upstream tests only a NEGATIVE or
EXACTLY-ZERO consolidated net, so a reversed posting more than offset by positive
ones carries no review flag at all. Five real June Lendistry loans are in exactly
that state. They are classified Reversal / Transaction on the raw
has_negative_posting fact -- which changes no eligibility, no value and no review
flag, and removes them from BULK approval only.

NOT COVERED HERE. D4's own completeness/readiness arithmetic
(test_review_completeness.py), D3 per-deal isolation (test_msr_deal_isolation.py)
and every business rule (their own suites) are untouched by this work and keep
their existing coverage.
"""

import json
import os
import re
import unittest

import pandas as pd

import checks
import corrections
import test_review_completeness
from test_review_completeness import PER_SERVICER, ROSTER as ROSTER_SIZE


HEADER_ROW = 9

# The five real June loans with a negative posting inside a POSITIVE net.
NET_POSITIVE_REVERSAL_LOANS = {
    ("Lendistry 2024", "4704259107"),
    ("Lendistry 2024", "4490649110"),
    ("Lendistry 2024", "5085549105"),
    ("Lendistry 2025", "6653309107"),
    ("Lendistry 2025", "7085379101"),
}

JSX = os.path.join(os.path.dirname(__file__), "..", "frontend", "src",
                   "components", "CorrectionReview.jsx")

# Real-fixture locations, via the suite's EXISTING mechanism (the same
# FALCONBRIDGE_* variables test_soup2025_scope.py, test_monthly_servicer_report.py
# and test_br002_frontend_backend_parity.py already use). No machine-specific path
# is hard-coded here: the fixtures live outside the repository, move between
# machines, and have moved on this one, so a literal path turns a relocated file
# into a silent skip.
#
#   FALCONBRIDGE_SRF_DIR                    June SUBMITTED (pre-correction) SRFs
#   FALCONBRIDGE_SOUP_XLSX                  SOUP whose latest period is MAY 2026
#   FALCONBRIDGE_JULY_SOUP_XLSX             SOUP whose latest period is JUNE 2026
#   FALCONBRIDGE_MEMO_XLSX                  Servicer Memo Summary
#   FALCONBRIDGE_JULY_LENDISTRY_2025_XLSX   July 2026 Lendistry 2025 SRF
#
# NOTE the June set for classification is the SUBMITTED one, not the accepted vF
# one. The 127/77/50 population is what a reviewer is shown BEFORE corrections are
# applied; the accepted files have already had them applied and produce almost no
# recommendations. The accepted set is what the LOCKED REPORTING BASELINE is
# measured on, which is a different question and a different suite.
#
# THE TWO SOUP GENERATIONS ARE NOT INTERCHANGEABLE COPIES -- they are different
# REPORTING PERIODS, and pairing a lender file with the wrong one silently
# produces a plausible but wrong answer rather than an error:
#
#   a JUNE audit's prior finalized period is MAY   -> FALCONBRIDGE_SOUP_XLSX
#   a JULY audit's prior finalized period is JUNE  -> FALCONBRIDGE_JULY_SOUP_XLSX
#
# CLIENT-CONFIRMED (2026-08-14). Measured cost of getting this wrong, on the real
# July Lendistry 2025 file paired with the MAY-latest SOUP: 89 recommendations
# instead of 91, of which 52 carried a delinquency value one rung understated
# (60->90, 90->120, 120->150, 150->180, DELINQUENT->30 DAYS), 2 were false
# positives against loans that had actually cured in June, and one real Loan
# Status contradiction (GP 6666709108, a current-period Bankruptcy Date beside a
# Liquidation Date on a live balance) never surfaced at all. The application also
# dated every output June_2026. ReportingPeriodPairing below is what makes that
# mistake fail loudly instead of quietly.
SRF_DIR = os.environ.get("FALCONBRIDGE_SRF_DIR")
SOUP_XLSX = os.environ.get("FALCONBRIDGE_SOUP_XLSX")
JULY_SOUP_XLSX = os.environ.get("FALCONBRIDGE_JULY_SOUP_XLSX")
MEMO_XLSX = os.environ.get("FALCONBRIDGE_MEMO_XLSX")
JULY_LENDISTRY_2025_XLSX = os.environ.get("FALCONBRIDGE_JULY_LENDISTRY_2025_XLSX")


def _require(*paths):
    """Skip reason naming the MISSING variable/path, or None when all are usable.
    A skip must say which fixture to set, never just 'not present'."""
    missing = [p for p in paths if not p or not os.path.exists(p)]
    return None if not missing else f"fixture(s) unavailable: {missing}"


def _eligible(**over):
    """A minimal built recommendation/candidate, as build_eligible_corrections
    emits it. Standard unless overridden."""
    base = {
        "gp": "1234567890", "field": "Delinquency Status",
        "corrected_value": "30 DAYS", "reported_value": "CURRENT",
        "review_required": False, "review_reason": None,
        "reversal_equivalents": None, "offsetting_zero_net_reversal": False,
        "has_negative_posting": False,
    }
    base.update(over)
    return base


# ---------------------------------------------------------------------------
class ClassifierUnit(unittest.TestCase):
    """The classifier in isolation — every branch and every precedence pair."""

    def test_a_clean_delinquency_recommendation_is_standard(self):
        cat, std = corrections.classify_recommendation(_eligible())
        self.assertEqual(cat, corrections.STANDARD_CATEGORY)
        self.assertTrue(std)

    def test_every_loan_status_is_review_required_even_unflagged(self):
        # The confirmed Full P&I deferral branch emits with review_required=False.
        # _apply_one_correction still refuses to write it without reviewer_approved,
        # so the UI must not offer it for bulk approval.
        cat, std = corrections.classify_recommendation(
            _eligible(field="Loan Status", corrected_value="DEFERRED"))
        self.assertEqual(cat, corrections.REVIEW_CATEGORY_LOAN_STATUS)
        self.assertFalse(std)

    def test_net_negative_reversal_is_review_required(self):
        cat, std = corrections.classify_recommendation(_eligible(
            reversal_equivalents=2, review_required=True,
            review_reason=corrections.REVIEW_REASON_REVERSAL))
        self.assertEqual(cat, corrections.REVIEW_CATEGORY_REVERSAL)
        self.assertFalse(std)

    def test_offsetting_zero_net_reversal_is_review_required(self):
        cat, std = corrections.classify_recommendation(_eligible(
            offsetting_zero_net_reversal=True, review_required=True,
            review_reason=corrections.REVIEW_REASON_REVERSAL_OFFSET))
        self.assertEqual(cat, corrections.REVIEW_CATEGORY_REVERSAL)
        self.assertFalse(std)

    def test_negative_posting_inside_a_positive_net_is_review_required(self):
        """The third discriminator. NO review flag, NO reversal flag -- this would
        be Standard on flags alone, which is the hole being closed."""
        e = _eligible(has_negative_posting=True)
        self.assertFalse(e["review_required"])
        self.assertIsNone(e["review_reason"])
        cat, std = corrections.classify_recommendation(e)
        self.assertEqual(cat, corrections.REVIEW_CATEGORY_REVERSAL)
        self.assertFalse(std)

    def test_cross_domain_is_data_domain_exception(self):
        cat, std = corrections.classify_recommendation(_eligible(
            review_required=True, review_reason=corrections.REVIEW_REASON_CROSS_DOMAIN))
        self.assertEqual(cat, corrections.REVIEW_CATEGORY_DATA_DOMAIN)
        self.assertFalse(std)

    def test_terminal_status_is_its_own_category(self):
        cat, std = corrections.classify_recommendation(_eligible(
            review_required=True, review_reason=corrections.REVIEW_REASON_TERMINAL_STATUS))
        self.assertEqual(cat, corrections.REVIEW_CATEGORY_TERMINAL)
        self.assertFalse(std)

    def test_an_unrecognised_future_review_reason_fails_closed(self):
        """A reason added upstream later must be Review Required by DEFAULT, never
        silently swept into the bulk group."""
        cat, std = corrections.classify_recommendation(_eligible(
            review_required=True, review_reason="Some rule invented next quarter."))
        self.assertEqual(cat, corrections.REVIEW_CATEGORY_OTHER)
        self.assertFalse(std)

    def test_review_required_with_no_reason_string_still_fails_closed(self):
        cat, std = corrections.classify_recommendation(_eligible(review_required=True))
        self.assertEqual(cat, corrections.REVIEW_CATEGORY_OTHER)
        self.assertFalse(std)

    def test_reversal_outranks_loan_status(self):
        """Mirrors the engine: `if _reversal: continue` suppresses every Loan
        Status branch upstream, so reversal must rank first here too."""
        cat, _ = corrections.classify_recommendation(_eligible(
            field="Loan Status", reversal_equivalents=1))
        self.assertEqual(cat, corrections.REVIEW_CATEGORY_REVERSAL)

    def test_loan_status_outranks_cross_domain_and_terminal(self):
        cat, _ = corrections.classify_recommendation(_eligible(
            field="Loan Status", review_required=True,
            review_reason=corrections.REVIEW_REASON_TERMINAL_STATUS))
        self.assertEqual(cat, corrections.REVIEW_CATEGORY_LOAN_STATUS)

    def test_cross_domain_outranks_terminal_when_both_present(self):
        both = f"{corrections.REVIEW_REASON_CROSS_DOMAIN} {corrections.REVIEW_REASON_TERMINAL_STATUS}"
        cat, _ = corrections.classify_recommendation(_eligible(
            review_required=True, review_reason=both))
        self.assertEqual(cat, corrections.REVIEW_CATEGORY_DATA_DOMAIN)

    def test_no_precedence_order_can_produce_standard_for_a_flagged_item(self):
        """Standard is the terminal fallback. Exhaustively: any single flag, in any
        combination, must yield is_standard False."""
        flags = [
            {"reversal_equivalents": 1},
            {"offsetting_zero_net_reversal": True},
            {"has_negative_posting": True},
            {"field": "Loan Status"},
            {"review_required": True, "review_reason": corrections.REVIEW_REASON_CROSS_DOMAIN},
            {"review_required": True, "review_reason": corrections.REVIEW_REASON_TERMINAL_STATUS},
            {"review_required": True, "review_reason": "unknown"},
        ]
        for i in range(1, 1 << len(flags)):
            over = {}
            for bit, f in enumerate(flags):
                if i & (1 << bit):
                    over.update(f)
            _, std = corrections.classify_recommendation(_eligible(**over))
            self.assertFalse(std, f"combination {over} classified Standard")

    def test_the_classifier_never_mutates_its_input(self):
        e = _eligible(has_negative_posting=True)
        snapshot = json.dumps(e, sort_keys=True, default=str)
        corrections.classify_recommendation(e)
        corrections.classify_recommendation(e)
        self.assertEqual(json.dumps(e, sort_keys=True, default=str), snapshot)

    def test_the_classifier_is_deterministic(self):
        e = _eligible(review_required=True,
                      review_reason=corrections.REVIEW_REASON_REVERSAL,
                      reversal_equivalents=3)
        self.assertEqual(corrections.classify_recommendation(e),
                         corrections.classify_recommendation(e))

    def test_categories_are_exactly_the_published_set(self):
        self.assertEqual(len(corrections.REVIEW_REQUIRED_CATEGORIES), 5)
        self.assertNotIn(corrections.STANDARD_CATEGORY,
                         corrections.REVIEW_REQUIRED_CATEGORIES)
        # Loan Status is displayed first; reversal outranks it only in PRECEDENCE.
        self.assertEqual(corrections.REVIEW_REQUIRED_CATEGORIES[0],
                         corrections.REVIEW_CATEGORY_LOAN_STATUS)


# ---------------------------------------------------------------------------
class ClassificationCannotAlterTheRecommendation(unittest.TestCase):
    """The value the audit produced is untouched by classifying it. Uses
    build_recommendations end to end on a synthetic STAT-01 result."""

    @classmethod
    def setUpClass(cls):
        cls.correction_data = [
            {  # clean -> Standard
                "gp": "1000000001", "reasons": ["DELINQUENCY STATUS MISMATCH"],
                "dq_match": False, "calculated_dq": "30 DAYS", "srf_dq": "CURRENT",
                "srf_status": "ACTIVE", "soup_status": "ACTIVE",
                "excel_rows": [10], "sheet_name": "Servicer Inputs",
                "header_excel_row": HEADER_ROW, "ending_balance": 1000.0,
                "posting_count": 1, "has_negative_posting": False,
            },
            {  # negative posting, positive net -> Reversal, still unflagged
                "gp": "1000000002", "reasons": ["DELINQUENCY STATUS MISMATCH"],
                "dq_match": False, "calculated_dq": "DELINQUENT", "srf_dq": "CURRENT",
                "srf_status": "ACTIVE", "soup_status": "ACTIVE",
                "excel_rows": [11], "sheet_name": "Servicer Inputs",
                "header_excel_row": HEADER_ROW, "ending_balance": 2000.0,
                "posting_count": 3, "has_negative_posting": True,
                "payment_received": 500.0,
            },
        ]

    def _build(self):
        return corrections.build_recommendations(
            {"correction_data": self.correction_data}, "Test Bank")

    def test_calculated_values_are_exactly_stat01s(self):
        recs, _, _ = self._build()
        by_gp = {r["gp"]: r for r in recs}
        self.assertEqual(by_gp["1000000001"]["calculated"], "30 DAYS")
        self.assertEqual(by_gp["1000000002"]["calculated"], "DELINQUENT")

    def test_the_review_flag_is_unchanged_by_classification(self):
        """The net-positive reversal loan is classified Reversal but keeps its
        ORIGINAL (absent) review flag -- classification is not a flag."""
        recs, _, _ = self._build()
        r = next(x for x in recs if x["gp"] == "1000000002")
        self.assertIsNone(r["reviewFlag"])
        self.assertEqual(r["category"], corrections.REVIEW_CATEGORY_REVERSAL)
        self.assertFalse(r["standard"])

    def test_eligibility_is_unchanged_by_classification(self):
        """Both loans are still candidates. Classification narrows the BULK group,
        never the candidate set -- removing one would remove it from D4's
        denominator, which would be a business rule change."""
        recs, candidates, skipped = self._build()
        self.assertEqual(len(recs), 2)
        self.assertEqual(len(candidates), 2)
        self.assertEqual(skipped, [])

    def test_candidates_carry_no_classification_the_write_path_could_read(self):
        """The write path acts on session CANDIDATES. Classification lives on the
        client-facing recommendation only, so no write can consult it."""
        _, candidates, _ = self._build()
        for c in candidates:
            self.assertNotIn("standard", c)
            self.assertNotIn("category", c)

    def test_every_recommendation_gets_exactly_one_category(self):
        recs, _, _ = self._build()
        for r in recs:
            self.assertIn(r["category"],
                          (corrections.STANDARD_CATEGORY,) + corrections.REVIEW_REQUIRED_CATEGORIES)
            self.assertIs(r["standard"], r["category"] == corrections.STANDARD_CATEGORY)


# ---------------------------------------------------------------------------
class EvidencePassthroughIsDisplayOnly(unittest.TestCase):
    """The new evidence fields reach the review page and change nothing en route."""

    def _si(self, **over):
        row = {
            "SBA GP Number": "2000000001", "Lender Loan Number": "L-1",
            "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
            "Payment Received": 500.0, "Principal": 100.0, "Net Interest": 400.0,
            "Servicing Fee": 10.0, "Scheduled Payment": 500.0,
            "Beginning Balance": 100000.0, "Ending Balance": 99900.0,
            "Loan Payment Instance": 1, "Last Payment Instance?": "Y",
            "_excel_row": HEADER_ROW + 1, "_sheet_name": "Servicer Inputs",
            "_header_excel_row": HEADER_ROW,
        }
        row.update(over)
        return pd.DataFrame([row])

    def test_the_new_columns_are_consolidated_when_present(self):
        out = checks.consolidate_cycle_postings(self._si(**{
            "Foreclosure Date": "01/15/2026", "REO Date": "",
            "Liquidation Proceeds": 1234.56, "Liquidation Loss": 78.90,
            "Subsequent Recoveries/(Losses)": -5.00,
        }))
        r = out.loc["2000000001"]
        self.assertEqual(r["foreclosure_date"], "01/15/2026")
        self.assertTrue(r["foreclosure_date_populated"])
        self.assertEqual(r["reo_date"], "")
        self.assertFalse(r["reo_date_populated"])
        self.assertEqual(r["liquidation_proceeds"], 1234.56)
        self.assertEqual(r["liquidation_loss"], 78.90)
        self.assertEqual(r["subsequent_recoveries"], -5.00)

    def test_absent_columns_yield_none_not_zero(self):
        """A file without these columns must not report $0.00 proceeds -- that
        would read as 'measured and zero' rather than 'not reported'."""
        out = checks.consolidate_cycle_postings(self._si())
        r = out.loc["2000000001"]
        self.assertIsNone(r["liquidation_proceeds"])
        self.assertIsNone(r["liquidation_loss"])
        self.assertIsNone(r["subsequent_recoveries"])
        self.assertEqual(r["foreclosure_date"], "")

    def test_amounts_sum_across_postings_like_every_other_money_column(self):
        two = pd.concat([
            self._si(**{"Liquidation Proceeds": 100.0, "_excel_row": 10,
                        "Loan Payment Instance": 1}),
            self._si(**{"Liquidation Proceeds": 250.0, "_excel_row": 11,
                        "Loan Payment Instance": 2}),
        ], ignore_index=True)
        out = checks.consolidate_cycle_postings(two)
        self.assertEqual(out.loc["2000000001"]["liquidation_proceeds"], 350.0)

    def test_the_new_facts_change_no_existing_consolidated_value(self):
        """Byte-for-byte: every pre-existing consolidated value is identical with
        and without the new columns present in the source."""
        keys = ["posting_count", "loan_status", "dq_status", "received",
                "has_negative_posting", "has_filed_posting", "servicing_fee",
                "principal", "net_interest", "scheduled", "scheduled_ambiguous",
                "beginning_balance", "ending_balance", "bankruptcy_date",
                "liquidation_date", "paid_in_full_date"]
        without = checks.consolidate_cycle_postings(self._si())
        with_new = checks.consolidate_cycle_postings(self._si(**{
            "Foreclosure Date": "01/15/2026", "REO Date": "02/20/2026",
            "Liquidation Proceeds": 9999.99, "Liquidation Loss": 8888.88,
            "Subsequent Recoveries/(Losses)": 777.77,
        }))
        for k in keys:
            self.assertEqual(without.loc["2000000001"][k],
                             with_new.loc["2000000001"][k],
                             f"{k} changed when review-evidence columns were added")

    def test_the_evidence_maps_are_separate_from_the_eligibility_gate_map(self):
        """_TERMINAL_EVENT_DATE_COLUMNS drives Loan Status ELIGIBILITY
        (bankruptcy_date_populated is a write gate). The review-evidence map must
        not be merged into it, or the eligibility layer silently gains new keys."""
        self.assertNotIn("Foreclosure Date", checks._TERMINAL_EVENT_DATE_COLUMNS)
        self.assertNotIn("REO Date", checks._TERMINAL_EVENT_DATE_COLUMNS)
        self.assertEqual(set(checks._TERMINAL_EVENT_DATE_COLUMNS),
                         {"Bankruptcy Date", "Paid-In-Full Date", "Liquidation Date"})
        self.assertEqual(set(checks._REVIEW_EVIDENCE_DATE_COLUMNS),
                         {"Foreclosure Date", "REO Date"})

    def test_the_override_domain_is_the_five_permitted_fields(self):
        """AMENDED 2026-08-15. Loan Status and Delinquency Status were the whole
        domain until the reviewer-approved liquidation-reversal balance decision
        added Ending Balance, Liquidation Date and Liquidation Loss.

        The principle this test defends is unchanged: DISPLAY EVIDENCE MUST NEVER
        BECOME AN OVERRIDE TARGET. The three additions are not display evidence --
        they are the fields a reviewer explicitly decides on, gated in
        corrections.py. The genuinely display-only facts are asserted absent
        below, which is the half that actually guards the boundary."""
        self.assertEqual(set(checks._VERIFIED_OVERRIDE_COLUMNS),
                         {"Loan Status", "Delinquency Status", "Ending Balance",
                          "Liquidation Date", "Liquidation Loss"})

    def test_display_only_evidence_is_still_never_an_override_target(self):
        """The unchanged half. Foreclosure Date, REO Date, Liquidation Proceeds,
        Subsequent Recoveries and the prior-period balance are shown to the
        reviewer and can never be written."""
        for field in ("Foreclosure Date", "REO Date", "Liquidation Proceeds",
                      "Subsequent Recoveries/(Losses)", "Bankruptcy Date",
                      "Paid-In-Full Date", "Beginning Balance",
                      "Payment Received", "Scheduled Payment", "Servicing Fee"):
            with self.subTest(field=field):
                self.assertNotIn(field, checks._VERIFIED_OVERRIDE_COLUMNS)


# ---------------------------------------------------------------------------
class FrontendContract(unittest.TestCase):
    """The review page's Loan Status domain, its bulk control and the removal of
    the superseded flat table. Asserted against the JSX source, the same way the
    existing SERVICER_SOUP_KEYS and DQ_LADDER parity tests are."""

    @classmethod
    def setUpClass(cls):
        with open(JSX, encoding="utf-8") as fh:
            cls.src = fh.read()

    def _array(self, name):
        m = re.search(rf"const {name} = \[(.*?)\]", self.src, re.S)
        self.assertIsNotNone(m, f"{name} not found in CorrectionReview.jsx")
        return [v for v in re.findall(r"'([^']+)'", m.group(1))]

    def test_the_loan_status_dropdown_is_exactly_the_permitted_domain(self):
        self.assertEqual(set(self._array("LOAN_STATUS_DOMAIN")),
                         corrections._PERMITTED_LOAN_STATUS)
        self.assertEqual(len(self._array("LOAN_STATUS_DOMAIN")), 9)

    def test_the_dq_dropdown_stays_in_the_permitted_dq_domain(self):
        self.assertTrue(set(self._array("DQ_LADDER")) <= corrections._PERMITTED_DQ_STATUS)

    def test_loan_status_is_chosen_from_a_select_never_typed(self):
        """No free text: the approved-value control is a <select> over
        domainFor(field), and no text input is bound to approved_value."""
        self.assertIn("domainFor(rec.field).map", self.src)
        self.assertNotRegex(
            self.src,
            r"type=\"text\"[^>]*approved_value|approved_value[^)]*type=\"text\"")

    def test_the_jsx_category_constants_match_the_backend(self):
        for name, expected in (
            ("STANDARD_CATEGORY", corrections.STANDARD_CATEGORY),
            ("REVIEW_CATEGORY_LOAN_STATUS", corrections.REVIEW_CATEGORY_LOAN_STATUS),
            ("REVIEW_CATEGORY_REVERSAL", corrections.REVIEW_CATEGORY_REVERSAL),
            ("REVIEW_CATEGORY_DATA_DOMAIN", corrections.REVIEW_CATEGORY_DATA_DOMAIN),
            ("REVIEW_CATEGORY_TERMINAL", corrections.REVIEW_CATEGORY_TERMINAL),
            ("REVIEW_CATEGORY_OTHER", corrections.REVIEW_CATEGORY_OTHER),
        ):
            m = re.search(rf"const {name} = '([^']+)'", self.src)
            self.assertIsNotNone(m, f"{name} missing from CorrectionReview.jsx")
            self.assertEqual(m.group(1), expected)

    def test_the_review_required_category_order_matches_the_backend(self):
        m = re.search(r"const REVIEW_REQUIRED_CATEGORIES = \[(.*?)\]", self.src, re.S)
        names = re.findall(r"(REVIEW_CATEGORY_\w+)", m.group(1))
        resolved = [re.search(rf"const {n} = '([^']+)'", self.src).group(1) for n in names]
        self.assertEqual(tuple(resolved), corrections.REVIEW_REQUIRED_CATEGORIES)

    def test_bulk_approval_is_sourced_from_the_server_standard_flag(self):
        """standardRecs is filtered on r.standard -- the SERVER's classification.
        The page must not re-derive which items are deterministic."""
        self.assertIn("recs.filter((r) => r.standard)", self.src)
        self.assertIn(
            "standardRecs.filter((r) => !decisionOf(r.id).decision && !loanStatusUndecided(r))",
            self.src)

    def test_bulk_approval_never_submits_and_never_calls_an_endpoint(self):
        m = re.search(r"const approveAllStandard = \(\) =>(.*?)\n\n", self.src, re.S)
        self.assertIsNotNone(m)
        body = m.group(1)
        for forbidden in ("axios", "submit(", "fetch(", "await", "/api/"):
            self.assertNotIn(forbidden, body,
                             f"approveAllStandard must not contain {forbidden!r}")

    def test_bulk_approval_never_overwrites_an_existing_decision(self):
        m = re.search(r"const approveAllStandard = \(\) =>(.*?)\n\n", self.src, re.S)
        self.assertIn("if (!(out[r.id] && out[r.id].decision))", m.group(1))

    def test_bulk_approval_is_not_wired_to_any_mount_effect(self):
        """It must never fire on page load. There is no useEffect in this file at
        all, and approveAllStandard is referenced exactly once as an onClick."""
        self.assertNotIn("useEffect", self.src)
        self.assertEqual(self.src.count("onClick={approveAllStandard}"), 1)
        self.assertEqual(len(re.findall(r"\bapproveAllStandard\b", self.src)), 2)

    def test_the_companion_hold_keys_on_an_undecided_loan_status(self):
        m = re.search(r"const loanStatusUndecided = \(rec\) =>(.*?)\n\n", self.src, re.S)
        self.assertIsNotNone(m)
        body = m.group(1)
        self.assertIn("loanKeyOf(other) === loanKeyOf(rec)", body)
        self.assertIn("other.field === 'Loan Status'", body)
        self.assertIn("!decisionOf(other.id).decision", body)

    def test_the_loan_card_groups_on_servicer_and_gp_never_gp_alone(self):
        # Matched on the return line, not a brace-balanced body: the key is a
        # template literal, so `${rec.servicer}` contains a `}` of its own and a
        # non-greedy \{(.*?)\} truncates the capture mid-expression.
        m = re.search(r"function loanKeyOf\(rec\) \{\s*return (.+)", self.src)
        self.assertIsNotNone(m, "loanKeyOf not found in CorrectionReview.jsx")
        key_expr = m.group(1)
        self.assertIn("rec.servicer", key_expr)
        self.assertIn("rec.gp", key_expr)
        # Both halves in one key, separated -- never gp alone.
        self.assertEqual(key_expr.strip(), "`${rec.servicer}||${rec.gp}`")

    def test_each_grouped_row_keeps_its_own_id_and_controls(self):
        self.assertIn("id={`rec-${rec.id}`}", self.src)            # jump anchor
        self.assertIn("onChange(rec.id, { decision:", self.src)     # own decision
        self.assertIn("onChange(rec.id, { approved_value:", self.src)
        self.assertIn("key={rec.id}", self.src)

    def test_the_superseded_flat_review_table_is_gone(self):
        """Exactly one table, one thead, one tbody, and no flat `recs.map` render
        driver -- the old undifferentiated list must not survive anywhere."""
        self.assertEqual(self.src.count("<table"), 1)
        self.assertEqual(self.src.count("</table>"), 1)
        self.assertEqual(self.src.count("<thead>"), 1)
        self.assertEqual(self.src.count("<tbody>"), 1)
        self.assertEqual(self.src.count("<h3"), 1)
        self.assertEqual(len(re.findall(r"\{recs\.map\(", self.src)), 0)
        self.assertEqual(self.src.count("{group.recs.map("), 1)
        for once in ("Reject all remaining", "Jump to next undecided",
                     "Approve All Standard Corrections", "STANDARD CORRECTIONS",
                     "Current-period evidence:", "Deferral evidence:"):
            self.assertEqual(self.src.count(once), 1, f"{once!r} appears more than once")

    def test_the_loan_card_shows_the_approved_evidence_fields(self):
        for token in ("priorEndingBalance", "liquidationProceeds", "liquidationLoss",
                      "subsequentRecoveries", "foreclosureDate", "reoDate",
                      "servicerLoanStatus", "priorLoanStatus", "paymentReceived",
                      "liquidationDate", "bankruptcyDate"):
            self.assertIn(token, self.src, f"loan card evidence {token} missing")

    def test_the_contradiction_banner_states_the_conflict_without_resolving_it(self):
        self.assertIn("CONTRADICTORY", self.src)
        self.assertIn("weigh both before approving", self.src)

    def test_d4_submission_gating_is_unchanged(self):
        self.assertIn("const blocked = !reviewer.trim() || decidedCount === 0", self.src)
        self.assertIn("review_completeness", self.src)
        self.assertIn("reporting_readiness", self.src)


# ---------------------------------------------------------------------------
class BulkApprovalSemantics(unittest.TestCase):
    """The page's two selectors, re-implemented in Python from the same inputs so
    the RULES are pinned by assertion rather than only by a source-text check."""

    @staticmethod
    def _loan_key(r):
        return (r["servicer"], str(r["gp"]))

    @classmethod
    def _bulk(cls, recs, decisions):
        held, eligible = [], []
        for r in recs:
            if not r["standard"] or decisions.get(r["id"]):
                continue
            companion_open = any(
                cls._loan_key(o) == cls._loan_key(r)
                and o["field"] == "Loan Status"
                and not decisions.get(o["id"])
                for o in recs
            )
            (held if companion_open else eligible).append(r)
        return eligible, held

    @staticmethod
    def _rec(servicer, gp, field, standard, category):
        return {"id": f"{servicer}||{gp}||{field}", "servicer": servicer, "gp": gp,
                "field": field, "standard": standard, "category": category}

    def setUp(self):
        self.recs = [
            self._rec("CBT", "1", "Delinquency Status", True, corrections.STANDARD_CATEGORY),
            self._rec("CBT", "2", "Delinquency Status", True, corrections.STANDARD_CATEGORY),
            # a mixed loan: Standard DQ + Review Required Loan Status
            self._rec("CBT", "3", "Delinquency Status", True, corrections.STANDARD_CATEGORY),
            self._rec("CBT", "3", "Loan Status", False, corrections.REVIEW_CATEGORY_LOAN_STATUS),
            self._rec("CBT", "4", "Delinquency Status", False, corrections.REVIEW_CATEGORY_REVERSAL),
            self._rec("CBT", "5", "Delinquency Status", False, corrections.REVIEW_CATEGORY_DATA_DOMAIN),
            self._rec("CBT", "6", "Delinquency Status", False, corrections.REVIEW_CATEGORY_TERMINAL),
            self._rec("CBT", "7", "Delinquency Status", False, corrections.REVIEW_CATEGORY_OTHER),
            # same GP at a DIFFERENT servicer -- must not be treated as one loan
            self._rec("Poppy", "3", "Delinquency Status", True, corrections.STANDARD_CATEGORY),
        ]

    def test_only_standard_recommendations_are_ever_eligible(self):
        eligible, _ = self._bulk(self.recs, {})
        for r in eligible:
            self.assertTrue(r["standard"])
            self.assertEqual(r["category"], corrections.STANDARD_CATEGORY)

    def test_no_review_required_category_can_be_bulk_approved(self):
        eligible, held = self._bulk(self.recs, {})
        swept = {r["id"] for r in eligible} | {r["id"] for r in held}
        for r in self.recs:
            if not r["standard"]:
                self.assertNotIn(r["id"], swept, f"{r['category']} was swept")

    def test_the_standard_companion_of_an_undecided_loan_status_is_held(self):
        eligible, held = self._bulk(self.recs, {})
        self.assertEqual([r["id"] for r in held], ["CBT||3||Delinquency Status"])
        self.assertNotIn("CBT||3||Delinquency Status", [r["id"] for r in eligible])

    def test_deciding_the_loan_status_releases_the_companion(self):
        for decision in ("approve", "reject", "return"):
            eligible, held = self._bulk(self.recs, {"CBT||3||Loan Status": decision})
            self.assertEqual(held, [], f"still held after Loan Status {decision}")
            self.assertIn("CBT||3||Delinquency Status", [r["id"] for r in eligible])

    def test_the_same_gp_at_another_servicer_is_not_held(self):
        _, held = self._bulk(self.recs, {})
        self.assertNotIn("Poppy||3||Delinquency Status", [r["id"] for r in held])

    def test_an_existing_decision_is_never_re_swept(self):
        eligible, held = self._bulk(self.recs, {"CBT||1||Delinquency Status": "reject"})
        self.assertNotIn("CBT||1||Delinquency Status",
                         [r["id"] for r in eligible] + [r["id"] for r in held])

    def test_the_held_item_is_still_counted_as_a_standard_correction(self):
        """The 77/50 split is per RECOMMENDATION and is not changed by the hold."""
        self.assertEqual(sum(1 for r in self.recs if r["standard"]), 4)
        eligible, held = self._bulk(self.recs, {})
        self.assertEqual(len(eligible) + len(held), 4)


# ---------------------------------------------------------------------------
class RealJuneFixtures(unittest.TestCase):
    """The locked June counts and grouping, on the real June submitted files."""

    FILES = [
        ("Community Bank & Trust", "Community Bank and Trust",
         "2026.06_Community Bank and Trust_SOUP 2025-1 SRF Template_June 2026.xlsx"),
        ("Lendistry", "Lendistry 2024",
         "2026.06_Lendistry_SOUP 2024-1 SRF Template_June 2026.xlsx"),
        ("Lendistry", "Lendistry 2025",
         "2026.06_Lendistry_SOUP 2025-1 SRF Template_June 2026.xlsx"),
        ("Millennium Bank", "Millennium 2024",
         "2026.06_Millennium_SOUP 2024-1 SRF Template_June 2026.xlsx"),
        ("Millennium Bank", "Millennium 2025",
         "2026.06_Millennium_SOUP 2025-1 SFR Template_June 2026.xlsx"),
        ("Mission Valley Bank", "Mission Valley Bank",
         "2026.06_Mission Valley Bank_SOUP 2024-1 SRF Template_June 2026.xlsx"),
        ("Peoples Bank", "Peoples Bank",
         "2026.06_Peoples Bank_SOUP 2024-1 SRF June 2026 Submitted 7.8.26.xlsx"),
        ("Poppy Bank", "Poppy Bank",
         "2026.06_Poppy Bank_SOUP 2024-1 SRF Template_June 2026.xlsx"),
    ]

    @classmethod
    def setUpClass(cls):
        import file_readers
        reason = _require(SRF_DIR, SOUP_XLSX, MEMO_XLSX,
                          *[os.path.join(SRF_DIR or "", f) for _, _, f in cls.FILES])
        if reason:
            raise unittest.SkipTest(
                f"June submitted fixtures unavailable -- set FALCONBRIDGE_SRF_DIR, "
                f"FALCONBRIDGE_SOUP_XLSX, FALCONBRIDGE_MEMO_XLSX ({reason})")
        soup = file_readers.read_soup_file(open(SOUP_XLSX, "rb").read())
        memo = file_readers.resolve_memo_attribution(
            file_readers.read_service_memo(open(MEMO_XLSX, "rb").read()),
            file_readers.build_memo_loan_index(soup))
        cls.recs = []
        for key, label, fn in cls.FILES:
            k, l, lt, si, deal = file_readers.read_servicer_file(
                open(os.path.join(SRF_DIR, fn), "rb").read(),
                filename=f"{key}||{label}__{fn}")
            res = checks.run_servicer_reconciliation(
                l, lt, si, file_readers.filter_soup_by_servicer(soup, k),
                memo_df=memo, memo_lender_key=k, deal=deal)
            stat01 = next(r for r in res if r["id"].split("::")[0] == "STAT-01")
            recs, _, _ = corrections.build_recommendations(stat01, l)
            cls.recs.extend(recs)

    # AMENDED 2026-08-15: the locked partition moved 127 -> 129 and Loan Status
    # 5 -> 7. VERIFIED, not assumed: the increase is EXACTLY 2 new required
    # balance candidates, and every other figure is untouched.
    #
    #   total                       127 -> 129   (+2)
    #   balance candidates            0 ->   2   (+2)
    #   total minus balance         127 -> 127   ( 0)   <- the old baseline, intact
    #   Standard — Delinquency       77 ->  77   ( 0)
    #   Reversal / Transaction       39 ->  39   ( 0)
    #   Data Domain / Exception       4 ->   4   ( 0)
    #   Terminal Status Conflict      2 ->   2   ( 0)
    #   Loan Status                   5 ->   7   (+2)  <- the 2 balance candidates
    #
    # The 2 loans, each a lender-reported LIQUIDATION whose reversal to ACTIVE is
    # offered, whose cycle opened with a positive balance, AND whose reported
    # Ending Balance the lender zeroed:
    #   Lendistry 2025   GP 5886179101   restore 20,088.08   reported ending 0.00
    #   Lendistry 2025   GP 5991379104   restore 20,203.66   reported ending 0.00
    #
    # Mission Valley GP 4778059109 is a LIQUIDATION -> ACTIVE reversal too, and is
    # DELIBERATELY EXCLUDED: its reported Ending Balance is 276,702.10, not zero,
    # so neither offered choice would be correct for it (see condition 4 in
    # corrections.BALANCE_DECISION_FIELD). That exclusion is what keeps this
    # baseline at +2 rather than +3.
    #
    # test_the_increase_is_exactly_the_balance_candidates below re-derives this on
    # every run, so the baseline cannot drift without the reason being visible.
    LOCKED_TOTAL = 129
    LOCKED_STANDARD = 77
    LOCKED_REVIEW = 52
    LOCKED_BALANCE_CANDIDATES = 2
    PRE_BALANCE_TOTAL = 127
    PRE_BALANCE_LOAN_STATUS = 5

    def test_the_locked_partition_is_129_equals_77_plus_52(self):
        std = [r for r in self.recs if r["standard"]]
        rev = [r for r in self.recs if not r["standard"]]
        self.assertEqual(len(self.recs), self.LOCKED_TOTAL)
        self.assertEqual(len(std), self.LOCKED_STANDARD)
        self.assertEqual(len(rev), self.LOCKED_REVIEW)
        self.assertEqual(len(std) + len(rev), len(self.recs))

    def test_the_increase_is_exactly_the_balance_candidates(self):
        """The whole justification for moving a locked baseline, re-derived."""
        balance = [r for r in self.recs
                   if r["field"] == corrections.BALANCE_DECISION_FIELD]
        self.assertEqual(len(balance), self.LOCKED_BALANCE_CANDIDATES)
        self.assertEqual(len(self.recs) - len(balance), self.PRE_BALANCE_TOTAL)
        # and every one of them is a real liquidation reversal with a positive
        # restore value -- not an accidental emission
        for rec in balance:
            self.assertTrue(rec["requiresBalanceDecision"])
            self.assertGreater(rec["balanceRestoreValue"], 0)
            self.assertFalse(rec["standard"])

    def test_the_standard_partition_did_not_move(self):
        """The balance decision must never be sweepable by the bulk control."""
        self.assertEqual(len([r for r in self.recs if r["standard"]]),
                         self.LOCKED_STANDARD)

    def test_the_review_required_category_breakdown(self):
        from collections import Counter
        got = Counter(r["category"] for r in self.recs if not r["standard"])
        self.assertEqual(dict(got), {
            corrections.REVIEW_CATEGORY_REVERSAL: 39,
            # 5 pre-existing Loan Status items + 2 balance decisions, which are
            # deliberately classified WITH their Loan Status sibling so the
            # reviewer reads the pair as one judgement.
            corrections.REVIEW_CATEGORY_LOAN_STATUS: 7,
            corrections.REVIEW_CATEGORY_DATA_DOMAIN: 4,
            corrections.REVIEW_CATEGORY_TERMINAL: 2,
        })

    def test_the_non_loan_status_categories_are_untouched(self):
        from collections import Counter
        got = Counter(r["category"] for r in self.recs if not r["standard"])
        self.assertEqual(got[corrections.REVIEW_CATEGORY_REVERSAL], 39)
        self.assertEqual(got[corrections.REVIEW_CATEGORY_DATA_DOMAIN], 4)
        self.assertEqual(got[corrections.REVIEW_CATEGORY_TERMINAL], 2)

    def test_every_loan_status_recommendation_is_review_required(self):
        ls = [r for r in self.recs if r["field"] == "Loan Status"]
        self.assertEqual(len(ls), self.PRE_BALANCE_LOAN_STATUS)
        for r in ls:
            self.assertFalse(r["standard"])
            self.assertEqual(r["category"], corrections.REVIEW_CATEGORY_LOAN_STATUS)
        # two of them carry NO review flag and would otherwise be Standard
        self.assertEqual(sum(1 for r in ls if not r["reviewFlag"]), 2)

    def test_no_flagged_recommendation_is_standard(self):
        for r in self.recs:
            if r["reviewFlag"]:
                self.assertFalse(r["standard"], f"{r['id']} flagged yet Standard")

    def test_the_five_net_positive_reversal_loans_are_reversal_not_standard(self):
        found = set()
        for r in self.recs:
            if (r["servicer"], str(r["gp"])) in NET_POSITIVE_REVERSAL_LOANS \
                    and r["field"] == "Delinquency Status":
                found.add((r["servicer"], str(r["gp"])))
                self.assertTrue(r["hasNegativePosting"])
                self.assertIsNone(r["reviewFlag"])         # unflagged upstream
                self.assertFalse(r["standard"])            # yet never bulk-approvable
                self.assertEqual(r["category"], corrections.REVIEW_CATEGORY_REVERSAL)
        self.assertEqual(found, NET_POSITIVE_REVERSAL_LOANS)

    def test_the_two_mixed_loans_are_the_expected_cbt_pair(self):
        """The only June loans whose Standard companion is held out of bulk."""
        from collections import defaultdict
        by_loan = defaultdict(list)
        for r in self.recs:
            by_loan[(r["servicer"], str(r["gp"]))].append(r)
        mixed = {k for k, v in by_loan.items()
                 if any(x["standard"] for x in v) and any(not x["standard"] for x in v)}
        self.assertEqual(mixed, {("Community Bank and Trust", "6717929105"),
                                 ("Community Bank and Trust", "5753609106")})

    def test_multi_recommendation_loans_keep_distinct_ids(self):
        from collections import defaultdict
        by_loan = defaultdict(list)
        for r in self.recs:
            by_loan[(r["servicer"], str(r["gp"]))].append(r)
        multi = {k: v for k, v in by_loan.items() if len(v) > 1}
        self.assertEqual(len(multi), 5)
        for (servicer, gp), group in multi.items():
            self.assertEqual(len({r["id"] for r in group}), len(group))
            self.assertEqual(len({r["field"] for r in group}), len(group))
        self.assertEqual(len({r["id"] for r in self.recs}), len(self.recs))

    def test_prior_ending_balance_reached_every_recommendation(self):
        self.assertTrue(all(r["priorEndingBalance"] is not None for r in self.recs))


# ---------------------------------------------------------------------------
class JulyLiquidationContradiction(unittest.TestCase):
    """The client's representative contradictory-evidence case: Lendistry 2025,
    July 2026, GP 6649779105 -- a current-period Bankruptcy Date supporting
    BANKRUPTCY while a Liquidation Date sits against a POSITIVE Ending Balance.

    The point of this class is what the application must NOT do: it must not infer
    LIQUIDATION, must not preselect it, and must not resolve the contradiction. It
    states both facts and leaves the choice to the reviewer, who may select any of
    the nine permitted statuses."""

    GP = "6649779105"

    @classmethod
    def setUpClass(cls):
        import file_readers
        # FALCONBRIDGE_JULY_SOUP_XLSX, never FALCONBRIDGE_SOUP_XLSX: a July audit's
        # prior finalized period is JUNE. See the module header for what pairing
        # this file with the May-latest SOUP silently does to the answer.
        reason = _require(JULY_LENDISTRY_2025_XLSX, JULY_SOUP_XLSX, MEMO_XLSX)
        if reason:
            raise unittest.SkipTest(
                f"July Lendistry 2025 fixtures unavailable -- set "
                f"FALCONBRIDGE_JULY_LENDISTRY_2025_XLSX, FALCONBRIDGE_JULY_SOUP_XLSX, "
                f"FALCONBRIDGE_MEMO_XLSX ({reason})")
        soup = file_readers.read_soup_file(open(JULY_SOUP_XLSX, "rb").read())
        memo = file_readers.resolve_memo_attribution(
            file_readers.read_service_memo(open(MEMO_XLSX, "rb").read()),
            file_readers.build_memo_loan_index(soup))
        k, l, lt, si, deal = file_readers.read_servicer_file(
            open(JULY_LENDISTRY_2025_XLSX, "rb").read(),
            filename="Lendistry||Lendistry 2025__"
                     + os.path.basename(JULY_LENDISTRY_2025_XLSX))
        res = checks.run_servicer_reconciliation(
            l, lt, si, file_readers.filter_soup_by_servicer(soup, k),
            memo_df=memo, memo_lender_key=k, deal=deal)
        stat01 = next(r for r in res if r["id"].split("::")[0] == "STAT-01")
        cls.recs, cls.candidates, _ = corrections.build_recommendations(stat01, l)
        cls.loan = [r for r in cls.recs if str(r["gp"]) == cls.GP]
        cls.cycle_month = stat01.get("_cycle_month")
        cls.servicer = l

    def test_the_loan_produces_two_independent_recommendations(self):
        self.assertEqual(len(self.loan), 2)
        self.assertEqual({r["field"] for r in self.loan},
                         {"Loan Status", "Delinquency Status"})
        self.assertEqual({r["id"] for r in self.loan},
                         {f"Lendistry 2025||{self.GP}||Loan Status",
                          f"Lendistry 2025||{self.GP}||Delinquency Status"})

    def test_the_loan_status_recommendation_is_review_required(self):
        ls = next(r for r in self.loan if r["field"] == "Loan Status")
        self.assertFalse(ls["standard"])
        self.assertEqual(ls["category"], corrections.REVIEW_CATEGORY_LOAN_STATUS)

    def test_the_contradictory_evidence_is_stated_not_resolved(self):
        ls = next(r for r in self.loan if r["field"] == "Loan Status")
        self.assertIn(corrections.REVIEW_REASON_BANKRUPTCY_LIQUIDATION_CONTRADICTION,
                      ls["reviewFlag"])
        self.assertIn(corrections.REVIEW_REASON_BANKRUPTCY_CURRENT_PERIOD,
                      ls["reviewFlag"])

    def test_liquidation_is_never_the_recommended_value(self):
        """The application must not infer LIQUIDATION from a Liquidation Date."""
        ls = next(r for r in self.loan if r["field"] == "Loan Status")
        self.assertEqual(ls["calculated"], "BANKRUPTCY")
        self.assertNotEqual(ls["calculated"], "LIQUIDATION")
        self.assertNotIn("LIQUIDATION", corrections._PROPOSABLE_LOAN_STATUS)

    def test_the_reviewer_can_see_the_evidence_that_causes_the_contradiction(self):
        ls = next(r for r in self.loan if r["field"] == "Loan Status")
        dq = next(r for r in self.loan if r["field"] == "Delinquency Status")
        self.assertEqual(ls["liquidationDate"], "06/29/2026")
        self.assertGreater(float(ls["balance"]), 0.0)          # live balance
        self.assertEqual(ls["reported"], "ACTIVE")
        self.assertEqual(ls["priorLoanStatus"], "BANKRUPTCY")
        self.assertIsNotNone(ls["priorEndingBalance"])
        self.assertEqual(ls["bankruptcyDate"], "2026-03-12")
        self.assertEqual(dq["reported"], "120 DAYS")
        # 90 DAYS, corrected 2026-08-14 from 60 DAYS. With JUNE as the prior
        # finalized period the ladder starts one rung further along; the May-latest
        # pairing understated this and 51 other values by exactly one rung.
        self.assertEqual(dq["calculated"], "90 DAYS")
        for key in ("foreclosureDate", "reoDate", "liquidationProceeds",
                    "liquidationLoss", "subsequentRecoveries", "paymentReceived"):
            self.assertIn(key, ls)

    def test_the_reviewer_may_select_liquidation_and_it_validates(self):
        """A reviewer who weighs the evidence and concludes LIQUIDATION is correct
        must be able to say so -- explicitly, never by inference."""
        ok, value = corrections.validate_approved_value(
            "Loan Status", "LIQUIDATION", substituted=True)
        self.assertTrue(ok)
        self.assertEqual(value, "LIQUIDATION")

    def test_the_application_itself_still_cannot_generate_liquidation(self):
        ok, reason = corrections.validate_approved_value(
            "Loan Status", "LIQUIDATION", substituted=False)
        self.assertFalse(ok)
        self.assertIn("may generate", reason)

    def test_the_reviewer_may_select_any_of_the_nine_permitted_statuses(self):
        for status in sorted(corrections._PERMITTED_LOAN_STATUS):
            ok, value = corrections.validate_approved_value(
                "Loan Status", status, substituted=True)
            self.assertTrue(ok, f"{status} refused")
            self.assertEqual(value, status)

    def test_a_cross_domain_or_invented_status_is_refused(self):
        for bad in ("CURRENT", "60 DAYS", "CHARGE-OFF", "LIQUIDATED", ""):
            ok, _ = corrections.validate_approved_value(
                "Loan Status", bad, substituted=True)
            self.assertFalse(ok, f"{bad!r} was accepted as a Loan Status")

    def test_the_loan_status_candidate_still_requires_reviewer_approval_to_write(self):
        cand = next(c for c in self.candidates
                    if str(c["gp"]) == self.GP and c["field"] == "Loan Status")
        self.assertNotIn("reviewer_approved", cand)

    def test_july_classification_totals(self):
        """CANONICAL July baseline, measured against the JUNE-inclusive SOUP.

        CORRECTED 2026-08-14 from 89/57/32. That earlier figure was produced by
        pairing this July lender file with the MAY-latest SOUP -- the wrong prior
        finalized period. It is not a smaller version of this answer: 52 of its
        values were one delinquency rung understated, 2 were false positives
        against loans that cured in June, and the GP 6666709108 Loan Status
        contradiction was missing entirely. No business logic changed."""
        std = [r for r in self.recs if r["standard"]]
        self.assertEqual(len(self.recs), 91)
        self.assertEqual(len(std), 57)
        self.assertEqual(len(self.recs) - len(std), 34)

    def test_july_category_breakdown(self):
        from collections import Counter
        got = Counter(r["category"] for r in self.recs)
        self.assertEqual(got[corrections.STANDARD_CATEGORY], 57)
        self.assertEqual(got[corrections.REVIEW_CATEGORY_REVERSAL], 29)
        self.assertEqual(got[corrections.REVIEW_CATEGORY_LOAN_STATUS], 5)
        # nothing lands anywhere else in this population
        self.assertEqual(sum(got.values()), 91)

    def test_the_july_cycle_month_is_july_2026(self):
        """The reporting period the audit derived. Pairing this file with the
        May-latest SOUP yields 2026-06 and dates every output June_2026, which is
        the visible symptom of the wrong prior period."""
        self.assertEqual(str(self.cycle_month)[:7], "2026-07")

    def test_the_july_output_filenames_say_july_2026(self):
        self.assertEqual(corrections._filename_month_label(self.cycle_month), "July_2026")
        self.assertEqual(
            corrections._build_corrected_filename(self.servicer, self.cycle_month),
            "Lendistry_2025_July_2026_CORRECTED.xlsx")

    def test_every_july_loan_status_loan_pairs_with_a_delinquency_recommendation(self):
        """FIVE loans, corrected 2026-08-14 from four. 6666709108 is the one the
        May-latest pairing hid: its prior finalized status only becomes BANKRUPTCY
        once June is the prior period, and the current-period Bankruptcy Date
        branch is gated on LOAN STATUS MISMATCH, so with prior == reported ACTIVE
        no mismatch was raised and its populated Bankruptcy Date was never
        consulted. See docs/OPEN_QUESTION_bankruptcy_date_gating.md -- recorded,
        deliberately NOT fixed here."""
        from collections import defaultdict
        by_loan = defaultdict(set)
        for r in self.recs:
            by_loan[str(r["gp"])].add(r["field"])
        ls_loans = {gp for gp, fields in by_loan.items() if "Loan Status" in fields}
        self.assertEqual(ls_loans, {"5917099105", "6288639102", "6649779105",
                                    "6666709108", "7126119101"})
        for gp in ls_loans:
            self.assertEqual(by_loan[gp], {"Loan Status", "Delinquency Status"})
        # every one of them is Review Required, none bulk-approvable
        for r in self.recs:
            if r["field"] == "Loan Status":
                self.assertFalse(r["standard"])
                self.assertEqual(r["category"], corrections.REVIEW_CATEGORY_LOAN_STATUS)

    def test_gp_6666709108_is_the_second_contradictory_evidence_case(self):
        """A current-period Bankruptcy Date AND a Liquidation Date against a live
        balance -- the same shape as GP 6649779105, surfaced only once June is the
        prior finalized period."""
        loan = [r for r in self.recs if str(r["gp"]) == "6666709108"]
        self.assertEqual({r["field"] for r in loan},
                         {"Loan Status", "Delinquency Status"})
        ls = next(r for r in loan if r["field"] == "Loan Status")
        self.assertEqual(ls["reported"], "ACTIVE")
        self.assertEqual(ls["calculated"], "BANKRUPTCY")
        self.assertEqual(ls["priorLoanStatus"], "BANKRUPTCY")   # June, not May's ACTIVE
        self.assertEqual(ls["bankruptcyDate"], "2026-04-06")
        self.assertEqual(ls["liquidationDate"], "05/26/2026")
        self.assertGreater(float(ls["balance"]), 0.0)
        self.assertFalse(ls["standard"])


# ---------------------------------------------------------------------------
class ReportingPeriodPairing(unittest.TestCase):
    """THE GUARD. A lender file paired with the wrong SOUP generation does not
    error -- it returns a plausible, wrong answer. This is what makes that
    mistake detectable.

    CLIENT-CONFIRMED (2026-08-14):
        a JUNE audit's prior finalized period is MAY
        a JULY audit's prior finalized period is JUNE

    Asserted three ways, so no single accident defeats it:
      1. each configured SOUP's own latest Payment Month is the period it claims
         to be -- a swap of the two environment variables fails here immediately;
      2. the two SOUP generations are genuinely different datasets, not copies;
      3. the cycle month the audit DERIVES from the July file equals July only
         with the June-inclusive SOUP, and is demonstrably June with the other --
         which is the mispairing this suite previously shipped a baseline for.

    Reads facts only. Changes no calculation and asserts nothing about the
    reporting-period RULE itself, which has its own suite
    (test_reporting_period_rule.py)."""

    @staticmethod
    def _latest_month(soup_path):
        soup = file_readers_mod().read_soup_file(open(soup_path, "rb").read())
        months = pd.to_datetime(soup["Payment Month"], errors="coerce").dropna()
        return months.max(), len(soup)

    def test_the_june_soup_latest_period_is_may_2026(self):
        if _require(SOUP_XLSX):
            self.skipTest("set FALCONBRIDGE_SOUP_XLSX")
        latest, rows = self._latest_month(SOUP_XLSX)
        self.assertEqual(str(latest)[:7], "2026-05",
                         f"FALCONBRIDGE_SOUP_XLSX must be the MAY-latest dataset "
                         f"(a June audit's prior period); got {latest} over {rows} rows")

    def test_the_july_soup_latest_period_is_june_2026(self):
        if _require(JULY_SOUP_XLSX):
            self.skipTest("set FALCONBRIDGE_JULY_SOUP_XLSX")
        latest, rows = self._latest_month(JULY_SOUP_XLSX)
        self.assertEqual(str(latest)[:7], "2026-06",
                         f"FALCONBRIDGE_JULY_SOUP_XLSX must be the JUNE-inclusive "
                         f"dataset (a July audit's prior period); got {latest} "
                         f"over {rows} rows")

    def test_the_two_soup_generations_are_not_the_same_dataset(self):
        """If they were interchangeable there would be nothing to get wrong."""
        if _require(SOUP_XLSX, JULY_SOUP_XLSX):
            self.skipTest("set FALCONBRIDGE_SOUP_XLSX and FALCONBRIDGE_JULY_SOUP_XLSX")
        june_latest, june_rows = self._latest_month(SOUP_XLSX)
        july_latest, july_rows = self._latest_month(JULY_SOUP_XLSX)
        self.assertNotEqual(june_latest, july_latest)
        self.assertLess(june_latest, july_latest)
        self.assertLess(june_rows, july_rows)   # July's adds a period of history

    def test_the_july_file_derives_july_only_with_the_june_inclusive_soup(self):
        """The positive AND the negative case in one test: the correct pairing
        derives 2026-07, and the WRONG pairing derives 2026-06. Asserting the
        wrong case explicitly is the point -- it proves the mismatch is
        detectable rather than merely undocumented."""
        if _require(JULY_LENDISTRY_2025_XLSX, SOUP_XLSX, JULY_SOUP_XLSX, MEMO_XLSX):
            self.skipTest("set the June and July SOUP fixtures plus the July SRF")
        fr = file_readers_mod()
        derived = {}
        for tag, soup_path in (("june_inclusive", JULY_SOUP_XLSX),
                               ("may_latest", SOUP_XLSX)):
            soup = fr.read_soup_file(open(soup_path, "rb").read())
            memo = fr.resolve_memo_attribution(
                fr.read_service_memo(open(MEMO_XLSX, "rb").read()),
                fr.build_memo_loan_index(soup))
            k, l, lt, si, deal = fr.read_servicer_file(
                open(JULY_LENDISTRY_2025_XLSX, "rb").read(),
                filename="Lendistry||Lendistry 2025__"
                         + os.path.basename(JULY_LENDISTRY_2025_XLSX))
            res = checks.run_servicer_reconciliation(
                l, lt, si, fr.filter_soup_by_servicer(soup, k),
                memo_df=memo, memo_lender_key=k, deal=deal)
            stat01 = next(r for r in res if r["id"].split("::")[0] == "STAT-01")
            recs, _, _ = corrections.build_recommendations(stat01, l)
            derived[tag] = {
                "cycle": stat01.get("_cycle_month"),
                "label": corrections._filename_month_label(stat01.get("_cycle_month")),
                "total": len(recs),
                "standard": sum(1 for r in recs if r["standard"]),
            }
        # correct pairing
        self.assertEqual(str(derived["june_inclusive"]["cycle"])[:7], "2026-07")
        self.assertEqual(derived["june_inclusive"]["label"], "July_2026")
        self.assertEqual(derived["june_inclusive"]["total"], 91)
        self.assertEqual(derived["june_inclusive"]["standard"], 57)
        # wrong pairing -- must be DIFFERENT and must be visibly mis-dated
        self.assertEqual(str(derived["may_latest"]["cycle"])[:7], "2026-06")
        self.assertEqual(derived["may_latest"]["label"], "June_2026")
        self.assertNotEqual(derived["may_latest"]["total"],
                            derived["june_inclusive"]["total"])
        self.assertNotEqual(derived["may_latest"]["label"],
                            derived["june_inclusive"]["label"])


def file_readers_mod():
    import file_readers
    return file_readers


# ---------------------------------------------------------------------------
class ReviewerSelectedLoanStatusWritesThrough(unittest.TestCase):
    """The full chain for a reviewer-CHOSEN Loan Status, on a real workbook:

        dropdown selection -> validate_approved_value(substituted=True)
        -> reviewer_approved -> _apply_one_correction -> Correction Log
        -> verified_overrides -> Final Audited Portfolio Summary AND the MSR

    This is the requirement that the reported lender status must NOT be what the
    post-review reports carry once a reviewer has decided otherwise. Every link
    already existed; these tests pin that classification did not break any of it.
    """

    SHEET = "Servicer Inputs"
    HEADERS = ["SBA GP Number", "Lender Loan Number", "Loan Status",
               "Delinquency Status", "Payment Received", "Servicing Fee",
               "Scheduled Payment", "Beginning Balance", "Ending Balance",
               "Loan Payment Instance", "Last Payment Instance?"]
    GP = "6649779105"

    def _workbook(self):
        import io
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = self.SHEET
        for i, h in enumerate(self.HEADERS, start=1):
            ws.cell(row=HEADER_ROW, column=i, value=h)
        ws.cell(row=HEADER_ROW + 1, column=1, value=self.GP)
        ws.cell(row=HEADER_ROW + 1, column=2, value="L-1")
        ws.cell(row=HEADER_ROW + 1, column=3, value="ACTIVE")
        ws.cell(row=HEADER_ROW + 1, column=4, value="120 DAYS")
        ws.cell(row=HEADER_ROW + 1, column=5, value=0.0)
        ws.cell(row=HEADER_ROW + 1, column=6, value=0.0)
        ws.cell(row=HEADER_ROW + 1, column=7, value=500.0)
        ws.cell(row=HEADER_ROW + 1, column=8, value=7450.69)
        ws.cell(row=HEADER_ROW + 1, column=9, value=7450.69)
        ws.cell(row=HEADER_ROW + 1, column=10, value=1)
        ws.cell(row=HEADER_ROW + 1, column=11, value="Y")
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def _candidates(self):
        """One loan, TWO recommendations -- exactly the grouped-card shape."""
        common = {
            "gp": self.GP, "excel_rows": [HEADER_ROW + 1], "skipped_row_refs": [],
            "sheet_name": self.SHEET, "header_excel_row": HEADER_ROW,
            "review_required": True, "evidence_marker": None,
        }
        return [
            {**common, "id": f"Lendistry 2025||{self.GP}||Loan Status",
             "field": "Loan Status", "corrected_value": "BANKRUPTCY",
             "reported_value": "ACTIVE",
             "review_reason": corrections.REVIEW_REASON_BANKRUPTCY_LIQUIDATION_CONTRADICTION,
             "eligibility_rule": "Current-period Bankruptcy Date populated",
             "evidence_marker": corrections.CURRENT_PERIOD_EVIDENCE_MARKER},
            {**common, "id": f"Lendistry 2025||{self.GP}||Delinquency Status",
             "field": "Delinquency Status", "corrected_value": "60 DAYS",
             "reported_value": "120 DAYS", "review_required": False,
             "review_reason": None, "eligibility_rule": "STAT-01 deterministic calculation"},
        ]

    def _review(self, decisions):
        return corrections.generate_reviewed_workbook(
            self._workbook(), "Lendistry_July.xlsx", "Lendistry 2025",
            self._candidates(), decisions, "Investigator", "sha-test",
            pd.Timestamp("2026-07-01"))

    def _log_rows(self, file_bytes):
        import io
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws = wb[corrections.CORRECTION_LOG_SHEET_NAME]
        headers = [c.value for c in ws[1]]
        return [dict(zip(headers, [c.value for c in row]))
                for row in ws.iter_rows(min_row=2)]

    def test_a_reviewer_selected_liquidation_is_written_to_the_workbook(self):
        """The reviewer weighed the contradictory evidence and chose LIQUIDATION
        over the application's BANKRUPTCY proposal."""
        import io
        from openpyxl import load_workbook
        out = self._review([
            {"id": f"Lendistry 2025||{self.GP}||Loan Status",
             "decision": "approve", "approved_value": "LIQUIDATION"},
        ])
        self.assertEqual(out["applied_count"], 1)
        self.assertEqual(out["corrections_not_applied"], [])
        ws = load_workbook(io.BytesIO(out["file_bytes"]))[self.SHEET]
        self.assertEqual(ws.cell(row=HEADER_ROW + 1, column=3).value, "LIQUIDATION")

    def test_the_correction_log_records_the_reviewers_actual_choice(self):
        out = self._review([
            {"id": f"Lendistry 2025||{self.GP}||Loan Status",
             "decision": "approve", "approved_value": "LIQUIDATION",
             "reviewer_note": "liquidation date confirmed with servicer"},
        ])
        row = next(r for r in self._log_rows(out["file_bytes"])
                   if r["Field Updated"] == "Loan Status")
        self.assertEqual(row["Calculated Value"], "BANKRUPTCY")     # what the app proposed
        self.assertEqual(row["Reviewer Approved Value"], "LIQUIDATION")  # what the human chose
        self.assertEqual(row["Corrected Value"], "LIQUIDATION")
        self.assertEqual(row["Original Value"], "ACTIVE")
        self.assertEqual(row["Rows Updated"], 1)
        self.assertIn("Approved", str(row["Decision"]))
        self.assertIn("liquidation date confirmed", str(row["Reviewer Note"]))
        self.assertIn("Investigator", str(row["Reviewed By (self-declared)"]))

    def test_the_choice_lands_in_the_verified_override_layer(self):
        out = self._review([
            {"id": f"Lendistry 2025||{self.GP}||Loan Status",
             "decision": "approve", "approved_value": "LIQUIDATION"},
        ])
        self.assertEqual(out["verified_overrides"],
                         {(self.GP, "Loan Status"): "LIQUIDATION"})

    def test_the_final_audited_summary_uses_the_reviewed_status(self):
        """Not the lender's ACTIVE -- the reviewer's LIQUIDATION."""
        si = pd.DataFrame([{
            "SBA GP Number": self.GP, "Loan Status": "ACTIVE",
            "Delinquency Status": "120 DAYS", "Payment Received": 0.0,
            "Servicing Fee": 0.0, "Scheduled Payment": 500.0,
            "Beginning Balance": 7450.69, "Ending Balance": 7450.69,
            "Loan Payment Instance": 1, "Last Payment Instance?": "Y",
            "_excel_row": HEADER_ROW + 1, "_sheet_name": self.SHEET,
            "_header_excel_row": HEADER_ROW,
        }])
        lender = checks.verified_portfolio_summary_views(si, {}, "Lendistry 2025", None)
        reviewed = checks.verified_portfolio_summary_views(
            si, {(self.GP, "Loan Status"): "LIQUIDATION"}, "Lendistry 2025", None)
        self.assertEqual(lender["counts"].get("Liquidation", 0), 0)
        self.assertEqual(reviewed["counts"].get("Liquidation", 0), 1)
        # The dollars record keys are " ($)"-suffixed, and Liquidation/Paid In Full
        # dollars are the BEGINNING balance (checks.py -- the balance that left the
        # portfolio), not the ending one. Asserting the existing rule, not a new one.
        self.assertAlmostEqual(reviewed["dollars"].get("Liquidation ($)", 0.0),
                               7450.69, places=2)
        self.assertAlmostEqual(lender["dollars"].get("Liquidation ($)", 0.0),
                               0.0, places=2)

    def test_the_monthly_servicer_report_uses_the_reviewed_status(self):
        import monthly_servicer_report as msr
        si = pd.DataFrame([{
            "SBA GP Number": self.GP, "Lender Loan Number": "L-1",
            "Loan Status": "ACTIVE", "Delinquency Status": "120 DAYS",
            "Payment Received": 0.0, "Servicing Fee": 0.0, "Principal": 0.0,
            "Net Interest": 0.0, "Scheduled Payment": 500.0,
            "Beginning Balance": 7450.69, "Ending Balance": 7450.69,
            "Loan Payment Instance": 1, "Last Payment Instance?": "Y",
            "_excel_row": HEADER_ROW + 1, "_sheet_name": self.SHEET,
            "_header_excel_row": HEADER_ROW,
        }])
        lender_rows = msr.build_servicer_rows(si, "Lendistry 2025", {})
        reviewed_rows = msr.build_servicer_rows(
            si, "Lendistry 2025", {(self.GP, "Loan Status"): "LIQUIDATION"})

        def status_of(rows):
            for row in rows:
                for key, value in row.items():
                    if "Loan Status" in str(key):
                        return value
            return None

        self.assertEqual(status_of(lender_rows), "ACTIVE")
        self.assertEqual(status_of(reviewed_rows), "LIQUIDATION")

    def test_a_rejected_loan_status_leaves_the_lender_value_and_no_override(self):
        import io
        from openpyxl import load_workbook
        out = self._review([
            {"id": f"Lendistry 2025||{self.GP}||Loan Status", "decision": "reject"},
        ])
        self.assertEqual(out["applied_count"], 0)
        self.assertEqual(out["rejected_count"], 1)
        self.assertEqual(out["verified_overrides"], {})
        ws = load_workbook(io.BytesIO(out["file_bytes"]))[self.SHEET]
        self.assertEqual(ws.cell(row=HEADER_ROW + 1, column=3).value, "ACTIVE")
        row = next(r for r in self._log_rows(out["file_bytes"])
                   if r["Field Updated"] == "Loan Status")
        self.assertEqual(row["Rows Updated"], 0)
        self.assertIn("Rejected", str(row["Decision"]))

    def test_the_two_recommendations_are_decided_independently(self):
        """Grouping is display only. Approving the Loan Status must not touch the
        Delinquency Status, and each is logged separately."""
        import io
        from openpyxl import load_workbook
        out = self._review([
            {"id": f"Lendistry 2025||{self.GP}||Loan Status",
             "decision": "approve", "approved_value": "LIQUIDATION"},
            {"id": f"Lendistry 2025||{self.GP}||Delinquency Status", "decision": "reject"},
        ])
        self.assertEqual(out["applied_count"], 1)
        self.assertEqual(out["rejected_count"], 1)
        ws = load_workbook(io.BytesIO(out["file_bytes"]))[self.SHEET]
        self.assertEqual(ws.cell(row=HEADER_ROW + 1, column=3).value, "LIQUIDATION")
        self.assertEqual(ws.cell(row=HEADER_ROW + 1, column=4).value, "120 DAYS")  # untouched
        self.assertEqual(out["verified_overrides"],
                         {(self.GP, "Loan Status"): "LIQUIDATION"})
        fields = {r["Field Updated"] for r in self._log_rows(out["file_bytes"])}
        self.assertEqual(fields, {"Loan Status", "Delinquency Status"})

    def test_an_invented_status_is_refused_and_never_written(self):
        import io
        from openpyxl import load_workbook
        out = self._review([
            {"id": f"Lendistry 2025||{self.GP}||Loan Status",
             "decision": "approve", "approved_value": "LIQUIDATED"},
        ])
        self.assertEqual(out["applied_count"], 0)
        self.assertEqual(out["approved_not_applied_count"], 1)
        self.assertEqual(out["verified_overrides"], {})
        ws = load_workbook(io.BytesIO(out["file_bytes"]))[self.SHEET]
        self.assertEqual(ws.cell(row=HEADER_ROW + 1, column=3).value, "ACTIVE")


# ---------------------------------------------------------------------------
class BulkApprovalCannotBypassD4(test_review_completeness._Harness):
    """Bulk-approving every Standard correction must leave FINAL REPORTING blocked
    while a Review Required item is still pending.

    Deliberately driven through test_review_completeness's OWN harness rather than
    a private re-implementation: the thing under test is D4's real arithmetic on
    the real endpoint, and a second driver could drift from it and quietly assert
    against a shape D4 no longer produces.

    Candidate 0 of each servicer is turned into a Loan Status recommendation --
    precisely what bulk approval may never touch -- so "approve everything the
    bulk control could reach" is exactly "everything except candidate 0"."""

    LOAN_STATUS_INDEX = 0

    def _approve_all_but_loan_status(self, cand_id, index):
        # Mirrors the page: the bulk control sweeps Standard items only, so the
        # Loan Status of each servicer is left undecided.
        return None if index % PER_SERVICER == self.LOAN_STATUS_INDEX else "approve"

    def test_bulk_approving_the_standard_group_leaves_reporting_blocked(self):
        status, body, captured = self.run_review(self._approve_all_but_loan_status)
        self.assertEqual(status, 200)
        rc, rr = body["review_completeness"], body["reporting_readiness"]
        servicers = len(ROSTER_SIZE)
        self.assertEqual(rc["total"], servicers * PER_SERVICER)
        self.assertEqual(rc["pending"], servicers)          # one per servicer
        self.assertFalse(rc["review_complete"])
        self.assertFalse(rr["reporting_ready"])
        self.assertIn("pending", rr["blocking_reasons"])
        # BOTH population reports withheld...
        self.assertEqual(body.get("monthly_servicer_reports") or [], [])
        self.assertIsNone(body.get("portfolio_summary_final_b64"))
        self.assertNotIn("called", captured)
        # ...while the PER-DECISION artifacts are still returned.
        self.assertEqual(len(body.get("corrected_workbooks") or []), servicers)

    def test_deciding_the_loan_status_too_releases_reporting(self):
        status, body, captured = self.run_review(lambda cid, i: "approve")
        self.assertEqual(status, 200)
        rc, rr = body["review_completeness"], body["reporting_readiness"]
        self.assertEqual(rc["pending"], 0)
        self.assertTrue(rc["review_complete"])
        self.assertTrue(rr["reporting_ready"])
        self.assertEqual(rr["non_reportable"], 0)
        self.assertTrue(body.get("monthly_servicer_reports"))
        self.assertTrue(body.get("portfolio_summary_final_b64"))

    def test_reportable_plus_non_reportable_always_equals_total(self):
        for decide in (self._approve_all_but_loan_status,
                       lambda cid, i: "approve",
                       lambda cid, i: "reject" if i % 2 else None):
            _, body, _ = self.run_review(decide)
            rc, rr = body["review_completeness"], body["reporting_readiness"]
            self.assertEqual(rr["reportable"] + rr["non_reportable"], rc["total"])
            if rr["reporting_ready"]:
                self.assertTrue(rc["review_complete"])





if __name__ == "__main__":
    unittest.main()
