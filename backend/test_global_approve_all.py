"""Global Approve All — CLIENT-CONFIRMED (2026-08-17).

The review page gains a GLOBAL bulk pre-fill alongside the existing Standard
one. This file holds the backend half: the category/field scope it is allowed
to sweep, and the parity assertions that hold the JSX to it.

WHAT THIS IS NOT. BULK_APPROVABLE_CATEGORIES is DATA. It changes no
classification, no eligibility, no write gate and no validation.
classify_recommendation, _apply_one_correction's reviewer_approved gate and
validate_approved_value are all untouched — a bulk pre-fill still produces an
ordinary submitted decision that the server validates exactly as it validates a
hand-entered one. Nothing here creates a path around backend validation.

THE TWO CONDITIONS THAT ARE NOT CATEGORY MEMBERSHIP:

  * BULK_EXCLUDED_FIELDS — the required liquidation-reversal balance decision
    classifies INTO REVIEW_CATEGORY_LOAN_STATUS by design, so admitting that
    category would otherwise admit the balance decision with it. Choosing
    between the two server-supplied balance options is a financial judgement.
  * an engine value must exist — a mandatory lender-reported LIQUIDATION has no
    proposable status, so there is nothing to pre-fill.
"""

import os
import re
import unittest

import corrections


JSX = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "frontend", "src",
                   "components", "CorrectionReview.jsx")


class BulkApprovableScopeTests(unittest.TestCase):
    """What a global bulk action may and may not sweep."""

    def test_the_four_delinquency_categories_are_bulk_approvable(self):
        for category in (corrections.STANDARD_CATEGORY,
                         corrections.REVIEW_CATEGORY_DATA_DOMAIN,
                         corrections.REVIEW_CATEGORY_TERMINAL,
                         corrections.REVIEW_CATEGORY_REVERSAL):
            with self.subTest(category=category):
                self.assertIn(category, corrections.BULK_APPROVABLE_CATEGORIES)

    def test_loan_status_is_bulk_approvable(self):
        """Where the engine produced a recommended value, a bulk pre-fill
        writes that value and nothing else."""
        self.assertIn(corrections.REVIEW_CATEGORY_LOAN_STATUS,
                      corrections.BULK_APPROVABLE_CATEGORIES)

    def test_other_flagged_is_never_bulk_approvable(self):
        """The fail-closed bucket. A review reason added upstream in future
        lands here by default and must not be swept before classification."""
        self.assertNotIn(corrections.REVIEW_CATEGORY_OTHER,
                         corrections.BULK_APPROVABLE_CATEGORIES)

    def test_every_category_is_explicitly_in_or_out(self):
        every = set(corrections.REVIEW_REQUIRED_CATEGORIES) | {corrections.STANDARD_CATEGORY}
        classified = set(corrections.BULK_APPROVABLE_CATEGORIES) | {
            corrections.REVIEW_CATEGORY_OTHER}
        self.assertEqual(every, classified)

    def test_the_balance_decision_field_is_excluded_by_field_not_category(self):
        """The load-bearing detail: it shares its sibling's category, so only a
        field-level exclusion keeps it out."""
        self.assertIn(corrections.BALANCE_DECISION_FIELD, corrections.BULK_EXCLUDED_FIELDS)
        self.assertEqual(
            corrections.classify_recommendation(
                {"field": corrections.BALANCE_DECISION_FIELD})[0],
            corrections.REVIEW_CATEGORY_LOAN_STATUS)
        self.assertIn(corrections.REVIEW_CATEGORY_LOAN_STATUS,
                      corrections.BULK_APPROVABLE_CATEGORIES)

    def test_a_flagged_recommendation_classifies_outside_the_bulk_scope(self):
        category, standard = corrections.classify_recommendation(
            {"field": "Delinquency Status", "review_required": True})
        self.assertFalse(standard)
        self.assertEqual(category, corrections.REVIEW_CATEGORY_OTHER)
        self.assertNotIn(category, corrections.BULK_APPROVABLE_CATEGORIES)

    def test_the_scope_lists_are_data_only(self):
        """No behaviour hangs off them server-side; they exist for parity."""
        self.assertIsInstance(corrections.BULK_APPROVABLE_CATEGORIES, tuple)
        self.assertIsInstance(corrections.BULK_EXCLUDED_FIELDS, tuple)
        source = corrections.classify_recommendation.__doc__ or ""
        self.assertNotIn("BULK_APPROVABLE_CATEGORIES", source)

    def test_classification_precedence_is_unchanged(self):
        """Guards against this work having quietly moved an item into Standard."""
        self.assertEqual(
            corrections.classify_recommendation({"reversal_equivalents": []})[0],
            corrections.REVIEW_CATEGORY_REVERSAL)
        self.assertEqual(
            corrections.classify_recommendation({"field": "Loan Status"})[0],
            corrections.REVIEW_CATEGORY_LOAN_STATUS)
        self.assertEqual(
            corrections.classify_recommendation({"field": "Delinquency Status"}),
            (corrections.STANDARD_CATEGORY, True))


class JsxParityTests(unittest.TestCase):
    """The JSX mirrors the backend lists verbatim, asserted against the source
    the same way the existing category and domain parity tests are."""

    @classmethod
    def setUpClass(cls):
        with open(JSX, encoding="utf-8") as handle:
            cls.src = handle.read()

    def _resolved_array(self, name):
        m = re.search(rf"const {name} = \[(.*?)\]", self.src, re.S)
        self.assertIsNotNone(m, f"{name} not found in CorrectionReview.jsx")
        names = re.findall(r"(STANDARD_CATEGORY|REVIEW_CATEGORY_\w+)", m.group(1))
        return [re.search(rf"const {n} = '([^']+)'", self.src).group(1) for n in names]

    def test_the_jsx_bulk_scope_matches_the_backend_exactly(self):
        self.assertEqual(tuple(self._resolved_array("BULK_APPROVABLE_CATEGORIES")),
                         corrections.BULK_APPROVABLE_CATEGORIES)

    def test_the_jsx_bulk_scope_excludes_other_flagged(self):
        self.assertNotIn(corrections.REVIEW_CATEGORY_OTHER,
                         self._resolved_array("BULK_APPROVABLE_CATEGORIES"))

    def test_the_jsx_bulk_scope_includes_loan_status(self):
        self.assertIn(corrections.REVIEW_CATEGORY_LOAN_STATUS,
                      self._resolved_array("BULK_APPROVABLE_CATEGORIES"))

    def test_the_global_action_excludes_the_balance_decision(self):
        """requiresBalanceDecision is checked before anything else, so no
        balance can be pre-filled whatever its category."""
        m = re.search(r"const globalHoldReason = \(rec\) => \{(.*?)\n  \}", self.src, re.S)
        self.assertIsNotNone(m, "globalHoldReason not found in CorrectionReview.jsx")
        body = m.group(1)
        self.assertIn("rec.requiresBalanceDecision", body)
        self.assertLess(body.index("requiresBalanceDecision"), body.index("bulkApprovable"))

    def test_the_global_action_requires_an_engine_value(self):
        m = re.search(r"const globalHoldReason = \(rec\) => \{(.*?)\n  \}", self.src, re.S)
        self.assertIn("!rec.calculated", m.group(1))

    def test_the_global_action_preserves_the_companion_hold(self):
        m = re.search(r"const globalHoldReason = \(rec\) => \{(.*?)\n  \}", self.src, re.S)
        self.assertIn("loanStatusUndecided(rec)", m.group(1))

    def test_the_global_action_never_overwrites_a_decision(self):
        m = re.search(r"const approveAllGlobal = \(\) =>(.*?)\n\n", self.src, re.S)
        self.assertIsNotNone(m, "approveAllGlobal not found in CorrectionReview.jsx")
        self.assertIn("if (!(out[r.id] && out[r.id].decision))", m.group(1))

    def test_the_global_action_submits_nothing(self):
        m = re.search(r"const approveAllGlobal = \(\) =>(.*?)\n\n", self.src, re.S)
        body = m.group(1)
        for forbidden in ("axios", "fetch(", "submit(", "/api/"):
            with self.subTest(symbol=forbidden):
                self.assertNotIn(forbidden, body)

    def test_the_global_action_pre_fills_only_the_engine_value(self):
        """approved_value comes from r.calculated and is never computed."""
        m = re.search(r"const approveAllGlobal = \(\) =>(.*?)\n\n", self.src, re.S)
        self.assertIn("approved_value: r.calculated", m.group(1))

    def test_the_standard_action_is_unchanged(self):
        self.assertIn("const approveAllStandard = () =>", self.src)
        self.assertIn("Approve All Standard Corrections", self.src)
        m = re.search(r"const approveAllStandard = \(\) =>(.*?)\n\n", self.src, re.S)
        self.assertIn("bulkEligible.forEach", m.group(1))
        # It still pre-fills the decision only -- no value, exactly as before.
        self.assertNotIn("approved_value", m.group(1))

    def test_the_still_requires_decision_panel_exists(self):
        self.assertIn("Still Requires Your Decision", self.src)
        for reason in ("HOLD_REASON_BALANCE", "HOLD_REASON_NO_VALUE",
                       "HOLD_REASON_OTHER_FLAGGED", "HOLD_REASON_COMPANION"):
            with self.subTest(reason=reason):
                self.assertIn(reason, self.src)


if __name__ == "__main__":
    unittest.main()
