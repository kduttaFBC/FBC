"""Portfolio Summary "Not Making Payment": the filed-posting corroboration rule.

WHAT THIS COVERS
----------------
CLIENT-CONFIRMED (2026-08-12). "Not Making Payment" has two ways in:

    (a) consolidated Payment Received is PRESENT and exactly 0     [pre-existing]
    (b) consolidated Payment Received is BLANK, and
          Principal is blank or 0, and
          Beginning Balance == Ending Balance, and
          has_filed_posting is True

(b) exists because a blank Payment Received is ambiguous on its own. Measured on
Lendistry's July 2026 file: 37 loans reported a single FILED posting with every
payment field empty, no balance movement, and 120+ day delinquency on 35 of
them -- and 36 of those 37 had been filed by the same servicer as an explicit 0
the month before, for the same loans in the same condition. The servicer changed
how it writes "no payment this cycle"; nothing about the loans changed. Treating
that blank as "unknown" dropped $2,241,118.06 the audit had counted in June.

The discriminator is the SRF template's own posting-identity fields, never the
servicer name:

    Last Payment Instance? == "Y"            on any posting, OR
    Loan Payment Instance numerically set    on any posting

A blank payment field on a posting the servicer actually FILED asserts "no
payment". The same blank on a row carrying no posting identity asserts nothing
and stays unknown -- which is what keeps 2 unvalidated Lendistry-2024 loans
(GP 4490649110, GP 5462759105) correctly excluded.

WHAT IT DELIBERATELY DOES NOT COVER
-----------------------------------
Reversal handling, Human Review routing, the delinquency ladder, Loan Status
derivation and every other Portfolio Summary category are untouched by this
change, and this file asserts their invariance rather than re-testing them.

The real-file acceptance figures (Lendistry 2025 July: 60 loans /
$3,593,725.78) are NOT asserted here -- the July Lendistry SRF files are not
available as committed fixtures. The profiles below are built to match those
rows field-for-field so the rule itself is covered without them.
"""

import unittest

import pandas as pd

import checks


NMP = "Not Making Payment ($)"


def _posting(**over):
    """One Servicer Inputs posting. Defaults to a FILED posting reporting no
    payment activity -- the shape of all 37 Lendistry-2025 July rows."""
    row = {
        "SBA GP Number": "1000000001",
        "Lender Loan Number": "L-1",
        "Loan Status": "ACTIVE",
        "Delinquency Status": "120 DAYS",
        "Payment Received": "",
        "Principal": "",
        "Net Interest": "",
        "Scheduled Payment": "",
        "Payment Receipt Date": "",
        "Beginning Balance": 20000.0,
        "Ending Balance": 20000.0,
        "Loan Payment Instance": 1.0,
        "Last Payment Instance?": "Y",
        "_excel_row": 10,
    }
    row.update(over)
    return row


def _summary(postings):
    si = pd.DataFrame(postings)
    return checks.portfolio_summary_check(si, "Test Servicer")["records"][0]


def _views(postings):
    return checks.original_portfolio_summary_views(pd.DataFrame(postings), "Test Servicer")


class HasFiledPostingFactTests(unittest.TestCase):
    """The additive consolidation fact. A RAW FACT, never a judgement."""

    def _fact(self, postings):
        g = checks.consolidate_cycle_postings(pd.DataFrame(postings))
        return bool(g.iloc[0]["has_filed_posting"])

    def test_last_payment_instance_y_alone_is_enough(self):
        self.assertTrue(self._fact([_posting(**{"Loan Payment Instance": ""})]))

    def test_loan_payment_instance_alone_is_enough(self):
        self.assertTrue(self._fact([_posting(**{"Last Payment Instance?": ""})]))

    def test_neither_field_populated_is_false(self):
        self.assertFalse(self._fact([_posting(**{
            "Loan Payment Instance": "", "Last Payment Instance?": ""})]))

    def test_it_is_true_when_any_posting_carries_identity(self):
        # Matches has_negative_posting's "any posting in the cycle" semantics: a
        # reversal pair can carry the ordinal on one row and not the other.
        self.assertTrue(self._fact([
            _posting(**{"Loan Payment Instance": "", "Last Payment Instance?": "",
                        "Payment Received": 100.0}),
            _posting(**{"Loan Payment Instance": 2.0, "Last Payment Instance?": "Y",
                        "Payment Received": -100.0}),
        ]))

    def test_a_non_y_flag_does_not_count(self):
        self.assertFalse(self._fact([_posting(**{
            "Last Payment Instance?": "N", "Loan Payment Instance": ""})]))

    def test_a_non_numeric_instance_does_not_count(self):
        self.assertFalse(self._fact([_posting(**{
            "Last Payment Instance?": "", "Loan Payment Instance": "n/a"})]))

    def test_missing_columns_degrade_to_false_without_raising(self):
        row = _posting()
        del row["Loan Payment Instance"]
        del row["Last Payment Instance?"]
        self.assertFalse(self._fact([row]))

    def test_the_fact_changes_no_other_consolidated_value(self):
        # The whole point of "additive": every pre-existing field is identical
        # whether or not the identity columns are present.
        with_ident = checks.consolidate_cycle_postings(pd.DataFrame([_posting(
            **{"Payment Received": 250.0, "Principal": 100.0})]))
        row = _posting(**{"Payment Received": 250.0, "Principal": 100.0})
        del row["Loan Payment Instance"]
        del row["Last Payment Instance?"]
        without = checks.consolidate_cycle_postings(pd.DataFrame([row]))
        for col in ("received", "principal", "loan_status", "dq_status",
                    "beginning_balance", "ending_balance", "scheduled",
                    "posting_count", "has_negative_posting"):
            with self.subTest(column=col):
                self.assertEqual(with_ident.iloc[0][col], without.iloc[0][col])


class BlankPaymentIsClassifiedOnlyWhenCorroboratedTests(unittest.TestCase):
    """Clause (b), positive and negative."""

    def test_a_filed_no_activity_posting_counts(self):
        # The 37-loan profile.
        self.assertEqual(_summary([_posting()])[NMP], 20000.0)

    def test_an_unfiled_row_does_not_count(self):
        # The 2-loan Lendistry-2024 profile: balances present, no posting
        # identity, no payment data. Genuinely unknown -- must stay out.
        self.assertEqual(_summary([_posting(**{
            "Loan Payment Instance": "", "Last Payment Instance?": "",
            "Delinquency Status": "DELINQUENT"})])[NMP], 0.0)

    def test_blank_payment_with_real_principal_does_not_count(self):
        self.assertEqual(_summary([_posting(**{"Principal": 500.0})])[NMP], 0.0)

    def test_blank_payment_with_balance_movement_does_not_count(self):
        self.assertEqual(_summary([_posting(**{"Ending Balance": 19500.0})])[NMP], 0.0)

    def test_a_zero_principal_still_counts(self):
        # "blank OR zero" -- an explicit 0 principal is not movement.
        self.assertEqual(_summary([_posting(**{"Principal": 0.0})])[NMP], 20000.0)

    def test_every_corroboration_condition_is_load_bearing(self):
        # Drop one condition at a time; each omission alone must disqualify.
        for label, over in (
            ("principal", {"Principal": 500.0}),
            ("movement", {"Ending Balance": 19000.0}),
            ("filed", {"Loan Payment Instance": "", "Last Payment Instance?": ""}),
        ):
            with self.subTest(condition=label):
                self.assertEqual(_summary([_posting(**over)])[NMP], 0.0)


class ExistingExplicitZeroBehaviourIsPreservedTests(unittest.TestCase):
    """Clause (a) is untouched -- including the reversal-pair case, which is
    what 23 of the 60 Lendistry-2025 loans enter through."""

    def test_an_explicit_zero_still_counts(self):
        self.assertEqual(_summary([_posting(**{"Payment Received": 0.0})])[NMP], 20000.0)

    def test_a_reversal_pair_netting_to_zero_still_counts(self):
        # +299.99 then -299.99, balance restored. No reversal logic is involved
        # or changed: the consolidated net is 0, so clause (a) admits it exactly
        # as it did before this rule existed.
        rows = [
            _posting(**{"Payment Received": 299.99, "Principal": 137.65,
                        "Beginning Balance": 20000.0, "Ending Balance": 19862.35,
                        "Loan Payment Instance": 1.0, "Last Payment Instance?": "N"}),
            _posting(**{"Payment Received": -299.99, "Principal": -137.65,
                        "Beginning Balance": 19862.35, "Ending Balance": 20000.0,
                        "Loan Payment Instance": 2.0, "Last Payment Instance?": "Y"}),
        ]
        self.assertEqual(_summary(rows)[NMP], 20000.0)

    def test_a_real_payment_never_counts(self):
        self.assertEqual(_summary([_posting(**{
            "Payment Received": 250.0, "Principal": 100.0,
            "Ending Balance": 19900.0})])[NMP], 0.0)

    def test_an_unfiled_row_with_an_explicit_zero_still_counts(self):
        # has_filed_posting gates ONLY the blank branch. An explicitly reported
        # $0 is already unambiguous and must not start depending on it.
        self.assertEqual(_summary([_posting(**{
            "Payment Received": 0.0,
            "Loan Payment Instance": "", "Last Payment Instance?": ""})])[NMP], 20000.0)


class EligibilityGateStillWinsTests(unittest.TestCase):
    """The (beginning == 0 AND ending == 0) gate runs BEFORE categorisation and
    is not weakened by the new branch."""

    def test_a_zero_zero_loan_is_excluded_despite_qualifying(self):
        rec = _summary([_posting(**{"Beginning Balance": 0.0, "Ending Balance": 0.0})])
        self.assertEqual(rec[NMP], 0.0)
        self.assertEqual(rec["Outstanding ($)"], 0.0)


class CountsAndDollarsCannotDivergeTests(unittest.TestCase):
    """_unit_weight_frame rewrites both balance columns to 1.0, so the
    "Beginning == Ending" corroboration would be trivially TRUE in the counts
    pass. Without the _NMP_BALANCES_UNCHANGED passthrough, counts would classify
    a loan the dollars pass rejected. These are the tests that fail if that
    passthrough is removed."""

    def test_a_moving_balance_is_rejected_by_BOTH_views(self):
        rows = [_posting(**{"Ending Balance": 19000.0})]
        v = _views(rows)
        self.assertEqual(v["dollars"][NMP], 0.0)
        self.assertEqual(v["counts"]["Not Making Payment"], 0,
                         "counts must not admit a loan the dollars pass rejected")

    def test_a_qualifying_loan_is_accepted_by_BOTH_views(self):
        v = _views([_posting()])
        self.assertEqual(v["dollars"][NMP], 20000.0)
        self.assertEqual(v["counts"]["Not Making Payment"], 1)

    def test_the_passthrough_survives_unit_weighting(self):
        g = checks.consolidate_cycle_postings(pd.DataFrame([
            _posting(**{"SBA GP Number": "1000000001"}),
            _posting(**{"SBA GP Number": "1000000002", "Ending Balance": 19000.0}),
        ]))
        w = checks._unit_weight_frame(g)
        self.assertIn(checks._NMP_BALANCES_UNCHANGED, w.columns)
        self.assertTrue(bool(w.loc["1000000001", checks._NMP_BALANCES_UNCHANGED]))
        self.assertFalse(bool(w.loc["1000000002", checks._NMP_BALANCES_UNCHANGED]))
        # ...and the balances really were overwritten, which is why it is needed.
        self.assertEqual(float(w.loc["1000000002", "beginning_balance"]), 1.0)
        self.assertEqual(float(w.loc["1000000002", "ending_balance"]), 1.0)

    def test_unit_weight_frame_still_does_not_mutate_its_input(self):
        g = checks.consolidate_cycle_postings(pd.DataFrame([_posting()]))
        before = list(g.columns)
        checks._unit_weight_frame(g)
        self.assertEqual(list(g.columns), before)


class NoOtherCategoryIsAffectedTests(unittest.TestCase):
    """A newly-counted loan must land in Not Making Payment and NOTHING else."""

    def test_the_new_branch_touches_no_other_category(self):
        rec_off = _summary([_posting(**{
            "Loan Payment Instance": "", "Last Payment Instance?": ""})])
        rec_on = _summary([_posting()])
        for key in rec_on:
            if key == NMP or key.endswith("(%)"):
                continue
            with self.subTest(category=key):
                self.assertEqual(rec_on[key], rec_off[key])

    def test_the_delinquency_bucket_is_unchanged_by_the_new_branch(self):
        # 120 DAYS on both; being newly counted as Not Making Payment must not
        # move, duplicate, or drop the aging bucket.
        self.assertEqual(_summary([_posting()])["120 Days ($)"], 20000.0)

    def test_loan_status_categories_are_unchanged_by_the_new_branch(self):
        rec = _summary([_posting(**{"Loan Status": "DEFERRED"})])
        self.assertEqual(rec["Deferred ($)"], 20000.0)
        self.assertEqual(rec[NMP], 20000.0)   # overlapping, as designed
        self.assertEqual(rec["Bankruptcy ($)"], 0.0)
        self.assertEqual(rec["Foreclosure ($)"], 0.0)


class AuditedAndOriginalStayIdenticalTests(unittest.TestCase):
    """Not Making Payment never reads dq_status_for_aging_fn, so the Original
    and Audited summaries must report it identically -- the guarantee
    _build_portfolio_summary_record exists to provide."""

    def test_audited_matches_original_for_this_category(self):
        si = pd.DataFrame([_posting()])
        original = checks.portfolio_summary_check(si, "S")["records"][0]
        audited = checks.audited_portfolio_summary_record(
            si, {"1000000001": "CURRENT"}, "S")
        self.assertEqual(audited[NMP], original[NMP])
        # ...while the aging bucket legitimately differs, proving the audited
        # path really was exercised.
        self.assertEqual(audited["Current ($)"], 20000.0)
        self.assertEqual(original["120 Days ($)"], 20000.0)


class RealWorldProfilesTests(unittest.TestCase):
    """The two measured July-2026 profiles, field-for-field, so the rule is
    covered even though the source workbooks are not available as fixtures."""

    def test_the_lendistry_2025_profile_counts(self):
        # GP 4609059108 as filed: single posting, every payment field empty,
        # zero-length interest period, no balance movement, identity present.
        rec = _summary([_posting(**{
            "SBA GP Number": "4609059108", "Delinquency Status": "150 DAYS",
            "Beginning Balance": 10720.82, "Ending Balance": 10720.82,
            "Loan Payment Instance": 1.0, "Last Payment Instance?": "Y"})])
        self.assertAlmostEqual(rec[NMP], 10720.82, places=2)

    def test_the_lendistry_2024_profile_does_not_count(self):
        # GP 4490649110 as filed: same blanks, but NO posting identity at all.
        rec = _summary([_posting(**{
            "SBA GP Number": "4490649110", "Delinquency Status": "DELINQUENT",
            "Beginning Balance": 232756.11, "Ending Balance": 232756.11,
            "Loan Payment Instance": "", "Last Payment Instance?": ""})])
        self.assertEqual(rec[NMP], 0.0)

    def test_the_two_profiles_are_separated_within_one_population(self):
        rec = _summary([
            _posting(**{"SBA GP Number": "4609059108", "Beginning Balance": 10720.82,
                        "Ending Balance": 10720.82}),
            _posting(**{"SBA GP Number": "4490649110", "Beginning Balance": 232756.11,
                        "Ending Balance": 232756.11,
                        "Loan Payment Instance": "", "Last Payment Instance?": ""}),
        ])
        self.assertAlmostEqual(rec[NMP], 10720.82, places=2)


if __name__ == "__main__":
    unittest.main()
