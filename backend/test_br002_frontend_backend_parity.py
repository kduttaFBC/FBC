"""BR-002 parity: the reviewer's on-screen payment figure vs the authoritative
backend calculation.

`CorrectionReview.jsx` renders a payment-units figure beside every
recommendation, immediately left of the reviewer's decision dropdown:

    Received $390.00 + Fee $20.00 / Sched $500.00 = 0.82 units

That figure is computed IN JAVASCRIPT by `paymentUnits()`, re-deriving the
BR-002 sufficiency numerator that `checks._apply_payment_transition` owns. Two
implementations of one client-confirmed rule, in two languages, with no test
binding them -- until this file. There is no JS test runner in the repo, so the
binding is asserted from Python: the JS source is read and fingerprinted, and
the formula is transcribed here and compared against the backend's real
behaviour.

WHAT THIS FILE IS FOR. It protects the CURRENT behaviour and exposes drift. It
does not assert that the two implementations agree everywhere -- they
deliberately do not, and the divergences are documented below as tests rather
than smoothed over. It changes nothing.

MEASURED ON THE REAL JUNE FIXTURES (8 servicers, 126 recommendations):
  53 positive-payment rows   -> frontend ratio EXACTLY equals the backend's
  20 zero-payment-with-fee   -> divergent in kind, max ratio 0.000758, so every
                                one renders as "0.00 units" at two decimals
   0 negative-payment rows   -> none in the June fixtures; July has them, and
                                they now DO reach the page (see below)
   0 zero/absent-scheduled   -> would render blank

UPDATED 2026-08-12. Reversal rows now reach the review page. The old
eligibility-layer `continue` in corrections.py did not merely block
auto-correction of a reversal-driven delinquency -- it deleted the loan from
Human Review entirely, so no reviewer ever saw the mismatch (confirmed on
Lendistry 2025, July 2026, loan 5951029110). That routing was corrected: the
recommendation is now emitted with review_required=True and
REVIEW_REASON_REVERSAL. The rule that a reversal is never AUTOMATICALLY
correctable is unchanged.

This file's original warning was that such a row would display a wrong figure,
because the frontend ratio adds the servicing fee and the backend's reversal
branch deliberately excludes it. That warning was acted on rather than
inherited: `paymentUnits()` now returns null when `reversalEquivalents` is set,
so no BR-002 ratio is rendered for a reversal row at all, and the real
deficiency arithmetic is shown in its place. Pinned below.

THE THREE THINGS KEEPING THIS SAFE:
  1. reversal rows are suppressed from paymentUnits by an explicit guard --
     this is now a real guard on paymentUnits, not an accident of eligibility;
  2. servicing fees happen to be ~0.05% of scheduled payments in this
     portfolio;
  3. manual discipline keeps the two numerators aligned.
Each is pinned by a test below, so if any stops holding, this fails.

OUT OF SCOPE (explicitly): removing the duplicate frontend calculation and
making the backend the single authoritative source. This file only makes the
current arrangement observable.

Run with:
    python -m pytest test_br002_frontend_backend_parity.py
"""

import math
import os
import re
import unittest

import pandas as pd

import checks
import corrections
import file_readers


JSX_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "..", "frontend", "src", "components", "CorrectionReview.jsx")


def _jsx_source():
    with open(JSX_PATH) as fh:
        return fh.read()


def _payment_units_source():
    src = _jsx_source()
    match = re.search(r"function paymentUnits\(rec\)\s*\{(.*?)\n\}", src, re.S)
    return match.group(1) if match else None


def frontend_payment_units(payment_received, servicing_fee, scheduled_payment,
                           reversal_equivalents=None):
    """Faithful Python transcription of CorrectionReview.jsx paymentUnits().

    Mirrors the JS semantics deliberately, including `if (!sched) return null`
    (0, NaN, None and '' are all falsy in JS) and `|| 0` coalescing. If the JS
    changes, PaymentUnitsSourceFingerprintTests fails and points here.

    reversal_equivalents mirrors the 2026-08-12 guard: the JS returns null
    before computing anything when `rec.reversalEquivalents != null`, so a
    reversal row renders no ratio. Defaults to None (not a reversal), which is
    every pre-existing call site's case and leaves their behaviour unchanged."""
    if reversal_equivalents is not None:
        return None
    try:
        sched = float(scheduled_payment) if scheduled_payment is not None else float("nan")
    except (TypeError, ValueError):
        sched = float("nan")
    if not sched or math.isnan(sched):
        return None
    received = payment_received or 0
    fee = servicing_fee or 0
    return (float(received) + float(fee)) / sched


class PaymentUnitsSourceFingerprintTests(unittest.TestCase):
    """Binds this file to the real JS. If paymentUnits is edited, these fail and
    the transcription above must be re-checked before the parity tests below
    can be trusted."""

    def test_the_function_still_exists(self):
        self.assertIsNotNone(_payment_units_source(),
                             "paymentUnits() not found in CorrectionReview.jsx")

    def test_the_numerator_still_adds_servicing_fee_to_payment_received(self):
        body = _payment_units_source()
        self.assertIn("rec.paymentReceived", body)
        self.assertIn("rec.servicingFee", body)
        self.assertRegex(body, r"paymentReceived[^)]*\)\s*\+\s*Number\(rec\.servicingFee")

    def test_the_denominator_is_still_scheduled_payment(self):
        body = _payment_units_source()
        self.assertIn("rec.scheduledPayment", body)
        self.assertRegex(body, r"return\s+adjusted\s*/\s*sched")

    def test_a_falsy_scheduled_payment_still_short_circuits(self):
        self.assertRegex(_payment_units_source(), r"if\s*\(!sched\)\s*return null")

    def test_a_reversal_row_short_circuits_before_any_arithmetic(self):
        # The guard added 2026-08-12, when reversal rows started reaching the
        # review page. It must come BEFORE the ratio is computed, so no reversal
        # row can ever render a BR-002 figure that did not drive its result.
        body = _payment_units_source()
        self.assertRegex(body, r"if\s*\(rec\.reversalEquivalents\s*!=\s*null\)\s*return null")
        self.assertLess(body.index("reversalEquivalents"), body.index("return adjusted / sched"),
                        "the reversal guard must precede the ratio calculation")

    def test_the_reversal_arithmetic_is_rendered_in_its_place(self):
        # Suppressing the wrong figure is only half of it — the reviewer still
        # needs the numbers the calculated value actually came from.
        src = _jsx_source()
        self.assertIn("function reversalEvidence(rec)", src)
        # The CALL SITE, matched on the callee rather than on the caller's local
        # variable name. It was `reversalEvidence(r)` until the 2026-08-13 Human
        # Review reorganisation moved the row markup into a RecommendationRow
        # component whose prop is `rec`; pinning the identifier made this test fail
        # on a pure rename while the arithmetic it guards was untouched.
        self.assertRegex(src, r"reversalEvidence\(\w+\)")
        self.assertIn("Math.abs(received)", src)   # scheduled + |negative amount|
        self.assertIn("sched * 0.8", src)          # measured in 80%-of-scheduled units

    def test_the_display_is_two_decimal_places(self):
        # The precision at which any divergence becomes visible to a reviewer.
        self.assertIn("units.toFixed(2)", _jsx_source())

    def test_the_figure_is_rendered_next_to_the_decision_control(self):
        src = _jsx_source()
        # Callee-matched, not caller-matched -- see the note on the reversal
        # fingerprint above. The point of this test is that the ratio is COMPUTED
        # for the row and RENDERED beside the decision control, which is asserted
        # positionally below rather than by a variable name.
        self.assertRegex(src, r"const units = paymentUnits\(\w+\)")
        self.assertIn("units.toFixed(2)", src)
        # ...and the figure precedes the decision <select> within the same row.
        self.assertLess(src.index("units.toFixed(2)"),
                        src.index("aria-label={`Decision for loan"),
                        "the BR-002 figure must render before the decision control")


class PositivePaymentParityTests(unittest.TestCase):
    """The only branch where the displayed number is legible AND meaningful.
    Here the two implementations must agree exactly."""

    CASES = [          # received, fee, scheduled
        (500.0, 0.0, 500.0),
        (390.0, 20.0, 500.0),
        (400.0, 0.0, 500.0),
        (1000.0, 25.0, 500.0),
        (212.37, 0.11, 214.66),
        (1594.20, 0.24, 1594.20),
    ]

    def test_the_displayed_ratio_equals_the_backend_sufficiency_ratio(self):
        for received, fee, scheduled in self.CASES:
            with self.subTest(received=received, fee=fee, scheduled=scheduled):
                shown = frontend_payment_units(received, fee, scheduled)
                backend_ratio = (received + fee) / scheduled
                self.assertAlmostEqual(shown, backend_ratio, places=12)

    def test_the_backend_numerator_includes_the_servicing_fee(self):
        # 390/500 = 0.78 -> below the 0.80 bar on its own.
        # (390+20)/500 = 0.82 -> at or above it once the fee is restored.
        # If the fee is ever removed from the backend numerator this flips to a
        # miss and this test fails -- which is the drift this file exists to catch.
        without_fee, _ = checks._apply_payment_transition(3, 390.0, 500.0, servicing_fee=None)
        with_fee, _ = checks._apply_payment_transition(3, 390.0, 500.0, servicing_fee=20.0)
        self.assertEqual(without_fee, 4)      # miss: ratio 0.78
        self.assertEqual(with_fee, 3)         # held:  ratio 0.82
        self.assertAlmostEqual(frontend_payment_units(390.0, 20.0, 500.0), 0.82, places=10)

    def test_the_threshold_the_displayed_ratio_is_judged_against(self):
        self.assertEqual(checks.PAYMENT_SUFFICIENCY_THRESHOLD, 0.80)

    def test_the_screen_shows_a_ratio_not_an_outcome(self):
        # Three visibly different ratios produce the SAME outcome, and one
        # visibly small step produces a different one. The reviewer cannot read
        # the outcome off the number:
        #
        #   0.78 -> one more miss        0.80 -> held        (a 0.02 step, opposite results)
        #   0.80 / 0.95 / 1.00 -> all held steady            (net reduction is
        #                                                     installments_covered - 1)
        #   2.00 -> one rank recovered
        #
        # Documented, not a defect: the display is the rule's INPUT, not its output.
        for received, expected in ((390.0, 4), (400.0, 3), (475.0, 3), (500.0, 3), (1000.0, 2)):
            with self.subTest(received=received):
                new_missed, _ = checks._apply_payment_transition(3, received, 500.0)
                self.assertEqual(new_missed, expected)
        self.assertAlmostEqual(frontend_payment_units(475.0, 0.0, 500.0), 0.95, places=10)
        self.assertAlmostEqual(frontend_payment_units(500.0, 0.0, 500.0), 1.00, places=10)

    def test_one_covered_installment_only_offsets_the_current_cycle(self):
        # Pins the "net rank reduction = installments_covered - 1" rule that the
        # ratio on screen gives no hint of.
        for received, expected_delta in ((500.0, 0), (1000.0, -1), (1500.0, -2)):
            with self.subTest(received=received):
                new_missed, _ = checks._apply_payment_transition(3, received, 500.0)
                self.assertEqual(new_missed - 3, expected_delta)


class DocumentedDivergenceZeroPaymentTests(unittest.TestCase):
    """INTENTIONAL BACKEND BEHAVIOUR, NOT MIRRORED ON SCREEN.

    A zero Payment Received is a full miss, full stop -- the backend returns
    prior + 1 before it ever looks at the fee or the scheduled amount. The
    frontend has no zero branch and computes fee / scheduled, which is > 0
    whenever a fee is present.

    Documented here rather than changed in either implementation."""

    def test_the_backend_ignores_the_fee_entirely_when_nothing_was_received(self):
        for fee in (None, 0.0, 0.24, 50.0, 5000.0):
            with self.subTest(fee=fee):
                new_missed, _ = checks._apply_payment_transition(3, 0.0, 500.0, servicing_fee=fee)
                self.assertEqual(new_missed, 4, "a zero payment is always one more miss")

    def test_the_frontend_shows_a_non_zero_ratio_for_the_same_input(self):
        self.assertGreater(frontend_payment_units(0.0, 50.0, 500.0), 0.0)
        self.assertAlmostEqual(frontend_payment_units(0.0, 50.0, 500.0), 0.10, places=10)

    def test_on_the_real_portfolio_the_divergence_is_below_display_precision(self):
        # Every real zero-payment fee is <= 0.0758% of the scheduled payment, so
        # toFixed(2) renders 0.00. If a servicer ever reports a materially larger
        # fee this stops being invisible -- see the real-fixture test below.
        largest_observed = 0.78 / 1029.66
        self.assertLess(largest_observed, 0.005,
                        "0.005 is where toFixed(2) starts rendering 0.01")
        self.assertEqual(f"{largest_observed:.2f}", "0.00")


class DocumentedDivergenceNegativePaymentTests(unittest.TestCase):
    """INTENTIONAL BACKEND BEHAVIOUR, NOT MIRRORED ON SCREEN.

    A negative Payment Received uses the client-confirmed deficiency formula and
    DELIBERATELY excludes the servicing fee -- checks.py records why: 'a small
    negative payment plus a fee could otherwise sum to a positive figure and
    silently bypass the reversal path altogether'. The frontend would add it.

    Such rows DO now reach the review page (routing corrected 2026-08-12), so
    the divergence is no longer prevented by their absence — it is prevented by
    paymentUnits' explicit reversal guard, which renders no ratio for them at
    all. Both halves are pinned."""

    def test_the_backend_excludes_the_fee_from_the_reversal_branch(self):
        no_fee, equiv_no_fee = checks._apply_payment_transition(2, -202.15, 212.37,
                                                                servicing_fee=None)
        big_fee, equiv_big_fee = checks._apply_payment_transition(2, -202.15, 212.37,
                                                                  servicing_fee=1000.0)
        self.assertEqual((no_fee, equiv_no_fee), (big_fee, equiv_big_fee),
                         "the fee must not influence the reversal branch at all")

    def test_a_fee_can_never_flip_a_reversal_into_a_positive_payment(self):
        new_missed, equivalents = checks._apply_payment_transition(2, -10.0, 500.0,
                                                                   servicing_fee=5000.0)
        self.assertIsNotNone(equivalents, "still routed through the reversal path")
        self.assertGreater(new_missed, 2)

    def test_the_frontend_would_have_shown_a_positive_figure(self):
        # Exactly the outcome the paymentUnits guard exists to prevent: a
        # reversal that the ratio, fee included, renders as a POSITIVE figure.
        # This is the unguarded arithmetic, which is why the guard is required.
        self.assertGreater(frontend_payment_units(-10.0, 5000.0, 500.0), 0.0)

    def test_the_guard_means_a_reversal_row_renders_no_figure_at_all(self):
        # Same inputs as above, now carrying the reversal flag the backend sends.
        self.assertIsNone(frontend_payment_units(-10.0, 5000.0, 500.0,
                                                 reversal_equivalents=1))

    def test_reversal_rows_reach_review_flagged_rather_than_being_deleted(self):
        # SUPERSEDES test_reversal_rows_are_excluded_from_candidacy (2026-08-12).
        # The old contract asserted a reversal loan must never become a
        # candidate. That was over-broad: it removed the mismatch from Human
        # Review entirely instead of merely blocking auto-correction. The
        # never-automatic half of the rule is asserted separately below.
        eligible = corrections.build_eligible_corrections([{
            "gp": "1", "reasons": ["DELINQUENCY STATUS MISMATCH"], "dq_match": False,
            "calculated_dq": "60 DAYS", "srf_dq": "CURRENT", "srf_status": "ACTIVE",
            "reversal_equivalents": 2,          # a reversal
            "excel_rows": [10], "sheet_name": "Servicer Inputs", "header_excel_row": 9,
        }])
        self.assertEqual(len(eligible), 1, "the reviewer must be able to see this mismatch")
        self.assertEqual(eligible[0]["field"], "Delinquency Status")
        self.assertTrue(eligible[0]["review_required"])
        self.assertEqual(eligible[0]["review_reason"], corrections.REVIEW_REASON_REVERSAL)
        self.assertEqual(eligible[0]["reversal_equivalents"], 2,
                         "carried through so the review page can suppress the BR-002 ratio")


class MissingOrZeroScheduledPaymentTests(unittest.TestCase):

    def test_the_frontend_renders_nothing(self):
        for scheduled in (0, 0.0, None, "", float("nan")):
            with self.subTest(scheduled=scheduled):
                self.assertIsNone(frontend_payment_units(500.0, 10.0, scheduled))

    def test_the_backend_holds_the_rank_steady(self):
        for scheduled in (0, None, float("nan")):
            with self.subTest(scheduled=scheduled):
                new_missed, _ = checks._apply_payment_transition(3, 500.0, scheduled)
                self.assertEqual(new_missed, 3, "no ratio is computable, so no movement")

    def test_a_zero_payment_with_zero_scheduled_is_still_a_miss(self):
        # The zero branch fires BEFORE the scheduled check -- the real shape of
        # several June Community Bank rows (received 0, scheduled 0).
        new_missed, _ = checks._apply_payment_transition(3, 0.0, 0.0)
        self.assertEqual(new_missed, 4)
        self.assertIsNone(frontend_payment_units(0.0, 0.0, 0.0))


MEMO_PATH = os.environ.get("FALCONBRIDGE_MEMO_XLSX")
SOUP_PATH = os.environ.get("FALCONBRIDGE_SOUP_XLSX")
SRF_DIR = os.environ.get("FALCONBRIDGE_SRF_DIR")
REAL = all([MEMO_PATH, SOUP_PATH, SRF_DIR])

SERVICERS = [("Community Bank and Trust", "Community Bank & Trust", "Community Bank and Trust"),
             ("Lendistry_SOUP 2024", "Lendistry", "Lendistry 2024"),
             ("Lendistry_SOUP 2025", "Lendistry", "Lendistry 2025"),
             ("Millennium_SOUP 2024", "Millennium Bank", "Millennium 2024"),
             ("Millennium_SOUP 2025", "Millennium Bank", "Millennium 2025"),
             ("Mission Valley Bank", "Mission Valley Bank", "Mission Valley Bank"),
             ("Peoples Bank", "Peoples Bank", "Peoples Bank"),
             ("Poppy Bank", "Poppy Bank", "Poppy Bank")]


@unittest.skipUnless(REAL, "set FALCONBRIDGE_MEMO_XLSX / _SOUP_XLSX / _SRF_DIR")
class RealPortfolioParityTests(unittest.TestCase):
    """Every recommendation a reviewer would actually see, June 2026."""

    @classmethod
    def setUpClass(cls):
        import glob
        soup = file_readers.read_soup_file(open(SOUP_PATH, "rb").read())
        memo = file_readers.resolve_memo_attribution(
            file_readers.read_service_memo(open(MEMO_PATH, "rb").read()),
            file_readers.build_memo_loan_index(soup))
        cls.recs = []
        for pat, soup_key, label in SERVICERS:
            path = [x for x in glob.glob(os.path.join(SRF_DIR, "*.xlsx"))
                    if pat.lower() in os.path.basename(x).lower()][0]
            sk, dl, lt, si, deal = file_readers.read_servicer_file(
                open(path, "rb").read(), filename=f"{soup_key}||{label}__x.xlsx")
            results = checks.run_servicer_reconciliation(
                dl, lt, si, file_readers.filter_soup_by_servicer(soup, sk),
                memo_df=memo, memo_lender_key=sk, deal=deal)
            stat01 = next(r for r in results if r["id"].split("::")[0] == "STAT-01")
            found, _, _ = corrections.build_recommendations(stat01, label)
            cls.recs.extend(found)

    def _split(self):
        positive, zero_fee, negative = [], [], []
        for r in self.recs:
            received = float(r.get("paymentReceived") or 0)
            fee = float(r.get("servicingFee") or 0)
            if received < 0:
                negative.append(r)
            elif received == 0 and fee:
                zero_fee.append(r)
            elif received > 0:
                positive.append(r)
        return positive, zero_fee, negative

    def test_every_reversal_row_that_reaches_the_page_is_flagged_and_suppressed(self):
        # SUPERSEDES test_no_reversal_row_reaches_the_review_page (2026-08-12).
        # Reversal rows are no longer kept off the page; what must hold now is
        # that each one is flagged for review and renders NO BR-002 ratio.
        _positive, _zero_fee, negative = self._split()
        for r in negative:
            with self.subTest(gp=r["gp"]):
                self.assertIsNotNone(r["reversalEquivalents"],
                                     "a negative-payment row must carry the reversal flag")
                self.assertIn(corrections.REVIEW_REASON_REVERSAL, r["reviewFlag"] or "")
                self.assertIsNone(
                    frontend_payment_units(float(r["paymentReceived"]),
                                           float(r.get("servicingFee") or 0),
                                           float(r["scheduledPayment"]),
                                           reversal_equivalents=r["reversalEquivalents"]),
                    "a reversal row must display no payment-units figure")

    def test_only_reversal_rows_carry_the_reversal_flag(self):
        # The converse: the guard must not swallow ordinary rows' figures.
        positive, zero_fee, _n = self._split()
        for r in positive + zero_fee:
            with self.subTest(gp=r["gp"]):
                self.assertIsNone(r["reversalEquivalents"])

    def test_every_positive_row_displays_the_backend_ratio_exactly(self):
        positive, _z, _n = self._split()
        self.assertGreater(len(positive), 0)
        for r in positive:
            received = float(r["paymentReceived"])
            fee = float(r.get("servicingFee") or 0)
            scheduled = float(r["scheduledPayment"])
            with self.subTest(gp=r["gp"]):
                self.assertAlmostEqual(
                    frontend_payment_units(received, fee, scheduled),
                    (received + fee) / scheduled, places=12)

    def test_every_zero_payment_divergence_is_invisible_at_two_decimals(self):
        _p, zero_fee, _n = self._split()
        for r in zero_fee:
            shown = frontend_payment_units(0.0, float(r["servicingFee"]),
                                           float(r["scheduledPayment"]))
            with self.subTest(gp=r["gp"], fee=r["servicingFee"]):
                self.assertGreater(shown, 0.0, "the frontend does compute a non-zero ratio")
                self.assertEqual(f"{shown:.2f}", "0.00",
                                 "a servicer reporting a materially larger fee would make "
                                 "this divergence visible to the reviewer")

    def test_the_portfolio_shape_this_file_was_measured_against(self):
        positive, zero_fee, _negative = self._split()
        # The negative-row count is deliberately NOT asserted here anymore: it
        # was 0 in the June fixtures only because reversal rows were being
        # dropped, and it is a real, varying portfolio figure now. The shape
        # that still matters is that positive and zero-fee rows both exist.
        self.assertGreater(len(positive), 0)
        self.assertGreater(len(zero_fee), 0)
        largest = max(float(r["servicingFee"]) / float(r["scheduledPayment"])
                      for r in zero_fee)
        self.assertLess(largest, 0.005)


if __name__ == "__main__":
    unittest.main()
