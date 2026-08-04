"""
Regression test for the negative-payment (reversal) delinquency formula in
_apply_payment_transition, checks.py.

CLIENT-CONFIRMED (2026-08-04, second revision):

    total_payment_deficiency = scheduled_payment + abs(negative_amount)
    payment_equivalents_removed = max(1, floor(
        total_payment_deficiency / (scheduled_payment * PAYMENT_SUFFICIENCY_THRESHOLD)
    ))

This file exists SPECIFICALLY so that a future change to this formula (or to
PAYMENT_SUFFICIENCY_THRESHOLD, or to the deficiency construction) is caught
by a failing test rather than silently shipping — every expected value below
is a literal, hardcoded number confirmed against the client's own worked
examples, not re-derived from the function under test. If a legitimate,
client-confirmed formula change happens later, THESE EXPECTED VALUES must be
updated deliberately, with a new explicit confirmation — a failing test here
is a prompt to go get that confirmation, not something to silence by making
the test compute its own expectation.

Run with:
    python -m unittest test_reversal_thresholds.py
or via pytest, which auto-discovers unittest.TestCase classes.
"""

import unittest

import checks


class ReversalThresholdTableTests(unittest.TestCase):
    """Locks in the exact $100-scheduled-payment threshold table confirmed
    on 2026-08-04: transitions at $60 (1->2 levels), $140 (2->3), $220
    (3->4), and $300 (4->5) -- verified directly against the client's own
    worked example ($100 current deficiency + $60 removed = $160;
    $160 / $80 threshold unit = 2)."""

    SCHEDULED = 100.0

    def _equivalents(self, negative_amount):
        _, equivalents = checks._apply_payment_transition(0, -negative_amount, self.SCHEDULED)
        return equivalents

    def test_just_below_first_transition_is_one_level(self):
        self.assertEqual(self._equivalents(59.99), 1)

    def test_first_transition_at_60_is_two_levels(self):
        # Client's own worked example: $100 + $60 = $160; $160 / $80 = 2.0000
        self.assertEqual(self._equivalents(60.00), 2)

    def test_just_above_first_transition_is_two_levels(self):
        self.assertEqual(self._equivalents(60.01), 2)

    def test_just_below_second_transition_is_two_levels(self):
        self.assertEqual(self._equivalents(139.99), 2)

    def test_second_transition_at_140_is_three_levels(self):
        # $100 + $140 = $240; $240 / $80 = 3.0000
        self.assertEqual(self._equivalents(140.00), 3)

    def test_just_above_second_transition_is_three_levels(self):
        self.assertEqual(self._equivalents(140.01), 3)

    def test_just_below_third_transition_is_three_levels(self):
        self.assertEqual(self._equivalents(219.99), 3)

    def test_third_transition_at_220_is_four_levels(self):
        # $100 + $220 = $320; $320 / $80 = 4.0000
        self.assertEqual(self._equivalents(220.00), 4)

    def test_just_above_third_transition_is_four_levels(self):
        self.assertEqual(self._equivalents(220.01), 4)

    def test_just_below_fourth_transition_is_four_levels(self):
        self.assertEqual(self._equivalents(299.99), 4)

    def test_fourth_transition_at_300_is_five_levels(self):
        # $100 + $300 = $400; $400 / $80 = 5.0000
        self.assertEqual(self._equivalents(300.00), 5)

    def test_minimum_one_level_for_a_tiny_negative_amount(self):
        # Even $0.01 removed must never round down to zero levels -- ANY
        # negative net payment means no valid payment was received.
        self.assertEqual(self._equivalents(0.01), 1)

    def test_full_transition_table_at_once(self):
        """The complete table, as one assertion, so the whole shape is
        visible in a single failure diff if anything shifts. Starts at
        $0.01, not $0.00 -- a literal zero isn't a genuine reversal case
        (Payment Received == 0 is the separate "no payment" branch, not
        this formula's domain)."""
        table = [
            (0.01, 1), (20.00, 1), (40.00, 1), (59.99, 1),
            (60.00, 2), (60.01, 2), (80.00, 2), (100.00, 2), (120.00, 2), (139.99, 2),
            (140.00, 3), (140.01, 3), (160.00, 3), (180.00, 3), (200.00, 3), (219.99, 3),
            (220.00, 4), (220.01, 4), (240.00, 4), (260.00, 4), (280.00, 4), (299.99, 4),
            (300.00, 5),
        ]
        actual = [(amt, self._equivalents(amt)) for amt, _expected in table]
        expected = table
        self.assertEqual(actual, expected)


class ReversalWorkedExampleTests(unittest.TestCase):
    """Locks in specific real-loan worked examples confirmed during this
    formula's review, as end-to-end anchors independent of the $100 table
    above (different Scheduled Payment amounts, real prior-status baselines)."""

    def test_loan_5521259109(self):
        # $212.37 scheduled, -$202.15 received, prior SOUP status CURRENT.
        # Client-confirmed: deficiency $414.52, threshold unit $169.90,
        # ratio 2.44, floor 2 -> CURRENT -> DELINQUENT -> 30 DAYS.
        new_missed, equivalents = checks._apply_payment_transition(0, -202.15, 212.37)
        self.assertEqual(equivalents, 2)
        self.assertEqual(checks._missed_count_to_status(new_missed), "30 DAYS")

    def test_a_211_73_scheduled_payment_with_301_38_removed(self):
        # $211.73 scheduled, -$301.38 received, prior SOUP status CURRENT.
        # Deficiency $513.11, threshold unit $169.384, ratio 3.0293, floor 3
        # -> 60 DAYS. NOTE: this example was validated as "2 levels / 30
        # DAYS" under the now-superseded 100%-divisor formula (2026-08-04,
        # first revision) -- the value below is the CURRENT, correct result
        # under the 80%-divisor formula, recomputed and confirmed directly
        # against the function, not carried over from the earlier version.
        new_missed, equivalents = checks._apply_payment_transition(0, -301.38, 211.73)
        self.assertEqual(equivalents, 3)
        self.assertEqual(checks._missed_count_to_status(new_missed), "60 DAYS")


class ReversalSafeguardTests(unittest.TestCase):
    """Locks in the two explicit safeguards confirmed alongside this
    formula: an invalid Scheduled Payment must surface as UNVERIFIED
    (None, None), never a guessed equivalents count."""

    def test_negative_scheduled_payment_returns_none(self):
        result = checks._apply_payment_transition(0, -202.15, -212.37)
        self.assertEqual(result, (None, None))

    def test_zero_scheduled_payment_returns_none(self):
        result = checks._apply_payment_transition(0, -202.15, 0)
        self.assertEqual(result, (None, None))

    def test_none_scheduled_payment_returns_none(self):
        result = checks._apply_payment_transition(0, -202.15, None)
        self.assertEqual(result, (None, None))


if __name__ == "__main__":
    unittest.main()
