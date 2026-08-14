"""The Portfolio Summary percentage DENOMINATOR — CLIENT-CONFIRMED (2026-08-13).

THE RULE
--------
Every Portfolio Summary percentage divides by OUTSTANDING ENDING BALANCE.
Paid In Full and Liquidation are no exception.

Paid In Full and Liquidation are CURRENT PERIOD TERMINAL ACTIVITY measures whose
dollar numerators use BEGINNING Balance, while the denominator is the ENDING
Balance of the loans that REMAIN -- a set that explicitly excludes the PIF and
Liquidation loans themselves. Numerator and denominator therefore come from
different balance points AND from disjoint loan sets.

That mixed basis is INTENTIONAL and client-confirmed. This file exists so it
cannot drift back to a "tidier" single basis by accident, in either direction:
denominator moved to Beginning Balance, or numerator moved to Ending Balance.

WHY THESE FIGURES
-----------------
The invariance tests are the load-bearing ones, and they are built so they
cannot pass by coincidence. Each changes the Beginning Balance of the
OUTSTANDING loan only -- which leaves the real numerator and the real
denominator untouched, while moving what a Beginning-Balance denominator would
produce:

    frame            PIF $     Outstanding $ (ending)   correct %   Beginning-basis %
    baseline       100,000            500,000             20.00           10.00
    beginning x2   100,000            500,000             20.00            5.00

So a regression to a Beginning-Balance denominator does not merely change the
number, it changes it DIFFERENTLY between the two frames -- there is no constant
either implementation could return that satisfies both. Liquidation is pinned
the same way (50.00 correct, against 21.74 then 11.63).

SCOPE
-----
Documentation and regression protection for an EXISTING confirmed rule. No
calculation is asserted to be anything other than what production already does.
The June 2026 SOUP 2025 baseline (1,436 / 1,428 / 8 / $1,465,536.12 / 0 / $0.00)
is deliberately NOT re-pinned here -- test_final_summary_terminal_activity.py
already locks it, and one owner per fact is the point.

    FALCONBRIDGE_ACCEPTED_JUNE_DIR=<dir with the three accepted vF workbooks> \\
    FALCONBRIDGE_SOUP_XLSX=... python3 -m pytest test_portfolio_summary_percentage_basis.py
"""

import glob
import os
import unittest

import pandas as pd

import checks
import file_readers


ACCEPTED_DIR = os.environ.get("FALCONBRIDGE_ACCEPTED_JUNE_DIR") \
    or os.environ.get("FALCONBRIDGE_ACCEPTED_SRF_DIR")
SOUP_XLSX = os.environ.get("FALCONBRIDGE_SOUP_XLSX")
REAL = bool(ACCEPTED_DIR and SOUP_XLSX
            and os.path.isdir(ACCEPTED_DIR) and os.path.exists(SOUP_XLSX))

# The June 2026 accepted SOUP 2025 percentages. Dollars are owned by
# test_final_summary_terminal_activity.py; only the RATIO is asserted here.
JUNE_PCT = [
    ("Community Bank and Trust", "Community Bank & Trust", "community bank and trust", 2.14),
    ("Lendistry 2025", "Lendistry", "lendistry_soup 2025-1", 1.63),
    ("Millennium 2025", "Millennium Bank", "millennium_soup 2025-1", 0.21),
]


def _row(**over):
    row = {
        "SBA GP Number": "1000000001",
        "Loan Status": "ACTIVE",
        "Delinquency Status": "CURRENT",
        "Payment Received": 250.0,
        "Principal": 100.0,
        "Scheduled Payment": 250.0,
        "Beginning Balance": 10000.0,
        "Ending Balance": 9900.0,
        "Loan Payment Instance": 1.0,
        "Last Payment Instance?": "Y",
        "_excel_row": 10,
    }
    row.update(over)
    return row


def _views(rows, overrides=None, excluded=None):
    return checks.verified_portfolio_summary_views(
        pd.DataFrame(rows), overrides or {}, "S", excluded_gps=excluded)


def _outstanding_loan(beginning, ending=500000.0, gp="1000000001"):
    return _row(**{"SBA GP Number": gp, "Loan Status": "ACTIVE",
                   "Beginning Balance": beginning, "Ending Balance": ending})


def _pif_loan(beginning=100000.0, gp="2000000002"):
    return _row(**{"SBA GP Number": gp, "Loan Status": "PAID-IN-FULL",
                   "Beginning Balance": beginning, "Ending Balance": 0.0,
                   "Payment Received": beginning})


def _liquidation_loan(beginning=250000.0, gp="3000000003"):
    return _row(**{"SBA GP Number": gp, "Loan Status": "LIQUIDATION",
                   "Beginning Balance": beginning, "Ending Balance": beginning})


class PaidInFullPercentageBasisTests(unittest.TestCase):
    """PIF (%) = PIF Beginning Balance / Outstanding ENDING Balance."""

    def test_the_denominator_is_outstanding_ending_balance(self):
        d = _views([_outstanding_loan(900000.0), _pif_loan()])["dollars"]
        self.assertEqual(d["Outstanding ($)"], 500000.00)
        self.assertEqual(d["Paid In Full ($)"], 100000.00)
        self.assertEqual(d["Paid In Full (%)"], 20.00)

    def test_it_is_not_the_total_beginning_balance(self):
        """The same frame under a Beginning-Balance denominator reads 10.00."""
        d = _views([_outstanding_loan(900000.0), _pif_loan()])["dollars"]
        total_beginning = 900000.0 + 100000.0
        self.assertNotEqual(
            d["Paid In Full (%)"],
            round(d["Paid In Full ($)"] / total_beginning * 100, 2))

    def test_changing_beginning_balance_does_not_move_the_percentage(self):
        """THE invariance test. Only the OUTSTANDING loan's Beginning Balance
        changes, so neither the real numerator nor the real denominator moves --
        but a Beginning-Balance denominator would fall from 10.00 to 5.00."""
        base = _views([_outstanding_loan(900000.0), _pif_loan()])["dollars"]
        moved = _views([_outstanding_loan(1900000.0), _pif_loan()])["dollars"]

        self.assertEqual(base["Outstanding ($)"], moved["Outstanding ($)"])
        self.assertEqual(base["Paid In Full ($)"], moved["Paid In Full ($)"])
        self.assertEqual(base["Paid In Full (%)"], moved["Paid In Full (%)"])
        self.assertEqual(moved["Paid In Full (%)"], 20.00)

        # No single Beginning-Balance implementation could satisfy both frames.
        self.assertAlmostEqual(100000.0 / 1000000.0 * 100, 10.00, places=2)
        self.assertAlmostEqual(100000.0 / 2000000.0 * 100, 5.00, places=2)

    def test_the_numerator_is_beginning_balance_not_ending(self):
        """The other half of the mixed basis. A PIF loan ends at $0.00, so an
        Ending-Balance numerator would report $0.00 and 0.00%."""
        d = _views([_outstanding_loan(900000.0), _pif_loan()])["dollars"]
        self.assertEqual(d["Paid In Full ($)"], 100000.00)
        self.assertNotEqual(d["Paid In Full ($)"], 0.0)
        self.assertGreater(d["Paid In Full (%)"], 0.0)


class LiquidationPercentageBasisTests(unittest.TestCase):
    """Liquidation (%) = Liquidation Beginning Balance / Outstanding ENDING."""

    def test_the_denominator_is_outstanding_ending_balance(self):
        d = _views([_outstanding_loan(900000.0), _liquidation_loan()])["dollars"]
        self.assertEqual(d["Outstanding ($)"], 500000.00)
        self.assertEqual(d["Liquidation ($)"], 250000.00)
        self.assertEqual(d["Liquidation (%)"], 50.00)

    def test_it_is_not_the_total_beginning_balance(self):
        d = _views([_outstanding_loan(900000.0), _liquidation_loan()])["dollars"]
        total_beginning = 900000.0 + 250000.0
        self.assertNotEqual(
            d["Liquidation (%)"],
            round(d["Liquidation ($)"] / total_beginning * 100, 2))

    def test_changing_beginning_balance_does_not_move_the_percentage(self):
        base = _views([_outstanding_loan(900000.0), _liquidation_loan()])["dollars"]
        moved = _views([_outstanding_loan(1900000.0), _liquidation_loan()])["dollars"]

        self.assertEqual(base["Outstanding ($)"], moved["Outstanding ($)"])
        self.assertEqual(base["Liquidation ($)"], moved["Liquidation ($)"])
        self.assertEqual(base["Liquidation (%)"], moved["Liquidation (%)"])
        self.assertEqual(moved["Liquidation (%)"], 50.00)

        self.assertAlmostEqual(250000.0 / 1150000.0 * 100, 21.74, places=2)
        self.assertAlmostEqual(250000.0 / 2150000.0 * 100, 11.63, places=2)

    def test_a_liquidation_loan_keeps_its_balance_out_of_the_denominator(self):
        """A LIQUIDATION loan can carry a POSITIVE ending balance and still be
        excluded from Outstanding -- so its own balance must not appear in the
        denominator it is measured against."""
        with_liq = _views([_outstanding_loan(900000.0), _liquidation_loan()])["dollars"]
        alone = _views([_outstanding_loan(900000.0)])["dollars"]
        self.assertEqual(with_liq["Outstanding ($)"], alone["Outstanding ($)"])


class DenominatorMembershipTests(unittest.TestCase):

    def test_adding_a_paid_in_full_loan_does_not_change_outstanding(self):
        alone = _views([_outstanding_loan(900000.0)])["dollars"]
        with_pif = _views([_outstanding_loan(900000.0), _pif_loan()])["dollars"]
        self.assertEqual(alone["Outstanding ($)"], with_pif["Outstanding ($)"])
        self.assertEqual(with_pif["Outstanding ($)"], 500000.00)

    def test_outstanding_reports_one_hundred_percent_of_itself(self):
        d = _views([_outstanding_loan(900000.0), _pif_loan()])["dollars"]
        self.assertEqual(d["Outstanding (%)"], 100.0)

    def test_zero_outstanding_yields_zero_percent_not_an_error(self):
        """A population of nothing but terminal activity has an empty
        denominator. The guard must return 0.0 rather than raise or imply
        100%."""
        d = _views([_pif_loan()])["dollars"]
        self.assertEqual(d["Outstanding ($)"], 0.0)
        self.assertEqual(d["Outstanding (%)"], 0.0)
        self.assertEqual(d["Paid In Full ($)"], 100000.00)
        self.assertEqual(d["Paid In Full (%)"], 0.0)


class EveryCategorySharesTheDenominatorTests(unittest.TestCase):
    """Not just PIF and Liquidation -- the rule is universal, so a future change
    cannot special-case one bucket onto a different basis."""

    def _mixed(self):
        return [
            _outstanding_loan(900000.0, gp="1000000001"),
            _row(**{"SBA GP Number": "1000000004", "Loan Status": "DEFERRED",
                    "Delinquency Status": "30 DAYS", "Beginning Balance": 60000.0,
                    "Ending Balance": 60000.0, "Payment Received": 0.0}),
            _row(**{"SBA GP Number": "1000000005", "Loan Status": "BANKRUPTCY",
                    "Delinquency Status": "90 DAYS", "Beginning Balance": 40000.0,
                    "Ending Balance": 40000.0, "Payment Received": 0.0}),
            _pif_loan(), _liquidation_loan(),
        ]

    def test_every_category_percentage_is_its_dollars_over_outstanding(self):
        d = _views(self._mixed())["dollars"]
        outstanding = d["Outstanding ($)"]
        self.assertGreater(outstanding, 0.0)
        for category in checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER:
            with self.subTest(category=category):
                self.assertEqual(
                    d[f"{category} (%)"],
                    round(d[f"{category} ($)"] / outstanding * 100, 2))


class AllThreeSummaryLayersShareTheBasisTests(unittest.TestCase):
    """Original (PORT-01), Audited and Final Audited all delegate to
    _build_portfolio_summary_record, so the denominator must be identical in
    each. If one layer ever grows its own pct(), this fails."""

    def _frame(self):
        return pd.DataFrame([_outstanding_loan(900000.0), _pif_loan()])

    def test_original_uses_outstanding_ending_balance(self):
        d = checks.original_portfolio_summary_views(self._frame(), "S")["dollars"]
        self.assertEqual(d["Outstanding ($)"], 500000.00)
        self.assertEqual(d["Paid In Full (%)"], 20.00)

    def test_audited_uses_outstanding_ending_balance(self):
        d = checks.audited_portfolio_summary_record(self._frame(), {}, servicer_name="S")
        self.assertEqual(d["Outstanding ($)"], 500000.00)
        self.assertEqual(d["Paid In Full (%)"], 20.00)

    def test_final_audited_uses_outstanding_ending_balance(self):
        d = checks.verified_portfolio_summary_views(self._frame(), {}, "S")["dollars"]
        self.assertEqual(d["Outstanding ($)"], 500000.00)
        self.assertEqual(d["Paid In Full (%)"], 20.00)

    def test_the_three_layers_agree_on_the_denominator(self):
        frame = self._frame()
        a = checks.original_portfolio_summary_views(frame, "S")["dollars"]["Outstanding ($)"]
        b = checks.audited_portfolio_summary_record(frame, {}, servicer_name="S")["Outstanding ($)"]
        c = checks.verified_portfolio_summary_views(frame, {}, "S")["dollars"]["Outstanding ($)"]
        self.assertEqual(a, b)
        self.assertEqual(b, c)


@unittest.skipUnless(REAL, "set FALCONBRIDGE_ACCEPTED_JUNE_DIR (or _ACCEPTED_SRF_DIR) and _SOUP_XLSX")
class JuneAcceptedPercentagesTests(unittest.TestCase):
    """The rule on the real accepted June SOUP 2025 files. Dollars are owned by
    test_final_summary_terminal_activity.py; this asserts only that the reported
    percentage IS the terminal-activity dollars over Outstanding."""

    @staticmethod
    def _views_for(substring, soup_key, display):
        soup = file_readers.read_soup_file(open(SOUP_XLSX, "rb").read())
        path = next(p for p in sorted(glob.glob(os.path.join(ACCEPTED_DIR, "*.xlsx")))
                    if substring in os.path.basename(p).lower())
        _sk, _dl, _lt, si, _deal = file_readers.read_servicer_file(
            open(path, "rb").read(), filename=f"{soup_key}||{display}__x.xlsx")
        del soup
        return checks.verified_portfolio_summary_views(si, {}, servicer_name=display)

    def test_pif_percentage_is_pif_dollars_over_outstanding(self):
        for display, soup_key, substring, expected in JUNE_PCT:
            with self.subTest(servicer=display):
                d = self._views_for(substring, soup_key, display)["dollars"]
                self.assertEqual(d["Paid In Full (%)"], expected)
                self.assertEqual(
                    d["Paid In Full (%)"],
                    round(d["Paid In Full ($)"] / d["Outstanding ($)"] * 100, 2))

    def test_liquidation_is_zero_and_reports_zero_percent(self):
        for display, soup_key, substring, _expected in JUNE_PCT:
            with self.subTest(servicer=display):
                d = self._views_for(substring, soup_key, display)["dollars"]
                self.assertEqual(d["Liquidation ($)"], 0.00)
                self.assertEqual(d["Liquidation (%)"], 0.00)

    def test_the_denominator_excludes_the_paid_in_full_balances(self):
        """Outstanding excludes the PIF loans, so including them would give a
        strictly larger denominator and a strictly smaller ratio.

        Asserted UNROUNDED. On the real June figures PIF is a fraction of a
        percent of the portfolio, so both bases round to the same 2dp value for
        Millennium (0.21 either way) -- real data cannot discriminate here. The
        synthetic invariance tests above are the ones that catch a basis
        regression; this one confirms the exclusion holds on production data."""
        for display, soup_key, substring, expected in JUNE_PCT:
            with self.subTest(servicer=display):
                d = self._views_for(substring, soup_key, display)["dollars"]
                self.assertEqual(d["Paid In Full (%)"], expected)
                reported = d["Paid In Full ($)"] / d["Outstanding ($)"] * 100
                including_pif = (d["Paid In Full ($)"]
                                 / (d["Outstanding ($)"] + d["Paid In Full ($)"]) * 100)
                self.assertLess(including_pif, reported)


if __name__ == "__main__":
    unittest.main()
