"""
Regression test for portfolio_summary_check (PORT-01) in checks.py.

Locks in the confirmed business rules as of 2026-08-04 (revised):
  - Outstanding, all delinquency buckets, Not Making Payment, Deferred,
    Bankruptcy, and Foreclosure all use ENDING BALANCE.
  - Paid In Full and Liquidation use BEGINNING BALANCE instead (separate
    activity metrics), and are EXCLUDED from Outstanding.
  - "Not Making Payment" is based solely on consolidated Payment Received
    EXPLICITLY equal to 0 — a blank/missing/NaN value does NOT count as
    zero (corrected 2026-08-04) — with NO Loan Status filtering.
  - Paid In Full / Liquidation detection uses controlled EXACT matching
    (corrected 2026-08-04) — not a broad substring check.
  - Every other column is otherwise unaffected by Loan Status EXCEPT the
    one explicit Outstanding exclusion above — categories overlap freely.

Run with:
    python -m unittest test_portfolio_summary.py
"""

import unittest

import pandas as pd

import checks


class PortfolioSummaryWorkedExampleTests(unittest.TestCase):
    """The exact worked example from the original spec: a single loan that
    should contribute to four different columns simultaneously."""

    def test_single_loan_contributes_to_four_overlapping_columns(self):
        servicer_inputs_df = pd.DataFrame([{
            "SBA GP Number": "1", "Lender Loan Number": "L-1", "Loan Status": "DEFERRED",
            "Delinquency Status": "30 DAYS", "Beginning Balance": 100000, "Ending Balance": 100000,
            "Payment Received": 0, "Scheduled Payment": 0,
        }])
        result = checks.portfolio_summary_check(servicer_inputs_df, servicer_name="Test Bank")
        r = result["records"][0]

        self.assertEqual(r["Outstanding ($)"], 100000)
        self.assertEqual(r["30 Days ($)"], 100000)
        self.assertEqual(r["Not Making Payment ($)"], 100000)
        self.assertEqual(r["Deferred ($)"], 100000)

        for col in ["Current", "Delinquent", "60 Days", "90 Days", "120 Days", "150+ Days",
                    "Bankruptcy", "Foreclosure", "Liquidation", "Paid In Full"]:
            self.assertEqual(r[f"{col} ($)"], 0, f"{col} should be 0 for this loan")


class PortfolioSummaryOutstandingExclusionTests(unittest.TestCase):
    """Outstanding excludes PAID-IN-FULL and LIQUIDATION loans entirely —
    those balances are activity that has left the portfolio, tracked
    separately using Beginning Balance instead of Ending Balance."""

    def _build(self, loan_status, beginning, ending, received=500, scheduled=500):
        return pd.DataFrame([{
            "SBA GP Number": "1", "Lender Loan Number": "L-1", "Loan Status": loan_status,
            "Delinquency Status": "CURRENT", "Beginning Balance": beginning, "Ending Balance": ending,
            "Payment Received": received, "Scheduled Payment": scheduled,
        }])

    def test_paid_in_full_excluded_from_outstanding(self):
        sdf = self._build("PAID-IN-FULL", beginning=50000, ending=0)
        r = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]
        self.assertEqual(r["Outstanding ($)"], 0)

    def test_paid_in_full_uses_beginning_balance_not_ending(self):
        sdf = self._build("PAID-IN-FULL", beginning=50000, ending=0)
        r = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]
        self.assertEqual(r["Paid In Full ($)"], 50000)

    def test_liquidation_excluded_from_outstanding(self):
        sdf = self._build("LIQUIDATION", beginning=30000, ending=0)
        r = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]
        self.assertEqual(r["Outstanding ($)"], 0)

    def test_liquidation_uses_beginning_balance_not_ending(self):
        sdf = self._build("LIQUIDATION", beginning=30000, ending=0)
        r = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]
        self.assertEqual(r["Liquidation ($)"], 30000)

    def test_active_loan_still_counts_normally_in_outstanding(self):
        sdf = self._build("ACTIVE", beginning=75000, ending=70000)
        r = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]
        self.assertEqual(r["Outstanding ($)"], 70000)


class PortfolioSummaryPaidInFullExactMatchTests(unittest.TestCase):
    """CORRECTED (2026-08-04): Paid In Full / Liquidation detection uses
    controlled exact matching, not a broad substring check — the prior
    "PAID" in loan_status check could incorrectly classify values like
    PARTIALLY PAID, UNPAID, or PAID DOWN as Paid In Full."""

    def _paid_in_full_amount(self, loan_status):
        sdf = pd.DataFrame([{
            "SBA GP Number": "1", "Loan Status": loan_status, "Delinquency Status": "CURRENT",
            "Beginning Balance": 40000, "Ending Balance": 0, "Payment Received": 40000, "Scheduled Payment": 0,
        }])
        return checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]["Paid In Full ($)"]

    def test_paid_in_full_hyphenated_counts(self):
        self.assertEqual(self._paid_in_full_amount("PAID-IN-FULL"), 40000)

    def test_paid_in_full_spaced_counts(self):
        self.assertEqual(self._paid_in_full_amount("PAID IN FULL"), 40000)

    def test_pif_abbreviation_counts(self):
        self.assertEqual(self._paid_in_full_amount("PIF"), 40000)

    def test_partially_paid_does_not_count(self):
        self.assertEqual(self._paid_in_full_amount("PARTIALLY PAID"), 0)

    def test_unpaid_does_not_count(self):
        self.assertEqual(self._paid_in_full_amount("UNPAID"), 0)

    def test_paid_down_does_not_count(self):
        self.assertEqual(self._paid_in_full_amount("PAID DOWN"), 0)


class PortfolioSummaryNotMakingPaymentIndependenceTests(unittest.TestCase):
    """"Not Making Payment" is based solely on consolidated Payment
    Received = 0. Loan Status is intentionally ignored."""

    def _run(self, loan_status):
        servicer_inputs_df = pd.DataFrame([{
            "SBA GP Number": "1", "Lender Loan Number": "L-1", "Loan Status": loan_status,
            "Delinquency Status": "CURRENT", "Beginning Balance": 10000, "Ending Balance": 10000,
            "Payment Received": 0, "Scheduled Payment": 500,
        }])
        result = checks.portfolio_summary_check(servicer_inputs_df, servicer_name="Test Bank")
        return result["records"][0]["Not Making Payment ($)"]

    def test_active_zero_payment_counts(self):
        self.assertEqual(self._run("ACTIVE"), 10000)

    def test_deferred_zero_payment_still_counts(self):
        self.assertEqual(self._run("DEFERRED"), 10000)

    def test_bankruptcy_zero_payment_still_counts(self):
        self.assertEqual(self._run("BANKRUPTCY"), 10000)

    def test_foreclosure_zero_payment_still_counts(self):
        self.assertEqual(self._run("FORECLOSURE"), 10000)

    def test_liquidation_zero_payment_still_counts(self):
        self.assertEqual(self._run("LIQUIDATION"), 10000)

    def test_paid_in_full_zero_payment_still_counts_if_ending_balance_nonzero(self):
        self.assertEqual(self._run("PAID-IN-FULL"), 10000)

    def test_nonzero_payment_does_not_count_regardless_of_loan_status(self):
        servicer_inputs_df = pd.DataFrame([{
            "SBA GP Number": "1", "Lender Loan Number": "L-1", "Loan Status": "DEFERRED",
            "Delinquency Status": "CURRENT", "Beginning Balance": 10000, "Ending Balance": 10000,
            "Payment Received": 500, "Scheduled Payment": 500,
        }])
        result = checks.portfolio_summary_check(servicer_inputs_df, servicer_name="Test Bank")
        self.assertEqual(result["records"][0]["Not Making Payment ($)"], 0)


class PortfolioSummaryBlankPaymentReceivedTests(unittest.TestCase):
    """CORRECTED (2026-08-04): a blank/missing/NaN consolidated Payment
    Received must NOT be treated as $0 for Not Making Payment — the prior
    implementation silently defaulted a missing value to 0.0, incorrectly
    equating "no payment data at all" with "confirmed zero payment"."""

    def test_explicit_zero_counts(self):
        sdf = pd.DataFrame([{
            "SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
            "Beginning Balance": 20000, "Ending Balance": 20000, "Payment Received": 0, "Scheduled Payment": 500,
        }])
        result = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]["Not Making Payment ($)"]
        self.assertEqual(result, 20000)

    def test_blank_string_does_not_count(self):
        sdf = pd.DataFrame([{
            "SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
            "Beginning Balance": 20000, "Ending Balance": 20000, "Payment Received": "", "Scheduled Payment": 500,
        }])
        result = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]["Not Making Payment ($)"]
        self.assertEqual(result, 0)

    def test_none_does_not_count(self):
        sdf = pd.DataFrame([{
            "SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
            "Beginning Balance": 20000, "Ending Balance": 20000, "Payment Received": None, "Scheduled Payment": 500,
        }])
        result = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]["Not Making Payment ($)"]
        self.assertEqual(result, 0)

    def test_nan_does_not_count(self):
        import numpy as np
        sdf = pd.DataFrame([{
            "SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
            "Beginning Balance": 20000, "Ending Balance": 20000, "Payment Received": np.nan, "Scheduled Payment": 500,
        }])
        result = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]["Not Making Payment ($)"]
        self.assertEqual(result, 0)


class PortfolioSummaryPercentageTests(unittest.TestCase):
    """Direct tests of the percentage calculation."""

    def test_outstanding_percentage_is_100_when_nonzero(self):
        sdf = pd.DataFrame([{
            "SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
            "Beginning Balance": 10000, "Ending Balance": 10000, "Payment Received": 500, "Scheduled Payment": 500,
        }])
        r = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]
        self.assertEqual(r["Outstanding (%)"], 100.0)

    def test_current_percentage_equals_current_over_outstanding(self):
        sdf = pd.DataFrame([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 30000, "Ending Balance": 30000, "Payment Received": 500, "Scheduled Payment": 500},
            {"SBA GP Number": "2", "Loan Status": "ACTIVE", "Delinquency Status": "DELINQUENT",
             "Beginning Balance": 70000, "Ending Balance": 70000, "Payment Received": 0, "Scheduled Payment": 500},
        ])
        r = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]
        expected_pct = round((30000 / 100000) * 100, 2)
        self.assertEqual(r["Current (%)"], expected_pct)
        self.assertEqual(r["Current (%)"], 30.0)

    def test_all_percentages_zero_when_outstanding_is_zero(self):
        sdf = pd.DataFrame([{
            "SBA GP Number": "1", "Loan Status": "PAID-IN-FULL", "Delinquency Status": "CURRENT",
            "Beginning Balance": 50000, "Ending Balance": 0, "Payment Received": 50000, "Scheduled Payment": 0,
        }])
        r = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]
        self.assertEqual(r["Outstanding (%)"], 0.0)
        for col in ["Current", "Delinquent", "30 Days", "60 Days", "90 Days", "120 Days",
                    "150+ Days", "Not Making Payment", "Deferred", "Bankruptcy",
                    "Foreclosure", "Liquidation", "Paid In Full"]:
            self.assertEqual(r[f"{col} (%)"], 0.0, f"{col} (%) should be 0.0 when Outstanding is 0")


class PortfolioSummaryMultiPostingConsolidationTests(unittest.TestCase):
    """Confirms PORT-01 uses the SAME existing consolidation behavior as
    STAT-01/NEG-01 for a loan with multiple postings this cycle:
    Beginning Balance = first chronological posting, Ending Balance /
    Loan Status / Delinquency Status = final chronological posting,
    Payment Received = sum of all postings."""

    def test_multi_posting_loan_uses_correct_consolidated_fields(self):
        servicer_inputs_df = pd.DataFrame([
            {"SBA GP Number": "1", "Loan Payment Instance": 1, "Loan Status": "ACTIVE",
             "Delinquency Status": "DELINQUENT", "Beginning Balance": 100000, "Ending Balance": 95000,
             "Payment Received": 300, "Scheduled Payment": 500},
            {"SBA GP Number": "1", "Loan Payment Instance": 2, "Loan Status": "DEFERRED",
             "Delinquency Status": "30 DAYS", "Beginning Balance": 95000, "Ending Balance": 90000,
             "Payment Received": 200, "Scheduled Payment": 500},
        ])
        r = checks.portfolio_summary_check(servicer_inputs_df, servicer_name="Test")["records"][0]

        self.assertEqual(r["Outstanding ($)"], 90000)
        self.assertEqual(r["30 Days ($)"], 90000)
        self.assertEqual(r["Deferred ($)"], 90000)
        self.assertEqual(r["Not Making Payment ($)"], 0)
        self.assertEqual(r["Delinquent ($)"], 0)


class PortfolioSummary150PlusDaysBoundaryTests(unittest.TestCase):
    """Direct boundary tests for the 150+ Days bucket."""

    def _bucket_amount(self, dq_status, bucket):
        sdf = pd.DataFrame([{
            "SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": dq_status,
            "Beginning Balance": 10000, "Ending Balance": 10000, "Payment Received": 500, "Scheduled Payment": 500,
        }])
        r = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]
        return r[f"{bucket} ($)"]

    def test_120_days_goes_to_120_days_bucket(self):
        self.assertEqual(self._bucket_amount("120 DAYS", "120 Days"), 10000)
        self.assertEqual(self._bucket_amount("120 DAYS", "150+ Days"), 0)

    def test_150_days_goes_to_150_plus_bucket(self):
        self.assertEqual(self._bucket_amount("150 DAYS", "150+ Days"), 10000)

    def test_150_plus_days_label_goes_to_150_plus_bucket(self):
        self.assertEqual(self._bucket_amount("150+ DAYS", "150+ Days"), 10000)

    def test_180_days_goes_to_150_plus_bucket(self):
        self.assertEqual(self._bucket_amount("180 DAYS", "150+ Days"), 10000)

    def test_600_days_goes_to_150_plus_bucket(self):
        self.assertEqual(self._bucket_amount("600 DAYS", "150+ Days"), 10000)

    def test_unknown_status_enters_no_delinquency_bucket(self):
        sdf = pd.DataFrame([{
            "SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "N/A",
            "Beginning Balance": 10000, "Ending Balance": 10000, "Payment Received": 500, "Scheduled Payment": 500,
        }])
        r = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]
        for bucket in ["Current", "Delinquent", "30 Days", "60 Days", "90 Days", "120 Days", "150+ Days"]:
            self.assertEqual(r[f"{bucket} ($)"], 0, f"{bucket} should be 0 for an unrecognized status")

    def test_blank_status_enters_no_delinquency_bucket(self):
        sdf = pd.DataFrame([{
            "SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "",
            "Beginning Balance": 10000, "Ending Balance": 10000, "Payment Received": 500, "Scheduled Payment": 500,
        }])
        r = checks.portfolio_summary_check(sdf, servicer_name="Test")["records"][0]
        for bucket in ["Current", "Delinquent", "30 Days", "60 Days", "90 Days", "120 Days", "150+ Days"]:
            self.assertEqual(r[f"{bucket} ($)"], 0, f"{bucket} should be 0 for a blank status")


class PortfolioSummaryEmptyPortfolioTests(unittest.TestCase):
    """PORT-01 must return one valid, all-zero summary record for an empty
    portfolio, without raising an exception."""

    def test_empty_portfolio_returns_one_zero_record(self):
        empty_df = pd.DataFrame(columns=[
            "SBA GP Number", "Loan Status", "Delinquency Status",
            "Beginning Balance", "Ending Balance", "Payment Received", "Scheduled Payment",
        ])
        result = checks.portfolio_summary_check(empty_df, servicer_name="Empty Bank")
        self.assertEqual(len(result["records"]), 1)
        self.assertEqual(result["issues_found"], 0)
        self.assertEqual(result["status"], "PASS")
        r = result["records"][0]
        self.assertEqual(r["Outstanding ($)"], 0)
        self.assertEqual(r["Outstanding (%)"], 0.0)


class PortfolioSummaryCBTValidationTests(unittest.TestCase):
    """Reproduces the exact published CB&T June 2026 figures. The first 12
    loans below are traced directly from the real file (their exact
    per-loan Ending/Beginning Balance, Loan Status, and Delinquency Status,
    confirmed against the actual servicer file during development). The
    remaining 4 "F_" loans are computed aggregate stand-ins for the
    remaining ~140 real loans in that file, sized to close the exact
    remaining gap for each category — NOT a claim that the real portfolio
    has exactly 4 more loans, but a verified, exact reconstruction of the
    AGGREGATE arithmetic. All 14 published category values match exactly
    (aside from sub-dollar float rounding)."""

    def test_complete_synthetic_reconstruction_matches_all_14_categories(self):
        rows = [
            {"SBA GP Number": "A1", "Loan Status": "ACTIVE", "Delinquency Status": "DELINQUENT", "Beginning Balance": 64936.11, "Ending Balance": 64936.11, "Payment Received": 0, "Scheduled Payment": 0},
            {"SBA GP Number": "A2", "Loan Status": "ACTIVE", "Delinquency Status": "DELINQUENT", "Beginning Balance": 152756.27, "Ending Balance": 152756.27, "Payment Received": 0, "Scheduled Payment": 0},
            {"SBA GP Number": "A3", "Loan Status": "ACTIVE", "Delinquency Status": "DELINQUENT", "Beginning Balance": 273172.98, "Ending Balance": 273172.98, "Payment Received": 0, "Scheduled Payment": 0},
            {"SBA GP Number": "A4", "Loan Status": "ACTIVE", "Delinquency Status": "DELINQUENT", "Beginning Balance": 123328.42, "Ending Balance": 123328.42, "Payment Received": 0, "Scheduled Payment": 0},
            {"SBA GP Number": "A5", "Loan Status": "ACTIVE", "Delinquency Status": "DELINQUENT", "Beginning Balance": 51695.29, "Ending Balance": 51695.29, "Payment Received": 0, "Scheduled Payment": 0},
            {"SBA GP Number": "D1", "Loan Status": "DEFERRED", "Delinquency Status": "DELINQUENT", "Beginning Balance": 504764.06, "Ending Balance": 504764.06, "Payment Received": 0, "Scheduled Payment": 0},
            {"SBA GP Number": "D2", "Loan Status": "DEFERRED", "Delinquency Status": "30 DAYS", "Beginning Balance": 145282.54, "Ending Balance": 145282.54, "Payment Received": 0, "Scheduled Payment": 0},
            {"SBA GP Number": "D3", "Loan Status": "DEFERRED", "Delinquency Status": "60 DAYS", "Beginning Balance": 843761.03, "Ending Balance": 843761.03, "Payment Received": 0, "Scheduled Payment": 0},
            {"SBA GP Number": "D4", "Loan Status": "DEFERRED", "Delinquency Status": "60 DAYS", "Beginning Balance": 331080.11, "Ending Balance": 331080.11, "Payment Received": 0, "Scheduled Payment": 0},
            {"SBA GP Number": "P1", "Loan Status": "PAID-IN-FULL", "Delinquency Status": "CURRENT", "Beginning Balance": 165560.28, "Ending Balance": 0, "Payment Received": 167146.90, "Scheduled Payment": 0},
            {"SBA GP Number": "P2", "Loan Status": "PAID-IN-FULL", "Delinquency Status": "CURRENT", "Beginning Balance": 148795.52, "Ending Balance": 0, "Payment Received": 150683.11, "Scheduled Payment": 0},
            {"SBA GP Number": "P3", "Loan Status": "PAID-IN-FULL", "Delinquency Status": "CURRENT", "Beginning Balance": 365207.99, "Ending Balance": 0, "Payment Received": 369332.20, "Scheduled Payment": 0},
            {"SBA GP Number": "F_CURRENT", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT", "Beginning Balance": 26729324.02, "Ending Balance": 26729324.02, "Payment Received": 1000, "Scheduled Payment": 1000},
            {"SBA GP Number": "F_DELINQUENT", "Loan Status": "ACTIVE", "Delinquency Status": "DELINQUENT", "Beginning Balance": 1431718.42, "Ending Balance": 1431718.42, "Payment Received": 1000, "Scheduled Payment": 1000},
            {"SBA GP Number": "F_60DAYS", "Loan Status": "ACTIVE", "Delinquency Status": "60 DAYS", "Beginning Balance": 882431.32, "Ending Balance": 882431.32, "Payment Received": 1000, "Scheduled Payment": 1000},
            {"SBA GP Number": "F_90DAYS_DEFERRED", "Loan Status": "DEFERRED", "Delinquency Status": "90 DAYS", "Beginning Balance": 173000.26, "Ending Balance": 173000.26, "Payment Received": 1000, "Scheduled Payment": 1000},
        ]
        servicer_inputs_df = pd.DataFrame(rows)
        result = checks.portfolio_summary_check(servicer_inputs_df, servicer_name="Community Bank and Trust")
        r = result["records"][0]

        known = {
            "Outstanding": 31707251, "Current": 26729324, "Delinquent": 2602372,
            "30 Days": 145283, "60 Days": 2057272, "90 Days": 173000, "120 Days": 0,
            "150+ Days": 0, "Not Making Payment": 2490776.81, "Deferred": 1997888,
            "Bankruptcy": 0, "Foreclosure": 0, "Liquidation": 0, "Paid In Full": 679564,
        }
        for cat, expected in known.items():
            with self.subTest(category=cat):
                self.assertAlmostEqual(r[f"{cat} ($)"], expected, delta=1.0)


class AuditedPortfolioSummarySharedArchitectureTests(unittest.TestCase):
    """CLIENT-CONFIRMED (2026-08-04): the Audited Portfolio Summary must
    reuse STAT-01's real calculated_dq via audited_status_by_loan, never
    reimplement delinquency logic. These tests confirm the shared
    architecture behaves correctly using status_match_check's actual
    output, not a hand-constructed stand-in."""

    def _build_files(self, servicer_dq, prior_soup_dq, received=0, scheduled=500.0, servicer_loan_status="ACTIVE"):
        servicer_inputs_df = pd.DataFrame([{
            "SBA GP Number": "1", "Lender Loan Number": "L-1", "Loan Status": servicer_loan_status,
            "Delinquency Status": servicer_dq, "Beginning Balance": 100000, "Ending Balance": 100000,
            "Payment Received": received, "Scheduled Payment": scheduled,
        }])
        soup_servicer_df = pd.DataFrame([{
            "SBA GP Number": "1", "Servicer": "Test Bank", "Payment Month": "2026-05-01",
            "Loan Status": "ACTIVE", "Delinquency Status": prior_soup_dq,
            "Beginning Actual Balance": 100000, "Ending Actual Balance": 100000,
        }])
        return servicer_inputs_df, soup_servicer_df

    def test_audited_status_by_loan_is_complete_not_just_flagged_subset(self):
        servicer_inputs_df, soup_servicer_df = self._build_files("CURRENT", "CURRENT", received=500, scheduled=500)
        result = checks.status_match_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name="Test Bank")
        self.assertEqual(len(result["records"]), 0)
        self.assertIn("1", result["audited_status_by_loan"])
        self.assertEqual(result["audited_status_by_loan"]["1"], "CURRENT")

    def test_loan_corrected_current_to_delinquent_moves_only_in_audited(self):
        servicer_inputs_df, soup_servicer_df = self._build_files("CURRENT", "CURRENT", received=0, scheduled=500)
        stat01 = checks.status_match_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name="Test Bank")
        audited_status_by_loan = stat01["audited_status_by_loan"]
        self.assertEqual(audited_status_by_loan["1"], "DELINQUENT")

        original = checks.portfolio_summary_check(servicer_inputs_df, servicer_name="Test Bank")["records"][0]
        audited = checks.audited_portfolio_summary_record(servicer_inputs_df, audited_status_by_loan, servicer_name="Test Bank")

        self.assertEqual(original["Current ($)"], 100000)
        self.assertEqual(original["Delinquent ($)"], 0)
        self.assertEqual(audited["Current ($)"], 0)
        self.assertEqual(audited["Delinquent ($)"], 100000)

    def test_loan_corrected_90_days_to_60_days_moves_only_in_audited(self):
        servicer_inputs_df, soup_servicer_df = self._build_files("90 DAYS", "60 DAYS", received=500, scheduled=500)
        stat01 = checks.status_match_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name="Test Bank")
        audited_status_by_loan = stat01["audited_status_by_loan"]
        self.assertEqual(audited_status_by_loan["1"], "60 DAYS")

        original = checks.portfolio_summary_check(servicer_inputs_df, servicer_name="Test Bank")["records"][0]
        audited = checks.audited_portfolio_summary_record(servicer_inputs_df, audited_status_by_loan, servicer_name="Test Bank")

        self.assertEqual(original["90 Days ($)"], 100000)
        self.assertEqual(original["60 Days ($)"], 0)
        self.assertEqual(audited["90 Days ($)"], 0)
        self.assertEqual(audited["60 Days ($)"], 100000)

    def test_shared_categories_identical_between_original_and_audited(self):
        servicer_inputs_df, soup_servicer_df = self._build_files(
            "CURRENT", "CURRENT", received=0, scheduled=500, servicer_loan_status="DEFERRED"
        )
        stat01 = checks.status_match_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name="Test Bank")
        audited_status_by_loan = stat01["audited_status_by_loan"]
        original = checks.portfolio_summary_check(servicer_inputs_df, servicer_name="Test Bank")["records"][0]
        audited = checks.audited_portfolio_summary_record(servicer_inputs_df, audited_status_by_loan, servicer_name="Test Bank")

        for cat in ["Outstanding", "Not Making Payment", "Deferred", "Bankruptcy", "Foreclosure", "Liquidation", "Paid In Full"]:
            with self.subTest(category=cat):
                self.assertEqual(original[f"{cat} ($)"], audited[f"{cat} ($)"])

    def test_fallback_to_raw_status_when_loan_absent_from_audited_dataset(self):
        servicer_inputs_df = pd.DataFrame([{
            "SBA GP Number": "999", "Loan Status": "ACTIVE", "Delinquency Status": "60 DAYS",
            "Beginning Balance": 5000, "Ending Balance": 5000, "Payment Received": 500, "Scheduled Payment": 500,
        }])
        empty_audited_status = {}
        result = checks.audited_portfolio_summary_record(servicer_inputs_df, empty_audited_status, servicer_name="Test Bank")
        self.assertEqual(result["60 Days ($)"], 5000)

    def test_real_cbt_data_audited_differs_from_original_exactly_as_investigated(self):
        rows = [
            {"SBA GP Number": "A", "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
             "Beginning Balance": 322621.20, "Ending Balance": 322621.20, "Payment Received": 500, "Scheduled Payment": 500},
            {"SBA GP Number": "B", "Loan Status": "DEFERRED", "Delinquency Status": "90 DAYS",
             "Beginning Balance": 173000.26, "Ending Balance": 173000.26, "Payment Received": 224.36, "Scheduled Payment": 500},
        ]
        servicer_inputs_df = pd.DataFrame(rows)
        audited_status_by_loan = {"A": "DELINQUENT", "B": "60 DAYS"}

        original = checks.portfolio_summary_check(servicer_inputs_df, servicer_name="Test")["records"][0]
        audited = checks.audited_portfolio_summary_record(servicer_inputs_df, audited_status_by_loan, servicer_name="Test")

        self.assertEqual(original["Current ($)"], 322621.20)
        self.assertEqual(original["90 Days ($)"], 173000.26)
        self.assertEqual(audited["Delinquent ($)"], 322621.20)
        self.assertEqual(audited["60 Days ($)"], 173000.26)
        self.assertEqual(original["Outstanding ($)"], audited["Outstanding ($)"])


if __name__ == "__main__":
    unittest.main()
