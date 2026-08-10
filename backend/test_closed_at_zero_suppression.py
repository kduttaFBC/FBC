"""Closed-at-zero suppression (client-confirmed 2026-08-10).

When the client's OWN prior SOUP history establishes that a loan reached
PAID-IN-FULL or LIQUIDATION with an Ending Actual Balance of 0, the loan is
closed. It must not appear anywhere in the current client-facing exception
population -- BAL-01, BAL-02, STAT-01, STAT-02 or recommendations -- whatever
Loan Status or balance the servicer subsequently reports.

This is the other half of _VALID_EXIT_TRANSITION_KEYWORDS, whose own comment
recorded this direction as "a different, unvalidated scenario" pending
confirmation.

It is an ELIGIBILITY gate, not a calculation: every value is still computed,
and audited_status_by_loan is still captured for the loan. The Audited
Portfolio Summary excludes it through an explicit exclusion set instead --
omitting the audited status would make the summary fall back to the
servicer's own reported value, which is the opposite of the intent.

Run with:
    python -m pytest test_closed_at_zero_suppression.py
"""

import unittest

import pandas as pd

import checks
import corrections


def _soup(rows):
    # Rows are (gp, month, status, dq, balance) with an optional 6th Deal
    # element; omitting it puts every row in one partition, which is what a
    # SOUP export without a Deal column reduces to.
    return pd.DataFrame([{
        "SBA GP Number": row[0], "Servicer": "T", "Payment Month": row[1],
        "Loan Status": row[2], "Delinquency Status": row[3],
        "Beginning Actual Balance": row[4], "Ending Actual Balance": row[4],
        "Deal": row[5] if len(row) > 5 else "",
    } for row in rows])


def _srf(rows):
    return pd.DataFrame([{
        "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}", "Loan Status": status,
        "Delinquency Status": dq, "Beginning Balance": bal, "Ending Balance": bal,
        "Payment Received": 0.0, "Scheduled Payment": 500.0, "Servicing Fee": 0.0,
    } for gp, status, dq, bal in rows])


class ClosedAtZeroDetectionTests(unittest.TestCase):
    """loans_closed_at_zero: what counts as closed, and what does not."""

    def test_paid_in_full_at_zero_is_closed(self):
        out = checks.loans_closed_at_zero(_soup([("1", "2026-05-01", "PAID-IN-FULL", "CURRENT", 0.0)]))
        self.assertIn("1", out)
        self.assertEqual(out["1"]["closed_status"], "PAID-IN-FULL")
        self.assertEqual(out["1"]["closed_balance"], 0.0)

    def test_liquidation_at_zero_is_closed(self):
        self.assertIn("1", checks.loans_closed_at_zero(
            _soup([("1", "2026-05-01", "LIQUIDATION", "CURRENT", 0.0)])))

    def test_closed_status_with_a_nonzero_balance_is_not_closed(self):
        # Zero balance is REQUIRED -- a status claim alone is not enough.
        self.assertEqual(checks.loans_closed_at_zero(
            _soup([("1", "2026-05-01", "PAID-IN-FULL", "CURRENT", 5000.0)])), {})

    def test_zero_balance_without_a_closed_status_is_not_closed(self):
        # A closed status is REQUIRED -- a zero balance alone is not enough.
        self.assertEqual(checks.loans_closed_at_zero(
            _soup([("1", "2026-05-01", "ACTIVE", "CURRENT", 0.0)])), {})

    def test_zero_is_judged_with_the_shared_balance_tolerance(self):
        self.assertIn("1", checks.loans_closed_at_zero(
            _soup([("1", "2026-05-01", "PAID-IN-FULL", "CURRENT", 0.50)])))
        self.assertEqual(checks.loans_closed_at_zero(
            _soup([("1", "2026-05-01", "PAID-IN-FULL", "CURRENT", 25.00)])), {})

    def test_whole_history_is_scanned_not_just_the_latest_row(self):
        # The check is not "is the FINAL row closed" -- a closure followed
        # only by further closed/zero reporting still qualifies. This is why
        # soup_latest is not enough on its own.
        out = checks.loans_closed_at_zero(_soup([
            ("1", "2026-03-01", "PAID-IN-FULL", "CURRENT", 0.0),
            ("1", "2026-04-01", "PAID-IN-FULL", "CURRENT", 0.0),
        ]))
        self.assertIn("1", out)
        self.assertEqual(out["1"]["closed_month"], "2026-04-01")

    def test_later_active_amortisation_contradicts_the_closure(self):
        # Qualification 1. A loan that keeps paying down after a closed row
        # is not closed -- the client's own later data refutes it. This is
        # the shape of the three Lendistry 2024 loans when their two deals
        # are not distinguished.
        self.assertEqual(checks.loans_closed_at_zero(_soup([
            ("1", "2026-03-01", "PAID-IN-FULL", "CURRENT", 0.0),
            ("1", "2026-04-01", "ACTIVE", "CURRENT", 12000.0),
            ("1", "2026-05-01", "ACTIVE", "30 DAYS", 11000.0),
        ])), {})

    def test_a_deal_exit_does_not_suppress_a_loan_active_in_another_deal(self):
        # Qualification 2, the real Lendistry 2024 shape: the PAID-IN-FULL
        # $0.00 row sits under one deal in the SAME month the loan continues
        # under another. Closing one deal is not closing the loan.
        self.assertEqual(checks.loans_closed_at_zero(_soup([
            ("1", "2026-03-01", "ACTIVE", "CURRENT", 13000.0, "SOUP2024_1"),
            ("1", "2026-03-01", "PAID-IN-FULL", "CURRENT", 0.0, "SOUP2025_1"),
            ("1", "2026-04-01", "ACTIVE", "CURRENT", 12000.0, "SOUP2024_1"),
        ])), {})

    def test_every_deal_closed_still_qualifies(self):
        out = checks.loans_closed_at_zero(_soup([
            ("1", "2026-03-01", "PAID-IN-FULL", "CURRENT", 0.0, "SOUP2024_1"),
            ("1", "2026-04-01", "LIQUIDATION", "CURRENT", 0.0, "SOUP2025_1"),
        ]))
        self.assertIn("1", out)
        # The reported closure is the latest one across the GP's deals.
        self.assertEqual(out["1"]["closed_month"], "2026-04-01")
        self.assertEqual(out["1"]["closed_deal"], "SOUP2025_1")

    def test_single_deal_closed_as_its_final_activity_qualifies(self):
        # GP 4338609104's shape: one deal, PAID-IN-FULL $0.00 last, nothing
        # after it. Unaffected by either qualification.
        out = checks.loans_closed_at_zero(_soup([
            ("1", "2026-04-01", "ACTIVE", "30 DAYS", 99666.56, "SOUP2025_1"),
            ("1", "2026-05-01", "PAID-IN-FULL", "60 DAYS", 0.0, "SOUP2025_1"),
        ]))
        self.assertIn("1", out)
        self.assertEqual(out["1"]["closed_deal"], "SOUP2025_1")

    def test_contradiction_is_scoped_to_the_same_deal(self):
        # An unrelated deal being active does not itself refute a closure --
        # it disqualifies the GP through qualification 2 (not every deal
        # closed), which is a different reason. Here the OTHER deal also
        # closes, so the GP qualifies despite mid-history activity in it.
        out = checks.loans_closed_at_zero(_soup([
            ("1", "2026-03-01", "PAID-IN-FULL", "CURRENT", 0.0, "A"),
            ("1", "2026-04-01", "ACTIVE", "CURRENT", 5000.0, "B"),
            ("1", "2026-05-01", "LIQUIDATION", "CURRENT", 0.0, "B"),
        ]))
        self.assertIn("1", out)

    def test_missing_columns_degrade_to_suppressing_nothing(self):
        self.assertEqual(checks.loans_closed_at_zero(pd.DataFrame([{"SBA GP Number": "1"}])), {})
        self.assertEqual(checks.loans_closed_at_zero(None), {})


class SuppressionAcrossAllExceptionChecksTests(unittest.TestCase):
    """A closed loan disappears from every client-facing exception surface."""

    def setUp(self):
        # GP 1 closed at zero in SOUP, then re-reported as LIQUIDATION with a
        # stale balance -- the GP 4338609104 shape.
        # GP 2 is an ordinary loan that must be unaffected.
        self.soup = _soup([
            ("1", "2026-04-01", "ACTIVE", "30 DAYS", 99666.56),
            ("1", "2026-05-01", "PAID-IN-FULL", "60 DAYS", 0.0),
            ("2", "2026-05-01", "ACTIVE", "CURRENT", 50000.0),
        ])
        self.srf = _srf([("1", "LIQUIDATION", "120 DAYS", 99666.56),
                         ("2", "ACTIVE", "CURRENT", 50000.0)])
        self.results = checks.run_servicer_reconciliation(
            "T", pd.DataFrame(), self.srf, self.soup, memo_df=None, memo_lender_key="T")

    def _check(self, cid):
        return next(r for r in self.results if r["id"].split("::")[0] == cid)

    def test_absent_from_every_exception_check(self):
        for cid in ("BAL-01", "BAL-02", "STAT-01", "STAT-02"):
            with self.subTest(check=cid):
                gps = [str(x.get("SBA GP Number", "")).strip() for x in self._check(cid)["records"]]
                self.assertNotIn("1", gps)

    def test_produces_no_correction_data_and_no_recommendations(self):
        stat01 = self._check("STAT-01")
        self.assertNotIn("1", [c["gp"] for c in stat01.get("correction_data", [])])
        recs, _, _ = corrections.build_recommendations(stat01, "T")
        self.assertNotIn("1", [r.get("gp") for r in recs])

    def test_unrelated_loan_is_untouched(self):
        stat01 = self._check("STAT-01")
        self.assertIn("2", [c["gp"] for c in stat01.get("correction_data", [])])

    def test_audited_status_is_still_captured_for_the_suppressed_loan(self):
        # Suppression is an eligibility decision, not a change to the audit's
        # reasoning -- the value is still computed and recorded.
        self.assertIn("1", self._check("STAT-01").get("audited_status_by_loan", {}))

    def test_trace_carries_the_full_evidence(self):
        trace = self._check("STAT-01").get("suppressed_closed_loans")
        self.assertEqual(len(trace), 1)
        t = trace[0]
        self.assertEqual(t["SBA GP Number"], "1")
        self.assertEqual(t["Prior Closed Status"], "PAID-IN-FULL")
        self.assertEqual(t["Prior Closed Month"], "2026-05-01")
        self.assertEqual(t["Prior Ending Actual Balance"], 0.0)
        self.assertEqual(t["Current Servicer Loan Status"], "LIQUIDATION")
        self.assertEqual(t["Current Servicer Ending Balance"], 99666.56)
        self.assertEqual(t["Checks Suppressed"], ["BAL-01", "BAL-02", "STAT-01", "STAT-02"])


class AuditedPortfolioExclusionTests(unittest.TestCase):
    """The suppressed loan contributes neither count nor balance to the
    Audited Portfolio Summary -- and the exclusion is EXPLICIT, not achieved
    by withholding the audited status."""

    def setUp(self):
        self.srf = _srf([("1", "LIQUIDATION", "120 DAYS", 99666.56),
                         ("2", "ACTIVE", "CURRENT", 50000.0)])
        self.audited = {"1": "90 DAYS", "2": "CURRENT"}

    def test_excluded_loan_contributes_nothing(self):
        with_loan = checks.audited_portfolio_summary_record(self.srf, self.audited, servicer_name="T")
        without = checks.audited_portfolio_summary_record(
            self.srf, self.audited, servicer_name="T", excluded_gps=["1"])
        self.assertEqual(with_loan["90 Days ($)"] - without["90 Days ($)"], 99666.56)
        self.assertEqual(without["Liquidation ($)"], 0.0)

    def test_omitting_the_audited_status_is_NOT_equivalent_to_excluding(self):
        # The distinction the client called out: a loan with no audited status
        # falls back to the servicer's own reported DQ (120 DAYS), so omission
        # would MISBUCKET it rather than remove it.
        omitted = checks.audited_portfolio_summary_record(
            self.srf, {"2": "CURRENT"}, servicer_name="T")
        excluded = checks.audited_portfolio_summary_record(
            self.srf, self.audited, servicer_name="T", excluded_gps=["1"])
        self.assertEqual(omitted["120 Days ($)"], 99666.56)   # fell back to the servicer's value
        self.assertEqual(excluded["120 Days ($)"], 0.0)
        self.assertNotEqual(omitted, excluded)

    def test_unexcluded_loan_is_unaffected(self):
        rec = checks.audited_portfolio_summary_record(
            self.srf, self.audited, servicer_name="T", excluded_gps=["1"])
        self.assertEqual(rec["Current ($)"], 50000.0)


class NotSuppressedRegressionTests(unittest.TestCase):
    """Everything that must continue through the existing logic unchanged."""

    def _run(self, soup_rows, srf_rows):
        soup = _soup(soup_rows)
        srf = _srf(srf_rows)
        results = checks.run_servicer_reconciliation(
            "T", pd.DataFrame(), srf, soup, memo_df=None, memo_lender_key="T")
        stat01 = next(r for r in results if r["id"].split("::")[0] == "STAT-01")
        return stat01, results

    def test_newly_reported_liquidation_with_no_prior_closure_is_not_suppressed(self):
        # The case the whole LIQUIDATION validation discussion is about -- it
        # must still reach the client.
        stat01, results = self._run(
            [("1", "2026-05-01", "ACTIVE", "CURRENT", 80000.0)],
            [("1", "LIQUIDATION", "CURRENT", 80000.0)])
        self.assertEqual(stat01.get("suppressed_closed_loans"), [])
        bal02 = next(r for r in results if r["id"].split("::")[0] == "BAL-02")
        self.assertIn("1", [str(x["SBA GP Number"]) for x in bal02["records"]])

    def test_active_deferred_foreclosure_reo_are_not_suppressed(self):
        for status in ("ACTIVE", "DEFERRED", "FORECLOSURE", "REO"):
            with self.subTest(status=status):
                stat01, _ = self._run(
                    [("1", "2026-05-01", "ACTIVE", "30 DAYS", 80000.0)],
                    [("1", status, "CURRENT", 80000.0)])
                self.assertEqual(stat01.get("suppressed_closed_loans"), [])

    def test_prior_paid_in_full_with_a_nonzero_balance_is_not_suppressed(self):
        stat01, _ = self._run(
            [("1", "2026-05-01", "PAID-IN-FULL", "CURRENT", 4000.0)],
            [("1", "ACTIVE", "CURRENT", 4000.0)])
        self.assertEqual(stat01.get("suppressed_closed_loans"), [])

    def test_prior_zero_balance_while_active_is_not_suppressed(self):
        stat01, _ = self._run(
            [("1", "2026-05-01", "ACTIVE", "CURRENT", 0.0)],
            [("1", "ACTIVE", "CURRENT", 0.0)])
        self.assertEqual(stat01.get("suppressed_closed_loans"), [])
