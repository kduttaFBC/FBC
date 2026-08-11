"""PORT-01 is labelled "Original Portfolio Summary" (client-confirmed 2026-08-10).

Three Portfolio Summary layers exist, and each must be nameable:

    Original Portfolio Summary       PORT-01 -- what the lender reported
    Audited Portfolio Summary        what the audit calculated
    Final Audited Portfolio Summary  what the reviewer verified (later commit)

The label is the entire change: PORT-01's 29 values are untouched. Previously
it rendered on the results screen as a bare "Portfolio Summary", which is why
its (unaudited) figures were read as the Audited workbook disagreeing with
itself rather than as a different report. Retained deliberately -- the
lender-reported starting position has audit value.

Run with:
    python -m pytest test_portfolio_summary_labels.py
"""

import unittest

import pandas as pd

import checks


def _srf(rows):
    return pd.DataFrame([{
        "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}", "Loan Status": ls,
        "Delinquency Status": dq, "Beginning Balance": bb, "Ending Balance": eb,
        "Payment Received": rec, "Scheduled Payment": 500.0, "Servicing Fee": 0.0,
    } for gp, ls, dq, bb, eb, rec in rows])


SAMPLE = _srf([
    ("1", "ACTIVE", "CURRENT", 100000.0, 99000.0, 500.0),
    ("2", "DEFERRED", "60 DAYS", 50000.0, 50000.0, 0.0),
    ("3", "LIQUIDATION", "120 DAYS", 25000.0, 25000.0, 0.0),
    ("4", "PAID-IN-FULL", "CURRENT", 10000.0, 0.0, 10000.0),
])


class Port01LabelTests(unittest.TestCase):

    def test_port01_is_labelled_original_portfolio_summary(self):
        result = checks.portfolio_summary_check(SAMPLE, "Test Bank")
        self.assertEqual(result["id"], "PORT-01")
        self.assertEqual(result["name"], "Original Portfolio Summary")

    def test_the_check_catalogue_agrees_with_the_check_itself(self):
        # Two independent name sources existed. A stale one is a drift trap
        # even while unused, so they are asserted equal rather than trusted.
        catalogue = dict(checks.ALL_CHECKS_META)
        self.assertEqual(catalogue["PORT-01"],
                         checks.portfolio_summary_check(SAMPLE, "Test Bank")["name"])

    def test_the_label_does_not_claim_to_be_audited(self):
        name = checks.portfolio_summary_check(SAMPLE, "Test Bank")["name"].upper()
        self.assertNotIn("AUDITED", name)
        self.assertNotIn("FINAL", name)
        self.assertIn("ORIGINAL", name)


class Port01ValuesAreUnchangedTests(unittest.TestCase):
    """The relabel must move no number. These assert the full record shape and
    the aging source, so a future change to either fails here first."""

    def test_result_envelope_is_unchanged(self):
        result = checks.portfolio_summary_check(SAMPLE, "Test Bank")
        # severity is INFO, but _result derives status = "PASS" whenever
        # issues_found == 0, so a summary reports PASS. Asserted as the code
        # actually behaves, not as the docstring at the call site reads.
        self.assertEqual(result["severity"], "INFO")
        self.assertEqual(result["status"], "PASS")
        self.assertEqual(result["issues_found"], 0)
        self.assertEqual(result["category"], "SBA Reconciliation")
        self.assertEqual(len(result["records"]), 1)

    def test_all_29_fields_are_present(self):
        record = checks.portfolio_summary_check(SAMPLE, "Test Bank")["records"][0]
        self.assertEqual(len(record), 29)
        self.assertEqual(record["Servicer"], "Test Bank")
        for category in checks._PORTFOLIO_SUMMARY_CATEGORY_ORDER:
            self.assertIn(f"{category} ($)", record)
            self.assertIn(f"{category} (%)", record)

    def test_aging_still_comes_from_the_servicers_own_reported_status(self):
        # PORT-01's defining property: unaudited. Loan 2 reports 60 DAYS and
        # must be bucketed there whatever any audit would have calculated.
        record = checks.portfolio_summary_check(SAMPLE, "Test Bank")["records"][0]
        self.assertEqual(record["60 Days ($)"], 50000.0)
        self.assertEqual(record["Current ($)"], 99000.0)

    def test_the_established_balance_rules_still_hold(self):
        # Outstanding excludes PAID-IN-FULL and LIQUIDATION; those two use
        # BEGINNING balance. Unchanged by the relabel.
        record = checks.portfolio_summary_check(SAMPLE, "Test Bank")["records"][0]
        self.assertEqual(record["Outstanding ($)"], 149000.0)     # loans 1 + 2 only
        self.assertEqual(record["Liquidation ($)"], 25000.0)      # beginning balance
        self.assertEqual(record["Paid In Full ($)"], 10000.0)     # beginning balance


if __name__ == "__main__":
    unittest.main()
