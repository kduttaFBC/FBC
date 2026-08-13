"""Reporting Exception Register — CLIENT-CONFIRMED (2026-08-12).

An exception records that an ALREADY-ISSUED Monthly Servicer Report carries a
value differing from what this application calculates from the corrected lender
source data, where the client confirmed the issued report is authoritative for
that period and supplied NO general business rule explaining it.

The ten protections the client asked for:

   1  all ten April exceptions reproduce the authoritative client April report
   2  the underlying standard calculation is unchanged
   3  removing the register returns the original lender-derived values
   4  a changed calculated value trips the stale-exception guard
   5  an April exception cannot reach May, June or any other period
   6  an exception for one GP cannot reach another GP
   7  an exception for Net Interest cannot reach another field
   8  June remains identical to the validated June baseline
   9  the complete regression suite passes  (run separately)
  10  no GP number and no April-specific value is embedded in calculation logic

Real-file tests skip unless the fixtures are present.
"""

import inspect
import json
import os
import tempfile
import unittest

import pandas as pd

import checks
import monthly_servicer_report as msr
import reporting_exceptions as rx


APRIL_PERIOD = "2026-04"
APRIL_DEAL = "SOUP2025_1"

# The five April loans and the two affected fields, exactly as recorded.
APRIL_EXPECTED = {
    ("3722547009", "Net Interest"): (223.53, 0.0),
    ("3722547009", "Total Payment Received"): (27823.63, 27600.10),
    ("5758789102", "Net Interest"): (144.82, 0.0),
    ("5758789102", "Total Payment Received"): (19980.30, 19835.48),
    ("6595329103", "Net Interest"): (1723.00, 0.0),
    ("6595329103", "Total Payment Received"): (119548.98, 117825.98),
    ("7098739101", "Net Interest"): (-122.83, -670.64),
    ("7098739101", "Total Payment Received"): (72429.89, 71882.08),
    ("7135239104", "Net Interest"): (549.92, 0.0),
    ("7135239104", "Total Payment Received"): (22308.29, 21758.37),
}
APRIL_TOTAL_DIFFERENCE = -3189.08


def _row(gp, net_interest, total):
    row = {c: None for c in msr.MSR_COLUMNS}
    row["SBA GP Number"] = gp
    row["Net Interest"] = net_interest
    row["Total Payment Received"] = total
    row["Lender/Servicer Name"] = "Test Bank"
    return row


def _entry(gp, field, system_value, authoritative, period=APRIL_PERIOD, deal=APRIL_DEAL):
    return {"reporting_period": period, "deal": deal, "sba_gp_number": gp, "field": field,
            "system_calculated_value": system_value,
            "authoritative_reported_value": authoritative,
            "difference": round(authoritative - system_value, 10),
            "reason": "client-confirmed historical reporting exception",
            "source": "test", "recorded_by": "test", "recorded_on": "2026-08-12",
            "standard_calculation_changed": False}


# ---------------------------------------------------------------------------
# 1. The shipped April register
# ---------------------------------------------------------------------------

class ShippedAprilRegisterTests(unittest.TestCase):

    def setUp(self):
        self.entries = rx.load_register()
        self.index = rx.index_register(self.entries)

    def test_the_register_holds_exactly_ten_entries(self):
        self.assertEqual(len(self.entries), 10)
        self.assertEqual(len(self.index), 10)

    def test_every_entry_is_april_2025_deal_only(self):
        for e in self.entries:
            with self.subTest(gp=e["sba_gp_number"], field=e["field"]):
                self.assertEqual(e["reporting_period"], APRIL_PERIOD)
                self.assertEqual(e["deal"], APRIL_DEAL)

    def test_only_the_two_confirmed_fields_appear(self):
        self.assertEqual({e["field"] for e in self.entries},
                         {"Net Interest", "Total Payment Received"})

    def test_no_other_field_is_present(self):
        for forbidden in ("Ending Actual Balance", "Liquidation Date", "Liquidation Loss",
                          "Modification Comments", "Loan Status", "Delinquency Status",
                          "Net Interest Rate", "Principal", "Beginning Actual Balance"):
            with self.subTest(field=forbidden):
                self.assertNotIn(forbidden, {e["field"] for e in self.entries})

    def test_exactly_the_five_confirmed_loans(self):
        self.assertEqual({e["sba_gp_number"] for e in self.entries},
                         {"3722547009", "5758789102", "6595329103", "7098739101", "7135239104"})

    def test_each_entry_records_calculated_authoritative_and_difference(self):
        for e in self.entries:
            key = (e["sba_gp_number"], e["field"])
            with self.subTest(key=key):
                system, authoritative = APRIL_EXPECTED[key]
                self.assertAlmostEqual(e["system_calculated_value"], system, places=2)
                self.assertAlmostEqual(e["authoritative_reported_value"], authoritative, places=2)
                self.assertAlmostEqual(e["difference"], authoritative - system, places=2)

    def test_each_entry_records_the_reason_and_source(self):
        for e in self.entries:
            with self.subTest(gp=e["sba_gp_number"]):
                self.assertEqual(e["reason"], "client-confirmed historical reporting exception")
                self.assertIn("April", e["source"])
                self.assertIs(e["standard_calculation_changed"], False)

    def test_both_fields_reconcile_to_the_same_total_difference(self):
        for field in ("Net Interest", "Total Payment Received"):
            with self.subTest(field=field):
                total = sum(e["difference"] for e in self.entries if e["field"] == field)
                self.assertAlmostEqual(total, APRIL_TOTAL_DIFFERENCE, places=2)


# ---------------------------------------------------------------------------
# 2, 3. The standard calculation is untouched
# ---------------------------------------------------------------------------

class StandardCalculationUnchangedTests(unittest.TestCase):

    def test_removing_the_register_returns_the_lender_derived_values(self):
        rows = [_row("3722547009", 223.53, 27823.63)]
        msr.build_servicer_rows.__doc__          # touch, no-op
        rx.apply_exceptions(rows, APRIL_PERIOD, APRIL_DEAL, {}, issues=[])
        self.assertEqual(rows[0]["Net Interest"], 223.53)
        self.assertEqual(rows[0]["Total Payment Received"], 27823.63)

    def test_no_register_at_all_changes_nothing(self):
        rows = [_row("3722547009", 223.53, 27823.63)]
        rx.apply_exceptions(rows, APRIL_PERIOD, APRIL_DEAL, None, issues=[])
        self.assertEqual(rows[0]["Net Interest"], 223.53)

    def test_an_empty_register_directory_loads_cleanly(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(rx.load_register(d), [])

    def test_checks_py_contains_no_exception_machinery(self):
        source = inspect.getsource(checks)
        for symbol in ("reporting_exception", "apply_exceptions", "authoritative_reported_value",
                       "system_calculated_value", "EXCEPTION_FIELDS"):
            with self.subTest(symbol=symbol):
                self.assertNotIn(symbol, source)

    def test_consolidate_cycle_postings_is_untouched_by_the_register(self):
        frame = pd.DataFrame([{
            "SBA GP Number": "3722547009", "Lender Loan Number": "L", "Loan Status": "ACTIVE",
            "Delinquency Status": "CURRENT", "Beginning Balance": 27600.10,
            "Ending Balance": 0.0, "Payment Received": 27823.63, "Net Interest": 223.53,
            "Principal": 27600.10, "Scheduled Payment": 100.0, "Servicing Fee": 0.0,
            "Loan Payment Instance": 1,
        }])
        row = checks.consolidate_cycle_postings(frame).loc["3722547009"]
        self.assertEqual(row["net_interest"], 223.53)
        self.assertEqual(row["received"], 27823.63)


# ---------------------------------------------------------------------------
# 4. The stale-exception guard
# ---------------------------------------------------------------------------

class StaleExceptionGuardTests(unittest.TestCase):

    def setUp(self):
        self.index = rx.index_register([_entry("3722547009", "Net Interest", 223.53, 0.0)])

    def test_a_matching_calculated_value_applies(self):
        rows = [_row("3722547009", 223.53, 27823.63)]
        issues = []
        applied = rx.apply_exceptions(rows, APRIL_PERIOD, APRIL_DEAL, self.index, issues=issues)
        self.assertEqual(len(applied), 1)
        self.assertEqual(rows[0]["Net Interest"], 0.0)
        self.assertEqual([i["code"] for i in issues], [rx.ISSUE_EXCEPTION_APPLIED])

    def test_a_changed_calculated_value_is_refused_and_surfaced(self):
        rows = [_row("3722547009", 999.99, 27823.63)]
        issues = []
        applied = rx.apply_exceptions(rows, APRIL_PERIOD, APRIL_DEAL, self.index, issues=issues)
        self.assertEqual(applied, [])
        self.assertEqual(rows[0]["Net Interest"], 999.99)     # untouched
        self.assertEqual([i["code"] for i in issues], [rx.ISSUE_EXCEPTION_STALE])
        self.assertIn("999.99", issues[0]["message"])
        self.assertIn("223.53", issues[0]["message"])

    def test_a_one_cent_change_trips_the_guard(self):
        rows = [_row("3722547009", 223.54, 27823.63)]
        issues = []
        rx.apply_exceptions(rows, APRIL_PERIOD, APRIL_DEAL, self.index, issues=issues)
        self.assertEqual(rows[0]["Net Interest"], 223.54)
        self.assertEqual(issues[0]["code"], rx.ISSUE_EXCEPTION_STALE)

    def test_float_noise_below_half_a_cent_still_applies(self):
        rows = [_row("3722547009", 223.5300000001, 27823.63)]
        issues = []
        rx.apply_exceptions(rows, APRIL_PERIOD, APRIL_DEAL, self.index, issues=issues)
        self.assertEqual(rows[0]["Net Interest"], 0.0)
        self.assertEqual(issues[0]["code"], rx.ISSUE_EXCEPTION_APPLIED)

    def test_a_missing_value_is_refused_not_applied(self):
        rows = [_row("3722547009", None, 27823.63)]
        issues = []
        rx.apply_exceptions(rows, APRIL_PERIOD, APRIL_DEAL, self.index, issues=issues)
        self.assertIsNone(rows[0]["Net Interest"])
        self.assertEqual(issues[0]["code"], rx.ISSUE_EXCEPTION_STALE)


# ---------------------------------------------------------------------------
# 5, 6, 7. An exception is never a rule
# ---------------------------------------------------------------------------

class ExceptionScopeTests(unittest.TestCase):

    def setUp(self):
        self.index = rx.index_register(rx.load_register())

    def test_an_april_exception_cannot_reach_another_period(self):
        for period in ("2026-05", "2026-06", "2026-03", "2025-04", None):
            with self.subTest(period=period):
                rows = [_row("3722547009", 223.53, 27823.63)]
                issues = []
                applied = rx.apply_exceptions(rows, period, APRIL_DEAL, self.index, issues=issues)
                self.assertEqual(applied, [])
                self.assertEqual(rows[0]["Net Interest"], 223.53)

    def test_an_april_exception_cannot_reach_another_deal(self):
        rows = [_row("3722547009", 223.53, 27823.63)]
        applied = rx.apply_exceptions(rows, APRIL_PERIOD, "SOUP2024_1", self.index, issues=[])
        self.assertEqual(applied, [])
        self.assertEqual(rows[0]["Net Interest"], 223.53)

    def test_an_exception_for_one_gp_cannot_reach_another_gp(self):
        rows = [_row("9999999999", 223.53, 27823.63)]
        applied = rx.apply_exceptions(rows, APRIL_PERIOD, APRIL_DEAL, self.index, issues=[])
        self.assertEqual(applied, [])
        self.assertEqual(rows[0]["Net Interest"], 223.53)

    def test_an_exception_for_net_interest_cannot_reach_another_field(self):
        """Only the named field moves, even on a loan that has an exception."""
        index = rx.index_register([_entry("3722547009", "Net Interest", 223.53, 0.0)])
        rows = [_row("3722547009", 223.53, 27823.63)]
        rx.apply_exceptions(rows, APRIL_PERIOD, APRIL_DEAL, index, issues=[])
        self.assertEqual(rows[0]["Net Interest"], 0.0)
        self.assertEqual(rows[0]["Total Payment Received"], 27823.63)   # untouched

    def test_only_the_two_permitted_fields_can_ever_be_targeted(self):
        self.assertEqual(rx.EXCEPTION_FIELDS, ("Net Interest", "Total Payment Received"))

    def test_an_entry_naming_any_other_field_is_rejected(self):
        for field in ("Ending Actual Balance", "Loan Status", "Delinquency Status",
                      "Liquidation Loss", "Principal"):
            with self.subTest(field=field):
                issues = []
                index = rx.index_register([_entry("3722547009", field, 1.0, 2.0)], issues=issues)
                self.assertEqual(index, {})
                self.assertEqual(issues[0]["code"], rx.ISSUE_EXCEPTION_REJECTED)

    def test_an_entry_missing_a_key_is_rejected(self):
        for key in rx.EXCEPTION_KEYS:
            with self.subTest(missing=key):
                entry = _entry("3722547009", "Net Interest", 223.53, 0.0)
                entry[key] = ""
                issues = []
                self.assertEqual(rx.index_register([entry], issues=issues), {})
                self.assertEqual(issues[0]["code"], rx.ISSUE_EXCEPTION_REJECTED)

    def test_an_entry_without_a_recorded_calculated_value_is_rejected(self):
        entry = _entry("3722547009", "Net Interest", 223.53, 0.0)
        del entry["system_calculated_value"]
        issues = []
        self.assertEqual(rx.index_register([entry], issues=issues), {})
        self.assertEqual(issues[0]["code"], rx.ISSUE_EXCEPTION_REJECTED)

    def test_an_unmatched_entry_is_surfaced_not_silent(self):
        issues = []
        rx.unmatched_entries(self.index, APRIL_PERIOD, APRIL_DEAL, set(), issues=issues)
        self.assertTrue(issues)
        self.assertEqual({i["code"] for i in issues}, {rx.ISSUE_EXCEPTION_UNMATCHED})


# ---------------------------------------------------------------------------
# 10. Nothing loan-specific or April-specific lives in calculation logic
# ---------------------------------------------------------------------------

class NoHardcodingTests(unittest.TestCase):

    GPS = ("3722547009", "5758789102", "6595329103", "7098739101", "7135239104")

    def test_no_gp_number_appears_in_any_calculation_module(self):
        import corrections
        import file_readers
        import report
        for module in (checks, corrections, file_readers, report, msr, rx):
            source = inspect.getsource(module)
            for gp in self.GPS:
                with self.subTest(module=module.__name__, gp=gp):
                    self.assertNotIn(gp, source)

    def test_no_april_authoritative_value_appears_in_code(self):
        for module in (checks, msr, rx):
            source = inspect.getsource(module)
            for value in ("27600.10", "19835.48", "117825.98", "71882.08", "21758.37",
                          "3189.08", "670.64", "1723.00"):
                with self.subTest(module=module.__name__, value=value):
                    self.assertNotIn(value, source)

    def test_no_terminal_status_or_payoff_rule_exists(self):
        source = inspect.getsource(rx)
        code = "\n".join(l for l in source.split('"""')[-1].splitlines()
                         if not l.strip().startswith("#"))
        for forbidden in ("PAID-IN-FULL", "SUBSTITUTION-OUT", "payoff", "terminal",
                          "Paid-In-Full Date", "Liquidation"):
            with self.subTest(term=forbidden):
                self.assertNotIn(forbidden, code)

    def test_the_register_lives_in_data_not_code(self):
        directory = rx.register_directory()
        self.assertTrue(os.path.isdir(directory))
        self.assertTrue(all(n.endswith(".json") for n in os.listdir(directory)))

    def test_the_module_never_reads_a_loan_status_to_decide(self):
        code = inspect.getsource(rx).split('"""')[-1]
        for symbol in ("loan_status", "Loan Status", "dq_status"):
            with self.subTest(symbol=symbol):
                self.assertNotIn(symbol, code)


# ---------------------------------------------------------------------------
# Integration through the report-writing layer
# ---------------------------------------------------------------------------

class ReportLayerIntegrationTests(unittest.TestCase):
    """build_servicer_rows applies exceptions only when given a period, deal
    and index; the default call path is unchanged."""

    FRAME = pd.DataFrame([{
        "SBA GP Number": "3722547009", "Lender Loan Number": "L-1", "Loan Status": "ACTIVE",
        "Delinquency Status": "CURRENT", "Beginning Balance": 27600.10, "Ending Balance": 0.0,
        "Payment Received": 27823.63, "Net Interest": 223.53, "Principal": 27600.10,
        "Net Interest Rate": 0.085, "Scheduled Payment": 100.0, "Servicing Fee": 0.0,
        "Loan Payment Instance": 1, "Paid-In-Full Date": None, "Liquidation Date": None,
        "Foreclosure Date": None, "REO Date": None, "Liquidation Proceeds": None,
        "Subsequent Recoveries/(Losses)": None, "Liquidation Loss": None,
        "Modification  Date": None, "Modification Comments (Free-form)": None,
    }])

    def test_the_default_call_applies_no_exception(self):
        rows = msr.build_servicer_rows(self.FRAME, "Community Bank & Trust")
        self.assertEqual(rows[0]["Net Interest"], 223.53)
        self.assertEqual(rows[0]["Total Payment Received"], 27823.63)

    def test_supplying_period_deal_and_index_applies_it(self):
        issues = []
        rows = msr.build_servicer_rows(
            self.FRAME, "Community Bank & Trust", issues=issues,
            reporting_period=APRIL_PERIOD, deal=APRIL_DEAL,
            exception_index=rx.index_register(rx.load_register()))
        self.assertEqual(rows[0]["Net Interest"], 0.0)
        self.assertEqual(rows[0]["Total Payment Received"], 27600.10)
        self.assertEqual({i["code"] for i in issues}, {rx.ISSUE_EXCEPTION_APPLIED})

    def test_the_wrong_period_leaves_the_calculated_values(self):
        rows = msr.build_servicer_rows(
            self.FRAME, "Community Bank & Trust",
            reporting_period="2026-06", deal=APRIL_DEAL,
            exception_index=rx.index_register(rx.load_register()))
        self.assertEqual(rows[0]["Net Interest"], 223.53)

    def test_every_other_column_is_untouched_by_an_exception(self):
        plain = msr.build_servicer_rows(self.FRAME, "Community Bank & Trust")[0]
        excepted = msr.build_servicer_rows(
            self.FRAME, "Community Bank & Trust", reporting_period=APRIL_PERIOD,
            deal=APRIL_DEAL, exception_index=rx.index_register(rx.load_register()))[0]
        for column in msr.MSR_COLUMNS:
            if column in rx.EXCEPTION_FIELDS:
                continue
            with self.subTest(column=column):
                self.assertEqual(plain[column], excepted[column])


if __name__ == "__main__":
    unittest.main()
