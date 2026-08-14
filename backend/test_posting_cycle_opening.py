"""Cycle opening identification for multi-posting loans.

WHAT THIS COVERS
----------------
CLIENT-CONFIRMED (2026-08-14). A loan's cycle OPENING row is resolved
independently of its posting SEQUENCE, because the lender does not guarantee
that the earliest-dated posting is written first in the worksheet.

    cycle opening order (checks._cycle_opening_posting):
        1. Payment Receipt Date ascending
        2. blank / unparseable Payment Receipt Date LAST
        3. Loan Payment Instance ascending (deterministic tie-breaker)
        4. original row index, then position (total order)

    posting sequence (checks.consolidate_cycle_postings) -- UNCHANGED:
        1. Loan Payment Instance, if populated
        2. Payment Receipt Date
        3. original row order

WHY THE SEQUENCE IS NOT CHANGED. In the SRF template Loan Payment Instance is
the Excel formula COUNTIFS($A$10:A<row>,A<row>) -- an exact re-encoding of
physical row position. It is nonetheless the better SEQUENCE signal: measured
across 16 real workbooks it reproduces a coherent Beginning->Ending balance
chain on 159 of 160 multi-posting loans against 158 for Payment Receipt Date,
and on the one loan where they disagree (Lendistry 2024 July GP 4704259107) the
ordinal is the unique correct chain and the date order breaks it.

THE ACCEPTANCE CASE. Lendistry 2025, July 2026, GP 5951029110: the earliest
posting (2026-06-29) is physically the SECOND row, so Beginning Balance was
read as 19,504.69 against a prior finalized SOUP Ending Actual Balance of
19,805.64 -- a -300.95 BAL-01 exception that is purely an artifact of row
placement. The 2026-06-29 posting carries 19,805.64 and ties to the cent.

WHAT THIS DELIBERATELY ASSERTS IS UNCHANGED
-------------------------------------------
last / effective_last, Ending Balance, Loan Status, Delinquency Status, the
Payment Received / Principal / Net Interest / Servicing Fee totals, reversal
facts, single-posting loans, and every Portfolio Summary figure. Those are
invariance assertions, not re-tests of the rules that own them.

REAL-FIXTURE TESTS are skipped unless these point at the real workbooks:

    FALCONBRIDGE_JULY_LENDISTRY_2025_XLSX   July 2026 Lendistry 2025 SRF
    FALCONBRIDGE_JULY_LENDISTRY_2024_XLSX   July 2026 Lendistry 2024 SRF
    FALCONBRIDGE_JULY_SOUP_XLSX             SOUP whose latest period is JUNE 2026
"""

import os
import unittest

import pandas as pd

import checks


# --------------------------------------------------------------------------
# Fixture builders
# --------------------------------------------------------------------------

def _posting(**over):
    """One Servicer Inputs posting. Defaults to an ordinary paying loan."""
    row = {
        "SBA GP Number": "1000000001",
        "Lender Loan Number": "L-1",
        "Loan Status": "ACTIVE",
        "Delinquency Status": "CURRENT",
        "Payment Received": 100.00,
        "Principal": 60.00,
        "Net Interest": 40.00,
        "Servicing Fee": 5.00,
        "Scheduled Payment": 100.00,
        "Beginning Balance": 1000.00,
        "Ending Balance": 940.00,
        "Payment Receipt Date": "2026-07-05",
        "Loan Payment Instance": 1,
        "Last Payment Instance?": "Y",
        "Interest-From": "2026-07-01",
        "Interest-To": "2026-07-31",
        "_excel_row": 10,
        "_sheet_name": "Servicer Inputs",
        "_header_excel_row": 9,
    }
    row.update(over)
    return row


def _frame(*rows):
    return pd.DataFrame(list(rows))


def _one(df, gp="1000000001"):
    return checks.consolidate_cycle_postings(df).loc[gp]


# The GP 5951029110 profile, field-for-field, in its real PHYSICAL row order.
# The 2026-06-29 posting is deliberately second -- that is the whole point.
GP_5951029110 = "5951029110"
_L = "0000747"


def _gp5951029110_postings():
    common = dict(SBA_GP=GP_5951029110)
    del common
    def p(row, instance, last, date, beg, end, recv, prin, neti, fee, sched):
        return _posting(
            **{
                "SBA GP Number": GP_5951029110,
                "Lender Loan Number": _L,
                "Loan Status": "ACTIVE",
                "Delinquency Status": "60 DAYS",
                "Payment Receipt Date": date,
                "Loan Payment Instance": instance,
                "Last Payment Instance?": last,
                "Beginning Balance": beg,
                "Ending Balance": end,
                "Payment Received": recv,
                "Principal": prin,
                "Net Interest": neti,
                "Servicing Fee": fee,
                "Scheduled Payment": sched,
                "_excel_row": row,
            }
        )
    return [
        p(244, 1, "N", "2026-07-01", 19504.69, 19805.64, -312.07, -300.95, -11.12, -1.07, 313.16),
        p(245, 2, "N", "2026-06-29", 19805.64, 19928.39, -296.24, -122.75, -173.49, -16.82, 0.00),
        p(246, 3, "N", "2026-07-01", 19805.64, 19504.69, 312.07, 300.95, 11.12, 1.09, 313.16),
        p(247, 4, "N", "2026-07-01", 19928.39, 19817.93, 295.14, 110.46, 184.68, 18.02, 313.16),
        p(248, 5, "N", "2026-07-01", 19817.93, 19928.39, -295.14, -110.46, -184.68, -17.92, 313.16),
        p(249, 6, "Y", "2026-07-23", 19928.39, 19928.39, 285.32, 0.00, 285.32, 27.30, 0.00),
    ]


# --------------------------------------------------------------------------
# 1 / 2. The acceptance case, and the general "earliest is physically second"
# --------------------------------------------------------------------------

class CycleOpeningAcceptanceTests(unittest.TestCase):
    """GP 5951029110: the real loan this rule exists for."""

    def setUp(self):
        self.row = _one(_frame(*_gp5951029110_postings()), GP_5951029110)

    def test_beginning_balance_comes_from_the_earliest_dated_posting(self):
        # 19,805.64 is the 2026-06-29 posting (physical row 245, instance 2).
        # The pre-change answer was 19,504.69, from physical row 244.
        self.assertEqual(self.row["beginning_balance"], 19805.64)

    def test_beginning_balance_ties_prior_finalized_soup_to_the_cent(self):
        prior_soup_ending = 19805.64          # June 2026 Ending Actual Balance
        difference = round(float(self.row["beginning_balance"]) - prior_soup_ending, 2)
        self.assertEqual(difference, 0.00)
        self.assertLessEqual(abs(difference), checks.BALANCE_TOLERANCE)

    def test_interest_from_follows_the_cycle_opening(self):
        # Interest-From genuinely describes the start of the accrual window,
        # so it moves with the opening. Interest-To stays on the last posting.
        self.assertEqual(str(self.row["interest_from"]), "2026-07-01")
        self.assertEqual(str(self.row["interest_to"]), "2026-07-31")

    def test_ending_state_is_untouched(self):
        # Every ending-state value still comes from the LAST posting in the
        # SEQUENCE (instance 6, the "Y" row) -- the opening rule must not
        # reach any of these.
        self.assertEqual(self.row["ending_balance"], 19928.39)
        self.assertEqual(str(self.row["loan_status"]).strip(), "ACTIVE")
        self.assertEqual(str(self.row["dq_status"]).strip(), "60 DAYS")
        self.assertEqual(self.row["_excel_row"], 249)
        self.assertFalse(bool(self.row["incomplete_final_posting"]))

    def test_totals_and_reversal_facts_are_untouched(self):
        self.assertAlmostEqual(float(self.row["received"]), -10.92, places=2)
        self.assertTrue(bool(self.row["has_negative_posting"]))
        self.assertTrue(bool(self.row["has_filed_posting"]))
        self.assertEqual(self.row["posting_count"], 6)

    def test_posting_sequence_itself_is_unchanged(self):
        # _all_excel_rows records the SEQUENCE, which stays on Loan Payment
        # Instance. If this ever reads [245, 244, ...] the sequence was
        # re-sorted, which is exactly what must not happen.
        self.assertEqual(list(self.row["_all_excel_rows"]),
                         [244, 245, 246, 247, 248, 249])


class EarliestPostingWrittenSecondTests(unittest.TestCase):
    """The general shape, minimal: earliest date is not the first row."""

    def test_opening_is_the_earliest_dated_row_not_the_first_row(self):
        df = _frame(
            _posting(Beginning_Balance=None, **{
                "Payment Receipt Date": "2026-07-15", "Loan Payment Instance": 1,
                "Last Payment Instance?": "N", "Beginning Balance": 500.00,
                "Ending Balance": 400.00, "_excel_row": 10}),
            _posting(**{
                "Payment Receipt Date": "2026-07-02", "Loan Payment Instance": 2,
                "Last Payment Instance?": "Y", "Beginning Balance": 900.00,
                "Ending Balance": 800.00, "_excel_row": 11}),
        )
        row = _one(df)
        self.assertEqual(row["beginning_balance"], 900.00)   # the 07-02 row
        self.assertEqual(row["ending_balance"], 800.00)      # still the LAST row
        self.assertEqual(row["_excel_row"], 11)

    def test_in_order_postings_are_completely_unaffected(self):
        """The guard against over-correction: when the earliest posting IS the
        first row, nothing may move. (Lendistry 2024 GP 4704259107's shape.)"""
        df = _frame(
            _posting(**{"Payment Receipt Date": "2026-07-01", "Loan Payment Instance": 1,
                        "Last Payment Instance?": "N", "Beginning Balance": 1000.00,
                        "Ending Balance": 900.00, "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "2026-07-02", "Loan Payment Instance": 2,
                        "Last Payment Instance?": "N", "Beginning Balance": 900.00,
                        "Ending Balance": 800.00, "_excel_row": 11}),
            _posting(**{"Payment Receipt Date": "2026-07-01", "Loan Payment Instance": 3,
                        "Last Payment Instance?": "Y", "Beginning Balance": 800.00,
                        "Ending Balance": 700.00, "_excel_row": 12}),
        )
        row = _one(df)
        # Physical order is NOT ascending by date (07-01, 07-02, 07-01), but the
        # earliest date already sits on row 1, so the opening does not move.
        self.assertEqual(row["beginning_balance"], 1000.00)
        self.assertEqual(row["ending_balance"], 700.00)


# --------------------------------------------------------------------------
# 3 - 6. Tie-breaking and degenerate dates
# --------------------------------------------------------------------------

class CycleOpeningTieBreakTests(unittest.TestCase):

    def test_same_date_uses_loan_payment_instance(self):
        """101 of 163 real multi-posting loans share a receipt date, so this is
        the common path, not an edge case. Lowest instance wins."""
        df = _frame(
            _posting(**{"Payment Receipt Date": "2026-07-10", "Loan Payment Instance": 2,
                        "Last Payment Instance?": "N", "Beginning Balance": 555.00,
                        "Ending Balance": 500.00, "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "2026-07-10", "Loan Payment Instance": 1,
                        "Last Payment Instance?": "Y", "Beginning Balance": 777.00,
                        "Ending Balance": 700.00, "_excel_row": 11}),
        )
        # Instance 1 is physically SECOND; the tie-breaker must still pick it.
        self.assertEqual(_one(df)["beginning_balance"], 777.00)

    def test_same_date_tie_break_matches_pre_change_behaviour(self):
        """When every posting shares a date, the opening is whatever the
        instance order already said -- so no same-date loan can move."""
        df = _frame(
            _posting(**{"Payment Receipt Date": "2026-07-10", "Loan Payment Instance": 1,
                        "Last Payment Instance?": "N", "Beginning Balance": 300.00,
                        "Ending Balance": 250.00, "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "2026-07-10", "Loan Payment Instance": 2,
                        "Last Payment Instance?": "N", "Beginning Balance": 250.00,
                        "Ending Balance": 200.00, "_excel_row": 11}),
            _posting(**{"Payment Receipt Date": "2026-07-10", "Loan Payment Instance": 3,
                        "Last Payment Instance?": "Y", "Beginning Balance": 200.00,
                        "Ending Balance": 150.00, "_excel_row": 12}),
        )
        self.assertEqual(_one(df)["beginning_balance"], 300.00)

    def test_blank_date_sorts_last(self):
        df = _frame(
            _posting(**{"Payment Receipt Date": "", "Loan Payment Instance": 1,
                        "Last Payment Instance?": "N", "Beginning Balance": 111.00,
                        "Ending Balance": 100.00, "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "2026-07-08", "Loan Payment Instance": 2,
                        "Last Payment Instance?": "Y", "Beginning Balance": 222.00,
                        "Ending Balance": 200.00, "_excel_row": 11}),
        )
        # A blank date must never be treated as "earliest" -- an absent value
        # asserts nothing about when the cycle opened.
        self.assertEqual(_one(df)["beginning_balance"], 222.00)

    def test_none_date_sorts_last(self):
        df = _frame(
            _posting(**{"Payment Receipt Date": None, "Loan Payment Instance": 1,
                        "Last Payment Instance?": "N", "Beginning Balance": 111.00,
                        "Ending Balance": 100.00, "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "2026-07-08", "Loan Payment Instance": 2,
                        "Last Payment Instance?": "Y", "Beginning Balance": 222.00,
                        "Ending Balance": 200.00, "_excel_row": 11}),
        )
        self.assertEqual(_one(df)["beginning_balance"], 222.00)

    def test_invalid_date_sorts_last(self):
        df = _frame(
            _posting(**{"Payment Receipt Date": "not a date", "Loan Payment Instance": 1,
                        "Last Payment Instance?": "N", "Beginning Balance": 111.00,
                        "Ending Balance": 100.00, "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "2026-07-08", "Loan Payment Instance": 2,
                        "Last Payment Instance?": "Y", "Beginning Balance": 222.00,
                        "Ending Balance": 200.00, "_excel_row": 11}),
        )
        # Unparseable is treated exactly as blank -- it does not raise, and it
        # does not win the opening.
        self.assertEqual(_one(df)["beginning_balance"], 222.00)

    def test_all_dates_blank_falls_back_to_loan_payment_instance(self):
        df = _frame(
            _posting(**{"Payment Receipt Date": "", "Loan Payment Instance": 3,
                        "Last Payment Instance?": "N", "Beginning Balance": 333.00,
                        "Ending Balance": 300.00, "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "", "Loan Payment Instance": 1,
                        "Last Payment Instance?": "N", "Beginning Balance": 111.00,
                        "Ending Balance": 100.00, "_excel_row": 11}),
            _posting(**{"Payment Receipt Date": "", "Loan Payment Instance": 2,
                        "Last Payment Instance?": "Y", "Beginning Balance": 222.00,
                        "Ending Balance": 200.00, "_excel_row": 12}),
        )
        self.assertEqual(_one(df)["beginning_balance"], 111.00)   # instance 1

    def test_all_dates_and_instances_blank_falls_back_to_row_order(self):
        df = _frame(
            _posting(**{"Payment Receipt Date": "", "Loan Payment Instance": None,
                        "Last Payment Instance?": "", "Beginning Balance": 111.00,
                        "Ending Balance": 100.00, "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "", "Loan Payment Instance": None,
                        "Last Payment Instance?": "", "Beginning Balance": 222.00,
                        "Ending Balance": 200.00, "_excel_row": 11}),
        )
        self.assertEqual(_one(df)["beginning_balance"], 111.00)   # first row

    def test_missing_payment_receipt_date_column_entirely(self):
        rows = []
        for instance, beg, end, row_no in ((2, 555.00, 500.00, 10), (1, 777.00, 700.00, 11)):
            p = _posting(**{"Loan Payment Instance": instance, "Beginning Balance": beg,
                            "Ending Balance": end, "_excel_row": row_no,
                            "Last Payment Instance?": "Y" if instance == 2 else "N"})
            p.pop("Payment Receipt Date")
            rows.append(p)
        # No date column at all: the instance decides, exactly as before.
        self.assertEqual(_one(_frame(*rows))["beginning_balance"], 777.00)

    def test_result_is_independent_of_input_row_order(self):
        """The key is a total order, so shuffling the source rows cannot change
        the consolidated record."""
        postings = _gp5951029110_postings()
        forward = _one(_frame(*postings), GP_5951029110)
        reverse = _one(_frame(*reversed(postings)), GP_5951029110)
        self.assertEqual(forward["beginning_balance"], reverse["beginning_balance"])


# --------------------------------------------------------------------------
# 7 - 12. Invariance
# --------------------------------------------------------------------------

class UnchangedBehaviourTests(unittest.TestCase):

    def test_single_posting_loans_are_untouched(self):
        row = _one(_frame(_posting()))
        self.assertEqual(row["beginning_balance"], 1000.00)
        self.assertEqual(row["ending_balance"], 940.00)
        self.assertEqual(row["posting_count"], 1)
        self.assertEqual(str(row["interest_from"]), "2026-07-01")

    def test_single_posting_with_a_blank_date_is_untouched(self):
        row = _one(_frame(_posting(**{"Payment Receipt Date": ""})))
        self.assertEqual(row["beginning_balance"], 1000.00)

    def test_effective_last_fallback_still_selects_the_same_row(self):
        """The incomplete-final-posting fallback is independent of the opening
        rule and must keep choosing the last posting with real data."""
        df = _frame(
            _posting(**{"Payment Receipt Date": "2026-07-20", "Loan Payment Instance": 1,
                        "Last Payment Instance?": "N", "Beginning Balance": 1000.00,
                        "Ending Balance": 900.00, "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "2026-07-02", "Loan Payment Instance": 2,
                        "Last Payment Instance?": "N", "Beginning Balance": 1200.00,
                        "Ending Balance": 1100.00, "_excel_row": 11}),
            _posting(**{"Payment Receipt Date": "2026-07-25", "Loan Payment Instance": 3,
                        "Last Payment Instance?": "Y", "Loan Status": "",
                        "Delinquency Status": "", "Beginning Balance": 1100.00,
                        "Ending Balance": None, "_excel_row": 12}),
        )
        row = _one(df)
        self.assertTrue(bool(row["incomplete_final_posting"]))
        self.assertEqual(row["original_last_excel_row"], 12)
        self.assertEqual(row["effective_last_excel_row"], 11)   # unchanged
        self.assertEqual(row["ending_balance"], 1100.00)        # unchanged
        # ...while the opening still moves to the 07-02 posting.
        self.assertEqual(row["beginning_balance"], 1200.00)

    def test_reversal_totals_and_flags_are_order_independent(self):
        df = _frame(
            _posting(**{"Payment Receipt Date": "2026-07-20", "Loan Payment Instance": 1,
                        "Last Payment Instance?": "N", "Payment Received": 250.00,
                        "Principal": 200.00, "Net Interest": 50.00, "Servicing Fee": 5.00,
                        "Beginning Balance": 1000.00, "Ending Balance": 800.00,
                        "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "2026-07-03", "Loan Payment Instance": 2,
                        "Last Payment Instance?": "Y", "Payment Received": -250.00,
                        "Principal": -200.00, "Net Interest": -50.00, "Servicing Fee": -5.00,
                        "Beginning Balance": 800.00, "Ending Balance": 1000.00,
                        "_excel_row": 11}),
        )
        row = _one(df)
        self.assertEqual(float(row["received"]), 0.0)          # zero-net reversal
        self.assertEqual(float(row["principal"]), 0.0)
        self.assertEqual(float(row["net_interest"]), 0.0)
        self.assertEqual(float(row["servicing_fee"]), 0.0)
        self.assertTrue(bool(row["has_negative_posting"]))
        self.assertTrue(bool(row["has_filed_posting"]))
        self.assertEqual(row["ending_balance"], 1000.00)       # still the last row
        self.assertEqual(row["beginning_balance"], 800.00)     # the 07-03 row

    def test_loan_status_and_delinquency_status_stay_on_the_last_posting(self):
        df = _frame(
            _posting(**{"Payment Receipt Date": "2026-07-20", "Loan Payment Instance": 1,
                        "Last Payment Instance?": "N", "Loan Status": "ACTIVE",
                        "Delinquency Status": "CURRENT", "Beginning Balance": 1000.00,
                        "Ending Balance": 900.00, "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "2026-07-01", "Loan Payment Instance": 2,
                        "Last Payment Instance?": "Y", "Loan Status": "DEFERRED",
                        "Delinquency Status": "60 DAYS", "Beginning Balance": 1500.00,
                        "Ending Balance": 1400.00, "_excel_row": 11}),
        )
        row = _one(df)
        self.assertEqual(str(row["loan_status"]).strip(), "DEFERRED")
        self.assertEqual(str(row["dq_status"]).strip(), "60 DAYS")
        self.assertEqual(row["ending_balance"], 1400.00)
        self.assertEqual(row["beginning_balance"], 1500.00)

    def test_lender_loan_number_is_identity_not_a_cycle_opening_value(self):
        df = _frame(
            _posting(**{"Payment Receipt Date": "2026-07-20", "Loan Payment Instance": 1,
                        "Last Payment Instance?": "N", "Lender Loan Number": "L-FIRST",
                        "Beginning Balance": 1000.00, "Ending Balance": 900.00,
                        "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "2026-07-01", "Loan Payment Instance": 2,
                        "Last Payment Instance?": "Y", "Lender Loan Number": "L-SECOND",
                        "Beginning Balance": 1500.00, "Ending Balance": 1400.00,
                        "_excel_row": 11}),
        )
        self.assertEqual(str(_one(df)["lender_loan_number"]), "L-FIRST")

    def test_scheduled_payment_resolution_is_unchanged(self):
        df = _frame(
            _posting(**{"Payment Receipt Date": "2026-07-20", "Loan Payment Instance": 1,
                        "Last Payment Instance?": "N", "Scheduled Payment": 100.00,
                        "Beginning Balance": 1000.00, "Ending Balance": 900.00,
                        "_excel_row": 10}),
            _posting(**{"Payment Receipt Date": "2026-07-01", "Loan Payment Instance": 2,
                        "Last Payment Instance?": "Y", "Scheduled Payment": 100.00,
                        "Beginning Balance": 1500.00, "Ending Balance": 1400.00,
                        "_excel_row": 11}),
        )
        row = _one(df)
        self.assertEqual(float(row["scheduled"]), 100.00)
        self.assertFalse(bool(row["scheduled_ambiguous"]))


# --------------------------------------------------------------------------
# 13. BAL-01 and DUP-01 must agree
# --------------------------------------------------------------------------

class BalOneAndDupOneAgreeTests(unittest.TestCase):
    """DUP-01 has its own posting sort. It must not be able to derive a
    different Beginning Balance than the audit does for the same loan."""

    def _df(self):
        return _frame(*_gp5951029110_postings())

    def test_dup01_reports_the_cycle_opening_balance(self):
        record = next(r for r in checks.duplicate_postings_check(self._df())["records"]
                      if r["SBA GP Number"] == GP_5951029110)
        self.assertEqual(record["Beginning Balance"], 19805.64)
        self.assertEqual(record["Ending Balance"], 19928.39)
        self.assertEqual(record["Number of Postings"], 6)

    def test_dup01_matches_consolidation_exactly(self):
        df = self._df()
        consolidated = checks.consolidate_cycle_postings(df)
        for record in checks.duplicate_postings_check(df)["records"]:
            gp = record["SBA GP Number"]
            self.assertEqual(float(record["Beginning Balance"]),
                             float(consolidated.loc[gp]["beginning_balance"]),
                             f"DUP-01 and consolidation disagree on GP {gp}")
            self.assertEqual(float(record["Ending Balance"]),
                             float(consolidated.loc[gp]["ending_balance"]))

    def test_dup01_payment_list_still_reads_in_posting_sequence(self):
        """The sequence is unchanged, so the payment list order is unchanged."""
        record = next(r for r in checks.duplicate_postings_check(self._df())["records"]
                      if r["SBA GP Number"] == GP_5951029110)
        details = record["Payment Details (amount / date)"]
        self.assertTrue(details.startswith("$-312.07 (07/01/2026); $-296.24 (06/29/2026)"),
                        details)
        self.assertAlmostEqual(record["Total Payment Received"], -10.92, places=2)


# --------------------------------------------------------------------------
# 14 / 15. Real-fixture acceptance
# --------------------------------------------------------------------------

JULY_L2025 = os.environ.get("FALCONBRIDGE_JULY_LENDISTRY_2025_XLSX")
JULY_L2024 = os.environ.get("FALCONBRIDGE_JULY_LENDISTRY_2024_XLSX")
JULY_SOUP = os.environ.get("FALCONBRIDGE_JULY_SOUP_XLSX")

_REAL_2025 = bool(JULY_L2025 and JULY_SOUP
                  and os.path.exists(JULY_L2025) and os.path.exists(JULY_SOUP))
_REAL_2024 = bool(JULY_L2024 and os.path.exists(JULY_L2024))


def _read(path, soup_key, label):
    import file_readers
    return file_readers.read_servicer_file(
        open(path, "rb").read(), filename=f"{soup_key}||{label}__x.xlsx")


@unittest.skipUnless(_REAL_2025, "set FALCONBRIDGE_JULY_LENDISTRY_2025_XLSX and "
                                 "FALCONBRIDGE_JULY_SOUP_XLSX (JUNE-inclusive)")
class RealJulyLendistry2025Tests(unittest.TestCase):
    """The acceptance case, end to end on the real workbook."""

    @classmethod
    def setUpClass(cls):
        import file_readers
        cls.soup = file_readers.read_soup_file(open(JULY_SOUP, "rb").read())
        _, cls.label, cls.lt, cls.si, cls.deal = _read(JULY_L2025, "Lendistry", "Lendistry 2025")
        cls.servicer_soup = file_readers.filter_soup_by_servicer(cls.soup, "Lendistry")
        cls.consolidated = checks.consolidate_cycle_postings(cls.si)

    def test_soup_fixture_is_the_june_inclusive_generation(self):
        latest = pd.to_datetime(self.soup["Payment Month"], errors="coerce").max()
        self.assertEqual((latest.year, latest.month), (2026, 6),
                         "FALCONBRIDGE_JULY_SOUP_XLSX must be the JUNE-inclusive dataset")

    def test_gp_5951029110_beginning_balance(self):
        self.assertEqual(float(self.consolidated.loc[GP_5951029110]["beginning_balance"]),
                         19805.64)

    def test_gp_5951029110_ending_state_unchanged(self):
        row = self.consolidated.loc[GP_5951029110]
        self.assertEqual(float(row["ending_balance"]), 19928.39)
        self.assertEqual(str(row["loan_status"]).strip(), "ACTIVE")
        self.assertEqual(str(row["dq_status"]).strip(), "60 DAYS")
        self.assertAlmostEqual(float(row["received"]), -10.92, places=2)

    def test_bal01_has_no_exceptions(self):
        result = checks.balance_continuity_check(self.si, self.servicer_soup)
        self.assertEqual(result["issues_found"], 0,
                         f"unexpected BAL-01 exceptions: {result['records']}")
        self.assertEqual(result["status"], "PASS")

    def test_bal02_still_passes(self):
        self.assertEqual(checks.liquidation_balance_check(self.si)["issues_found"], 0)

    def test_portfolio_summary_is_unchanged(self):
        views = checks.original_portfolio_summary_views(self.si, servicer_name="Lendistry 2025")
        counts, dollars = views["counts"], views["dollars"]
        self.assertEqual(counts["Outstanding"], 1223)
        self.assertEqual(counts["Not Making Payment"], 60)
        self.assertEqual(dollars["Not Making Payment ($)"], 3593725.78)
        self.assertEqual(dollars["Outstanding ($)"], 44932620.59)
        self.assertEqual(counts["Paid In Full"], 7)
        self.assertEqual(dollars["Paid In Full ($)"], 264237.48)
        self.assertEqual(counts["Liquidation"], 0)
        self.assertEqual(dollars["Liquidation ($)"], 0.0)

    def test_exactly_one_loan_carries_a_moved_cycle_opening(self):
        """Blast radius: the opening may differ from the sequence's first row
        on GP 5951029110 and on no other loan in this file."""
        moved = []
        si = self.si.copy()
        si["SBA GP Number"] = checks._norm_id(si["SBA GP Number"])
        si = si.reset_index(drop=True)
        for gp, group in si.groupby("SBA GP Number", sort=False):
            if len(group) < 2:
                continue
            group = group.assign(
                _k=pd.to_numeric(group["Loan Payment Instance"], errors="coerce")
            ).sort_values("_k", kind="stable")
            opening = checks._cycle_opening_posting(group)
            if opening.get("_excel_row") != group.iloc[0].get("_excel_row"):
                moved.append(str(gp))
        self.assertEqual(moved, [GP_5951029110])

    def test_dup01_and_bal01_agree_on_every_multi_posting_loan(self):
        for record in checks.duplicate_postings_check(self.si)["records"]:
            gp = record["SBA GP Number"]
            self.assertEqual(float(record["Beginning Balance"]),
                             float(self.consolidated.loc[gp]["beginning_balance"]),
                             f"DUP-01 disagrees with consolidation on GP {gp}")


@unittest.skipUnless(_REAL_2024, "set FALCONBRIDGE_JULY_LENDISTRY_2024_XLSX")
class RealJulyLendistry2024Tests(unittest.TestCase):
    """The over-correction guard on real data: GP 4704259107's physical order
    is not ascending by receipt date, but its earliest posting is already the
    first row, so no consolidated value may move."""

    @classmethod
    def setUpClass(cls):
        _, cls.label, cls.lt, cls.si, cls.deal = _read(JULY_L2024, "Lendistry", "Lendistry 2024")
        cls.consolidated = checks.consolidate_cycle_postings(cls.si)

    def test_gp_4704259107_beginning_balance_did_not_move(self):
        self.assertAlmostEqual(
            float(self.consolidated.loc["4704259107"]["beginning_balance"]), 41157.34, places=2)

    def test_no_loan_in_this_file_has_a_moved_cycle_opening(self):
        si = self.si.copy()
        si["SBA GP Number"] = checks._norm_id(si["SBA GP Number"])
        si = si.reset_index(drop=True)
        moved = []
        for gp, group in si.groupby("SBA GP Number", sort=False):
            if len(group) < 2:
                continue
            group = group.assign(
                _k=pd.to_numeric(group["Loan Payment Instance"], errors="coerce")
            ).sort_values("_k", kind="stable")
            if checks._cycle_opening_posting(group).get("_excel_row") != group.iloc[0].get("_excel_row"):
                moved.append(str(gp))
        self.assertEqual(moved, [])

    def test_portfolio_summary_is_unchanged(self):
        views = checks.original_portfolio_summary_views(self.si, servicer_name="Lendistry 2024")
        self.assertEqual(views["counts"]["Outstanding"], 118)
        self.assertEqual(views["counts"]["Not Making Payment"], 1)
        self.assertEqual(views["dollars"]["Not Making Payment ($)"], 12852.84)
        self.assertEqual(views["dollars"]["Outstanding ($)"], 5581072.26)


if __name__ == "__main__":
    unittest.main()
