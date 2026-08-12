"""BANKRUPTCY Loan Status from a current-period Bankruptcy Date —
CLIENT-CONFIRMED (2026-08-12).

The rule: BANKRUPTCY is valid for a reporting period when THAT period's own
Bankruptcy Date column contains a date. If Ending Balance is $0 because the loan
is PAID IN FULL or LIQUIDATION, the Bankruptcy Date is ignored for this rule.

What this closes. A servicer reporting ACTIVE against a prior SOUP BANKRUPTCY
produced a LOAN STATUS MISMATCH that matched none of corrections.py's three Loan
Status branches, and whose expected value (BANKRUPTCY) is deliberately absent
from _PROPOSABLE_LOAN_STATUS. So no recommendation existed at all: the loan never
reached the Human Review payload for that field and no reviewer could approve,
reject or return it. Observed on Lendistry 2025, July 2026, loans 7126119101,
5917099105, 6288639102, 6649779105 and 6666709108.

The four properties these tests exist to hold, in the client's terms:

  1  a populated current-period Bankruptcy Date on a loan with Ending Balance
     > $0 yields a BANKRUPTCY Loan Status recommendation, review_required
  2  prior SOUP BANKRUPTCY alone NEVER does (the decisive negative test)
  3  the zero-balance exception is preserved: $0 + PAID-IN-FULL / LIQUIDATION
     ignores the Bankruptcy Date
  4  reversal precedence is preserved, and no STAT-01 calculation or
     Delinquency Status figure moves

plus the marker-based write gate: BANKRUPTCY is writable through THIS
recommendation path only, and remains ungeneratable everywhere else.

The real-file tests skip unless the fixtures are present. Run with:
    FALCONBRIDGE_JULY_SOUP_XLSX=... FALCONBRIDGE_MEMO_XLSX=... \\
    FALCONBRIDGE_JULY_LENDISTRY_2025_XLSX=... \\
    python -m pytest test_bankruptcy_date_loan_status.py
"""

import datetime
import io
import os
import unittest

import pandas as pd
from openpyxl import Workbook, load_workbook

import checks
import corrections
import file_readers


SOUP_PATH = os.environ.get("FALCONBRIDGE_JULY_SOUP_XLSX")
MEMO_PATH = os.environ.get("FALCONBRIDGE_MEMO_XLSX")
SRF_PATH = os.environ.get("FALCONBRIDGE_JULY_LENDISTRY_2025_XLSX")
REAL = all([SOUP_PATH, MEMO_PATH, SRF_PATH])

# The five loans this change was raised for, with the current-period evidence
# measured off the real July 2026 Lendistry 2025 submission. Prior SOUP Loan
# Status is BANKRUPTCY for all five and the servicer reports ACTIVE for all five.
#   gp: (bankruptcy_date, ending_balance, liquidation_date_or_None)
JULY_TARGETS = {
    "7126119101": ("2026-02-28", 71252.97, None),
    "5917099105": ("2026-02-15", 10489.57, None),
    "6288639102": ("2026-03-23", 55907.68, None),
    "6649779105": ("2026-03-12", 7450.69, "06/29/2026"),
    "6666709108": ("2026-04-06", 70465.59, "05/26/2026"),
}
# Reports BANKRUPTCY itself, Bankruptcy Date populated, balance > 0. Consistent,
# so STAT-01 raises nothing for it — the control that proves this change adds no
# recommendation where there was no mismatch.
JULY_CONTROL_GP = "6733679102"
# The July run's Delinquency Status recommendation count, which this change must
# not move by one.
JULY_DQ_RECOMMENDATIONS = 86


# --------------------------------------------------------------------------
# Synthetic fixtures
# --------------------------------------------------------------------------

HEADERS = [
    "SBA GP Number", "Loan Status", "Delinquency Status",
    "Beginning Balance", "Ending Balance", "Payment Received", "Scheduled Payment",
    "Bankruptcy Date", "Paid-In-Full Date", "Liquidation Date",
]
HEADER_ROW = 3
SERVICER = "Test Bank"


def _build_workbook_bytes(rows, sheet_name="Servicer Inputs"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.cell(row=1, column=1, value="Servicer Name: Test Bank")
    ws.cell(row=2, column=1, value="Reporting Period: July 2026")
    for col, h in enumerate(HEADERS, start=1):
        ws.cell(row=HEADER_ROW, column=col, value=h)
    excel_rows = []
    for i, row in enumerate(rows):
        r = HEADER_ROW + 1 + i
        excel_rows.append(r)
        for col, h in enumerate(HEADERS, start=1):
            ws.cell(row=r, column=col, value=row.get(h))
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), excel_rows


def _read_cell(ws, row, header):
    return ws.cell(row=row, column=HEADERS.index(header) + 1).value


def _entry(gp="1", excel_rows=(4,), **overrides):
    """A correction_data entry shaped exactly as checks.py emits it, defaulting
    to the real July situation: LOAN STATUS MISMATCH, servicer says ACTIVE, prior
    SOUP says BANKRUPTCY, expected_loan_status_raw is None (BANKRUPTCY reached the
    display column only through the `or soup_status` fallback)."""
    entry = {
        "gp": gp, "excel_row": list(excel_rows)[-1], "excel_rows": list(excel_rows),
        "sheet_name": "Servicer Inputs", "header_excel_row": HEADER_ROW,
        "reasons": ["LOAN STATUS MISMATCH"],
        "dq_match": True, "calculated_dq": "90 DAYS", "srf_dq": "90 DAYS",
        "status_match": False,
        "expected_loan_status_raw": None,
        "expected_loan_status_effective": "BANKRUPTCY",
        "srf_status": "ACTIVE", "soup_status": "BANKRUPTCY",
        "deferral_type": None, "deferral_phase": None,
        "evidence_status": "no-deferral",
        "deferral_start": "", "deferral_end": "", "deferral_resume": "",
        "securitization_action": "",
        "reversal_equivalents": None, "has_negative_posting": False,
        "offsetting_zero_net_reversal": False,
        "bankruptcy_date": "2026-03-12", "bankruptcy_date_populated": True,
        "paid_in_full_date": "", "paid_in_full_date_populated": False,
        "liquidation_date": "", "liquidation_date_populated": False,
        "prior_dq": "90 DAYS", "payment_received": None, "scheduled_payment": None,
        "servicing_fee": 6.12, "ending_balance": 7450.69,
        "posting_count": len(excel_rows),
    }
    entry.update(overrides)
    return entry


def _loan_status_recs(correction_data, servicer=SERVICER):
    recs, candidates, _ = corrections.build_recommendations(
        {"correction_data": correction_data, "_cycle_month": None}, servicer)
    return ([r for r in recs if r["field"] == "Loan Status"],
            [c for c in candidates if c["field"] == "Loan Status"])


def _generate(file_bytes, candidates, decisions):
    return corrections.generate_reviewed_workbook(
        file_bytes, "test.xlsx", SERVICER, candidates, decisions,
        "K. Dutta", "sha256-fixture", None)


# ==========================================================================
# 1. Positive — a populated current-period Bankruptcy Date proposes BANKRUPTCY
# ==========================================================================

class BankruptcyDateProposesBankruptcyTests(unittest.TestCase):

    def test_populated_date_with_positive_balance_yields_review_required_recommendation(self):
        recs, _ = _loan_status_recs([_entry()])
        self.assertEqual(len(recs), 1)
        self.assertEqual(recs[0]["calculated"], "BANKRUPTCY")
        self.assertEqual(recs[0]["reported"], "ACTIVE")
        self.assertTrue(recs[0]["reviewFlag"])
        self.assertIn(corrections.REVIEW_REASON_BANKRUPTCY_CURRENT_PERIOD, recs[0]["reviewFlag"])

    def test_recommendation_is_never_auto_applied(self):
        _, candidates = _loan_status_recs([_entry()])
        self.assertTrue(candidates[0]["review_required"])

    def test_current_period_evidence_reaches_the_review_payload(self):
        """The reviewer must be able to see the date the proposal rests on
        without opening the workbook."""
        recs, _ = _loan_status_recs([_entry()])
        self.assertEqual(recs[0]["bankruptcyDate"], "2026-03-12")
        self.assertEqual(recs[0]["liquidationDate"], "")
        self.assertEqual(recs[0]["paidInFullDate"], "")
        self.assertEqual(recs[0]["balance"], 7450.69)

    def test_evidence_marker_is_not_echoed_to_the_client(self):
        """It is a write gate read off the session candidate. Echoing it would
        invite it back as client input."""
        recs, candidates = _loan_status_recs([_entry()])
        self.assertNotIn("evidence_marker", recs[0])
        self.assertNotIn("evidenceMarker", recs[0])
        self.assertEqual(candidates[0]["evidence_marker"],
                         corrections.CURRENT_PERIOD_EVIDENCE_MARKER)

    def test_a_balance_of_one_cent_is_still_outstanding(self):
        recs, _ = _loan_status_recs([_entry(ending_balance=0.01)])
        self.assertEqual(len(recs), 1)

    def test_floating_point_residue_below_a_cent_is_not_outstanding(self):
        """$0 is judged at cent precision, the precision every figure here is
        compared at — so residue behind a paid-off balance is still $0."""
        recs, _ = _loan_status_recs([_entry(ending_balance=0.000001)])
        self.assertEqual(recs, [])


# ==========================================================================
# 2. Negative — prior SOUP BANKRUPTCY alone is NEVER enough
# ==========================================================================

class PriorSoupBankruptcyAloneIsNotEvidenceTests(unittest.TestCase):
    """THE decisive test. The client was explicit that prior SOUP BANKRUPTCY must
    not on its own support the proposal, so soup_status is not merely
    insufficient in the new branch — it is not consulted at all."""

    def test_blank_current_bankruptcy_date_yields_no_recommendation(self):
        recs, _ = _loan_status_recs([_entry(
            soup_status="BANKRUPTCY",
            expected_loan_status_effective="BANKRUPTCY",
            bankruptcy_date="", bankruptcy_date_populated=False,
            ending_balance=50000.0,
        )])
        self.assertEqual(recs, [])

    def test_servicer_reported_status_stands_when_no_current_source_supports_bankruptcy(self):
        """No recommendation means the lender's ACTIVE is what the corrected
        workbook keeps — nothing is written to that cell."""
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "90 DAYS",
             "Beginning Balance": 50000, "Ending Balance": 50000, "Bankruptcy Date": ""},
        ])
        recs, candidates = _loan_status_recs([_entry(
            excel_rows=excel_rows, bankruptcy_date="", bankruptcy_date_populated=False,
            ending_balance=50000.0)])
        self.assertEqual(recs, [])
        self.assertEqual(candidates, [])
        _, ws = load_workbook(io.BytesIO(file_bytes)), None
        wb = load_workbook(io.BytesIO(file_bytes))
        self.assertEqual(_read_cell(wb["Servicer Inputs"], excel_rows[0], "Loan Status"), "ACTIVE")

    def test_whitespace_only_date_is_not_populated(self):
        """checks.py decides populated-ness; this asserts the branch trusts that
        verdict rather than re-deriving truthiness from the display string."""
        recs, _ = _loan_status_recs([_entry(
            bankruptcy_date="   ", bankruptcy_date_populated=False)])
        self.assertEqual(recs, [])


# ==========================================================================
# 3. The client's zero-balance exception
# ==========================================================================

class ZeroBalanceExceptionTests(unittest.TestCase):
    """If Ending Balance is $0 because the loan is PAID IN FULL or LIQUIDATION,
    the Bankruptcy Date is ignored for this rule."""

    def test_zero_balance_paid_in_full_status_ignores_bankruptcy_date(self):
        recs, _ = _loan_status_recs([_entry(
            srf_status="PAID-IN-FULL", ending_balance=0.0,
            paid_in_full_date="2026-07-15", paid_in_full_date_populated=True)])
        self.assertEqual(recs, [])

    def test_zero_balance_liquidation_status_ignores_bankruptcy_date(self):
        recs, _ = _loan_status_recs([_entry(
            srf_status="LIQUIDATION", ending_balance=0.0,
            liquidation_date="2026-07-20", liquidation_date_populated=True)])
        self.assertEqual(recs, [])

    def test_zero_balance_alone_ignores_bankruptcy_date(self):
        """Gate 2 excludes every $0 balance regardless of why it is $0 — a loan
        with nothing outstanding is never re-proposed as BANKRUPTCY."""
        recs, _ = _loan_status_recs([_entry(ending_balance=0.0)])
        self.assertEqual(recs, [])

    def test_paid_in_full_date_blocks_even_with_a_positive_balance(self):
        """The belt-and-braces half: a finished loan is refused even where gate 2
        would let it through, so a stale positive balance cannot resurrect it."""
        recs, _ = _loan_status_recs([_entry(
            ending_balance=1234.56,
            paid_in_full_date="2026-07-15", paid_in_full_date_populated=True)])
        self.assertEqual(recs, [])

    def test_liquidation_reported_status_blocks_even_with_a_positive_balance(self):
        recs, _ = _loan_status_recs([_entry(srf_status="LIQUIDATION", ending_balance=1234.56)])
        self.assertEqual(recs, [])

    def test_missing_balance_is_not_treated_as_outstanding(self):
        recs, _ = _loan_status_recs([_entry(ending_balance=None)])
        self.assertEqual(recs, [])


# ==========================================================================
# 4. Flag-and-propose for a contradictory Liquidation Date
# ==========================================================================

class LiquidationDateContradictionTests(unittest.TestCase):
    """CLIENT-CONFIRMED (2026-08-12): flag and propose, do NOT suppress. The
    zero-balance exception is keyed on a $0 balance, so a Liquidation Date beside
    a POSITIVE balance leaves the Bankruptcy Date standing — but the reviewer must
    see the contradiction prominently."""

    def _contradicted(self):
        return _entry(liquidation_date="06/29/2026", liquidation_date_populated=True,
                      ending_balance=7450.69)

    def test_recommendation_is_still_produced(self):
        recs, _ = _loan_status_recs([self._contradicted()])
        self.assertEqual(len(recs), 1)
        self.assertEqual(recs[0]["calculated"], "BANKRUPTCY")

    def test_both_reasons_are_stated(self):
        recs, _ = _loan_status_recs([self._contradicted()])
        self.assertIn(corrections.REVIEW_REASON_BANKRUPTCY_CURRENT_PERIOD, recs[0]["reviewFlag"])
        self.assertIn(corrections.REVIEW_REASON_BANKRUPTCY_LIQUIDATION_CONTRADICTION,
                      recs[0]["reviewFlag"])

    def test_liquidation_date_and_balance_both_reach_the_reviewer(self):
        recs, _ = _loan_status_recs([self._contradicted()])
        self.assertEqual(recs[0]["liquidationDate"], "06/29/2026")
        self.assertEqual(recs[0]["balance"], 7450.69)

    def test_uncontradicted_row_keeps_the_single_reason_verbatim(self):
        """With one reason present the joined string is that reason exactly, so
        the uncontradicted wording never drifts."""
        recs, _ = _loan_status_recs([_entry()])
        self.assertEqual(recs[0]["reviewFlag"],
                         corrections.REVIEW_REASON_BANKRUPTCY_CURRENT_PERIOD)


# ==========================================================================
# 5. Reversal precedence
# ==========================================================================

class ReversalPrecedencePreservedTests(unittest.TestCase):
    """The new branch sits after the existing `if _reversal: continue`, so a
    reversal still blocks every Loan Status recommendation."""

    def test_net_negative_reversal_blocks_the_bankruptcy_recommendation(self):
        recs, _ = _loan_status_recs([_entry(
            reversal_equivalents=1, has_negative_posting=True,
            reasons=["LOAN STATUS MISMATCH", "DELINQUENCY STATUS MISMATCH"],
            dq_match=False, calculated_dq="DELINQUENT", srf_dq="60 DAYS")])
        self.assertEqual(recs, [])

    def test_offsetting_zero_net_reversal_blocks_the_bankruptcy_recommendation(self):
        recs, _ = _loan_status_recs([_entry(
            offsetting_zero_net_reversal=True, has_negative_posting=True,
            reasons=["LOAN STATUS MISMATCH", "DELINQUENCY STATUS MISMATCH"],
            dq_match=False, calculated_dq="60 DAYS", srf_dq="90 DAYS")])
        self.assertEqual(recs, [])

    def test_the_delinquency_recommendation_still_reaches_review(self):
        """Blocking Loan Status must not resurrect the old defect of dropping the
        loan from Human Review altogether."""
        all_recs, _, _ = corrections.build_recommendations({"correction_data": [_entry(
            reversal_equivalents=1, has_negative_posting=True,
            reasons=["LOAN STATUS MISMATCH", "DELINQUENCY STATUS MISMATCH"],
            dq_match=False, calculated_dq="DELINQUENT", srf_dq="60 DAYS")],
            "_cycle_month": None}, SERVICER)
        self.assertEqual([r["field"] for r in all_recs], ["Delinquency Status"])


# ==========================================================================
# 6. No proposal when the servicer already agrees
# ==========================================================================

class ServicerAlreadyReportsBankruptcyTests(unittest.TestCase):

    def test_no_recommendation_when_reported_status_is_already_bankruptcy(self):
        recs, _ = _loan_status_recs([_entry(srf_status="BANKRUPTCY")])
        self.assertEqual(recs, [])

    def test_no_recommendation_without_a_loan_status_mismatch_reason(self):
        recs, _ = _loan_status_recs([_entry(reasons=["DELINQUENCY STATUS MISMATCH"],
                                           status_match=True)])
        self.assertEqual(recs, [])


# ==========================================================================
# 7. The marker-based write gate
# ==========================================================================

class MarkerBasedWriteGateTests(unittest.TestCase):
    """BANKRUPTCY is writable through this recommendation path only. The global
    generated domain is deliberately unchanged."""

    def test_proposable_domain_still_refuses_bankruptcy(self):
        self.assertNotIn("BANKRUPTCY", corrections._PROPOSABLE_LOAN_STATUS)
        self.assertIsNone(corrections._canonical_proposable_loan_status("BANKRUPTCY"))

    def test_generated_bankruptcy_refused_without_the_marker(self):
        ok, reason = corrections.validate_approved_value(
            "Loan Status", "BANKRUPTCY", substituted=False, evidence_marker=None)
        self.assertFalse(ok)
        self.assertIn("may generate", reason)

    def test_generated_bankruptcy_allowed_with_the_marker(self):
        ok, value = corrections.validate_approved_value(
            "Loan Status", "BANKRUPTCY", substituted=False,
            evidence_marker=corrections.CURRENT_PERIOD_EVIDENCE_MARKER)
        self.assertTrue(ok)
        self.assertEqual(value, "BANKRUPTCY")

    def test_marker_does_not_unlock_other_terminal_statuses(self):
        """It widens the generated domain by exactly one value."""
        for status in ("FORECLOSURE", "LIQUIDATION", "REO", "SUBSTITUTION-OUT"):
            ok, _ = corrections.validate_approved_value(
                "Loan Status", status, substituted=False,
                evidence_marker=corrections.CURRENT_PERIOD_EVIDENCE_MARKER)
            self.assertFalse(ok, f"{status} must not be generatable with the marker")

    def test_marker_does_not_unlock_a_cross_domain_value(self):
        ok, _ = corrections.validate_approved_value(
            "Loan Status", "90 DAYS", substituted=False,
            evidence_marker=corrections.CURRENT_PERIOD_EVIDENCE_MARKER)
        self.assertFalse(ok)

    def test_approve_with_no_approved_value_applies(self):
        """The API shape: approve the value the application derived, sending no
        approved_value at all (substituted=False). This is the path that would
        have failed had the gate relied on the UI always populating the field."""
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "90 DAYS",
             "Beginning Balance": 7450.69, "Ending Balance": 7450.69,
             "Bankruptcy Date": datetime.date(2026, 3, 12)},
        ])
        recs, candidates = _loan_status_recs([_entry(excel_rows=excel_rows)])
        outcome = _generate(file_bytes, candidates,
                            [{"id": recs[0]["id"], "decision": "approve"}])
        self.assertEqual(outcome["corrections_not_applied"], [])
        self.assertEqual(outcome["applied_count"], 1)
        wb = load_workbook(io.BytesIO(outcome["file_bytes"]))
        self.assertEqual(_read_cell(wb["Servicer Inputs"], excel_rows[0], "Loan Status"),
                         "BANKRUPTCY")
        self.assertEqual(outcome["verified_overrides"], {("1", "Loan Status"): "BANKRUPTCY"})

    def test_approve_with_the_ui_shape_applies(self):
        """CorrectionReview.jsx always sends approved_value (substituted=True)."""
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "90 DAYS",
             "Beginning Balance": 7450.69, "Ending Balance": 7450.69,
             "Bankruptcy Date": datetime.date(2026, 3, 12)},
        ])
        recs, candidates = _loan_status_recs([_entry(excel_rows=excel_rows)])
        outcome = _generate(file_bytes, candidates, [
            {"id": recs[0]["id"], "decision": "approve", "approved_value": "BANKRUPTCY"}])
        self.assertEqual(outcome["applied_count"], 1)
        wb = load_workbook(io.BytesIO(outcome["file_bytes"]))
        self.assertEqual(_read_cell(wb["Servicer Inputs"], excel_rows[0], "Loan Status"),
                         "BANKRUPTCY")

    def test_marker_cannot_be_minted_from_the_decision_payload(self):
        """The marker is read off the SERVER-SIDE candidate. A caller adding it
        to their decision must not unlock a BANKRUPTCY write on a candidate that
        never carried the evidence."""
        file_bytes, excel_rows = _build_workbook_bytes([
            {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "90 DAYS",
             "Beginning Balance": 50000, "Ending Balance": 50000, "Bankruptcy Date": ""},
        ])
        # A DEFERRED Loan Status candidate — a legitimately generated one that
        # carries no current-period bankruptcy evidence.
        recs, candidates = _loan_status_recs([_entry(
            excel_rows=excel_rows,
            expected_loan_status_raw="DEFERRED", deferral_type="full",
            deferral_phase="mid", evidence_status="documented",
            bankruptcy_date="", bankruptcy_date_populated=False,
            ending_balance=50000.0)])
        self.assertEqual(len(recs), 1)
        self.assertEqual(recs[0]["calculated"], "DEFERRED")
        outcome = _generate(file_bytes, candidates, [{
            "id": recs[0]["id"], "decision": "approve",
            "evidence_marker": corrections.CURRENT_PERIOD_EVIDENCE_MARKER,
            "approved_value": None,
        }])
        wb = load_workbook(io.BytesIO(outcome["file_bytes"]))
        self.assertEqual(_read_cell(wb["Servicer Inputs"], excel_rows[0], "Loan Status"),
                         "DEFERRED")

    def test_reject_and_return_leave_the_cell_untouched(self):
        for decision in ("reject", "return"):
            file_bytes, excel_rows = _build_workbook_bytes([
                {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "90 DAYS",
                 "Beginning Balance": 7450.69, "Ending Balance": 7450.69,
                 "Bankruptcy Date": datetime.date(2026, 3, 12)},
            ])
            recs, candidates = _loan_status_recs([_entry(excel_rows=excel_rows)])
            outcome = _generate(file_bytes, candidates,
                                [{"id": recs[0]["id"], "decision": decision}])
            self.assertEqual(outcome["applied_count"], 0, decision)
            wb = load_workbook(io.BytesIO(outcome["file_bytes"]))
            self.assertEqual(_read_cell(wb["Servicer Inputs"], excel_rows[0], "Loan Status"),
                             "ACTIVE", decision)


# ==========================================================================
# 8. The other three Loan Status branches are unchanged
# ==========================================================================

class ExistingBranchesUnchangedTests(unittest.TestCase):

    def test_terminal_branch_cannot_propose_bankruptcy_from_prior_soup(self):
        """A FORECLOSURE-reporting loan whose SOUP history says BANKRUPTCY and
        whose current Bankruptcy Date is blank gets no BANKRUPTCY proposal from
        the terminal-contradiction branch. This is what widening
        _PROPOSABLE_LOAN_STATUS would have broken."""
        recs, _ = _loan_status_recs([_entry(
            srf_status="FORECLOSURE", soup_status="BANKRUPTCY",
            expected_loan_status_effective="BANKRUPTCY",
            bankruptcy_date="", bankruptcy_date_populated=False,
            ending_balance=50000.0)])
        self.assertEqual(recs, [])

    def test_deferred_branch_still_emits(self):
        recs, _ = _loan_status_recs([_entry(
            expected_loan_status_raw="DEFERRED", deferral_type="full",
            deferral_phase="mid", evidence_status="documented",
            bankruptcy_date="", bankruptcy_date_populated=False)])
        self.assertEqual([r["calculated"] for r in recs], ["DEFERRED"])

    def test_bankruptcy_and_deferred_do_not_both_fire(self):
        """A confirmed full-P&I deferral with a stray Bankruptcy Date would
        otherwise produce two contradictory Loan Status recommendations for one
        loan, whose ids collide on (servicer, gp, field)."""
        recs, _ = _loan_status_recs([_entry(
            expected_loan_status_raw="DEFERRED", deferral_type="full",
            deferral_phase="mid", evidence_status="documented",
            bankruptcy_date="2026-03-12", bankruptcy_date_populated=True)])
        self.assertEqual(len(recs), 1, [r["calculated"] for r in recs])
        self.assertEqual(len({r["id"] for r in recs}), len(recs))


# ==========================================================================
# 9. consolidate_cycle_postings carries the raw facts
# ==========================================================================

class ConsolidationCarriesTerminalEventDatesTests(unittest.TestCase):

    def _consolidate(self, rows):
        df = pd.DataFrame(rows)
        df["_sheet_name"] = "Servicer Inputs"
        df["_header_excel_row"] = HEADER_ROW
        df["_excel_row"] = list(range(HEADER_ROW + 1, HEADER_ROW + 1 + len(rows)))
        return checks.consolidate_cycle_postings(df)

    def _base(self, **over):
        row = {"SBA GP Number": "1", "Loan Status": "ACTIVE", "Delinquency Status": "90 DAYS",
               "Beginning Balance": 7450.69, "Ending Balance": 7450.69,
               "Payment Received": 0.0, "Scheduled Payment": 0.0,
               "Bankruptcy Date": "", "Paid-In-Full Date": "", "Liquidation Date": ""}
        row.update(over)
        return row

    def test_a_date_object_counts_as_populated(self):
        cons = self._consolidate([self._base(**{"Bankruptcy Date": datetime.date(2026, 3, 12)})])
        self.assertTrue(cons.loc["1", "bankruptcy_date_populated"])
        self.assertEqual(cons.loc["1", "bankruptcy_date"], "2026-03-12")

    def test_a_string_date_counts_as_populated(self):
        cons = self._consolidate([self._base(**{"Liquidation Date": "06/29/2026"})])
        self.assertTrue(cons.loc["1", "liquidation_date_populated"])
        self.assertEqual(cons.loc["1", "liquidation_date"], "06/29/2026")

    def test_blank_whitespace_and_none_are_not_populated(self):
        for blank in ("", "   ", None):
            cons = self._consolidate([self._base(**{"Bankruptcy Date": blank})])
            self.assertFalse(cons.loc["1", "bankruptcy_date_populated"], repr(blank))
            self.assertEqual(cons.loc["1", "bankruptcy_date"], "")

    def test_populated_on_any_posting_counts(self):
        """A reversal pair or an incomplete final posting can leave the date on
        one row and blank on another. The loan does not stop having a bankruptcy
        date because the last row omitted it."""
        cons = self._consolidate([
            self._base(**{"Bankruptcy Date": datetime.date(2026, 3, 12)}),
            self._base(**{"Bankruptcy Date": ""}),
        ])
        self.assertTrue(cons.loc["1", "bankruptcy_date_populated"])
        self.assertEqual(cons.loc["1", "bankruptcy_date"], "2026-03-12")

    def test_missing_columns_degrade_to_not_populated(self):
        """A servicer template without these columns must not raise."""
        df = pd.DataFrame([{"SBA GP Number": "1", "Loan Status": "ACTIVE",
                            "Delinquency Status": "CURRENT", "Ending Balance": 100.0,
                            "Beginning Balance": 100.0, "_sheet_name": "Servicer Inputs",
                            "_header_excel_row": HEADER_ROW, "_excel_row": 4}])
        cons = checks.consolidate_cycle_postings(df)
        self.assertFalse(cons.loc["1", "bankruptcy_date_populated"])
        self.assertEqual(cons.loc["1", "bankruptcy_date"], "")

    def test_empty_frame_declares_the_new_columns(self):
        """Zero loans still has to return the full column contract, or every
        consumer indexing a new column breaks on an empty submission."""
        cons = checks.consolidate_cycle_postings(
            pd.DataFrame(columns=["SBA GP Number", "Loan Status", "Delinquency Status",
                                  "Ending Balance", "Beginning Balance", "_excel_row",
                                  "_sheet_name", "_header_excel_row"]))
        self.assertTrue(cons.empty)
        for col in ("bankruptcy_date", "bankruptcy_date_populated",
                    "paid_in_full_date", "paid_in_full_date_populated",
                    "liquidation_date", "liquidation_date_populated"):
            self.assertIn(col, cons.columns)


# ==========================================================================
# 10. Real July 2026 Lendistry 2025 fixtures
# ==========================================================================

@unittest.skipUnless(REAL, "set FALCONBRIDGE_JULY_SOUP_XLSX / _MEMO_XLSX / "
                           "_JULY_LENDISTRY_2025_XLSX")
class JulyLendistry2025RealFixtureTests(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        soup_df = file_readers.read_soup_file(open(SOUP_PATH, "rb").read())
        memo_df = file_readers.read_service_memo(open(MEMO_PATH, "rb").read())
        soup_key, display, loan_terms, inputs, deal = file_readers.read_servicer_file(
            open(SRF_PATH, "rb").read(),
            filename="Lendistry||Lendistry 2025__july.xlsx")
        cls.display = display
        cls.inputs = inputs
        results = checks.run_servicer_reconciliation(
            display, loan_terms, inputs,
            file_readers.filter_soup_by_servicer(soup_df, soup_key),
            memo_df=memo_df, memo_lender_key=soup_key, deal=deal)
        cls.stat01 = next(r for r in results if r["id"].split("::")[0] == "STAT-01")
        cls.correction_data = {str(c["gp"]).strip(): c for c in cls.stat01["correction_data"]}
        recs, candidates, skipped = corrections.build_recommendations(cls.stat01, display)
        cls.recs = recs
        cls.candidates = {c["id"]: c for c in candidates}
        cls.skipped = skipped
        cls.by_gp = {}
        for r in recs:
            cls.by_gp.setdefault(str(r["gp"]).strip(), []).append(r)

    def test_every_target_loan_now_gets_a_bankruptcy_loan_status_recommendation(self):
        for gp, (bk_date, balance, liq) in JULY_TARGETS.items():
            with self.subTest(gp=gp):
                ls = [r for r in self.by_gp.get(gp, []) if r["field"] == "Loan Status"]
                self.assertEqual(len(ls), 1)
                self.assertEqual(ls[0]["calculated"], "BANKRUPTCY")
                self.assertEqual(ls[0]["reported"], "ACTIVE")
                self.assertEqual(ls[0]["priorLoanStatus"], "BANKRUPTCY")
                self.assertEqual(ls[0]["bankruptcyDate"], bk_date)
                self.assertAlmostEqual(ls[0]["balance"], balance, places=2)
                self.assertTrue(ls[0]["reviewFlag"])

    def test_the_two_contradicted_loans_are_flagged_not_suppressed(self):
        for gp in ("6649779105", "6666709108"):
            with self.subTest(gp=gp):
                ls = [r for r in self.by_gp[gp] if r["field"] == "Loan Status"]
                self.assertEqual(len(ls), 1)
                self.assertEqual(ls[0]["liquidationDate"], JULY_TARGETS[gp][2])
                self.assertGreater(ls[0]["balance"], 0)
                self.assertIn(corrections.REVIEW_REASON_BANKRUPTCY_LIQUIDATION_CONTRADICTION,
                              ls[0]["reviewFlag"])

    def test_the_three_uncontradicted_loans_carry_only_the_bankruptcy_reason(self):
        for gp in ("7126119101", "5917099105", "6288639102"):
            with self.subTest(gp=gp):
                ls = [r for r in self.by_gp[gp] if r["field"] == "Loan Status"]
                self.assertEqual(ls[0]["reviewFlag"],
                                 corrections.REVIEW_REASON_BANKRUPTCY_CURRENT_PERIOD)

    def test_control_loan_is_untouched(self):
        """6733679102 reports BANKRUPTCY itself with the date populated and a
        positive balance. Consistent, so STAT-01 raises nothing and this change
        must add nothing."""
        self.assertNotIn(JULY_CONTROL_GP, self.correction_data)
        self.assertEqual(self.by_gp.get(JULY_CONTROL_GP, []), [])

    def test_delinquency_recommendation_count_is_unchanged(self):
        dq = [r for r in self.recs if r["field"] == "Delinquency Status"]
        self.assertEqual(len(dq), JULY_DQ_RECOMMENDATIONS)

    def test_total_recommendations_are_the_dq_set_plus_exactly_five(self):
        ls = [r for r in self.recs if r["field"] == "Loan Status"]
        self.assertEqual(len(ls), len(JULY_TARGETS))
        self.assertEqual({str(r["gp"]).strip() for r in ls}, set(JULY_TARGETS))
        self.assertEqual(len(self.recs), JULY_DQ_RECOMMENDATIONS + len(JULY_TARGETS))
        self.assertEqual(self.skipped, [])

    def test_delinquency_figures_for_the_targets_are_unchanged(self):
        """No STAT-01 calculation and no ladder result may move."""
        expected = {
            "7126119101": ("150 DAYS", "120 DAYS"),
            "5917099105": ("180 DAYS", "150 DAYS"),
            "6288639102": ("120 DAYS", "90 DAYS"),
            "6649779105": ("120 DAYS", "90 DAYS"),
            "6666709108": ("150 DAYS", "120 DAYS"),
        }
        for gp, (reported, calculated) in expected.items():
            with self.subTest(gp=gp):
                dq = [r for r in self.by_gp[gp] if r["field"] == "Delinquency Status"]
                self.assertEqual(len(dq), 1)
                self.assertEqual(dq[0]["reported"], reported)
                self.assertEqual(dq[0]["calculated"], calculated)

    def test_expected_loan_status_still_comes_from_the_soup_fallback(self):
        """The eligibility decision was made on the current-period Bankruptcy
        Date; STAT-01's own derivation is deliberately untouched."""
        for gp in JULY_TARGETS:
            with self.subTest(gp=gp):
                c = self.correction_data[gp]
                self.assertIsNone(c["expected_loan_status_raw"])
                self.assertEqual(c["expected_loan_status_effective"], "BANKRUPTCY")
                self.assertTrue(c["bankruptcy_date_populated"])

    def test_only_bankruptcy_dated_loans_gain_a_loan_status_recommendation(self):
        """Blast radius: the whole file has six rows with a Bankruptcy Date, five
        of which mismatch. Nothing else may be swept in."""
        inputs = self.inputs.copy()
        inputs["_gp"] = inputs["SBA GP Number"].astype(str).str.strip()
        dated = {gp for gp, grp in inputs.groupby("_gp")
                 if any(str(v).strip() for v in grp["Bankruptcy Date"])}
        self.assertEqual(dated, set(JULY_TARGETS) | {JULY_CONTROL_GP})
        ls_gps = {str(r["gp"]).strip() for r in self.recs if r["field"] == "Loan Status"}
        self.assertTrue(ls_gps <= dated)

    def test_approving_every_target_writes_bankruptcy_to_every_posting_row(self):
        decisions = []
        for gp in JULY_TARGETS:
            ls = [r for r in self.by_gp[gp] if r["field"] == "Loan Status"][0]
            decisions.append({"id": ls["id"], "decision": "approve"})
        outcome = corrections.generate_reviewed_workbook(
            open(SRF_PATH, "rb").read(), "july.xlsx", self.display,
            list(self.candidates.values()), decisions,
            "Regression Test", "sha256-real", "2026-07")
        self.assertEqual(outcome["corrections_not_applied"], [])
        self.assertEqual(outcome["applied_count"], len(JULY_TARGETS))
        wb = load_workbook(io.BytesIO(outcome["file_bytes"]), data_only=False)
        ws = wb["Servicer Inputs"]
        header = 9
        cols = {}
        for col in range(1, (ws.max_column or 0) + 1):
            value = ws.cell(row=header, column=col).value
            if value:
                cols[str(value).strip()] = col
        for gp in JULY_TARGETS:
            ls = [r for r in self.by_gp[gp] if r["field"] == "Loan Status"][0]
            for row in ls["rows"]:
                with self.subTest(gp=gp, row=row):
                    self.assertEqual(str(ws.cell(row=row, column=cols["SBA GP Number"]).value), gp)
                    self.assertEqual(ws.cell(row=row, column=cols["Loan Status"]).value,
                                     "BANKRUPTCY")

    def test_approving_loan_status_does_not_touch_delinquency_status(self):
        """Two recommendations exist per target loan. Approving one must leave the
        other's cell exactly as the lender submitted it."""
        gp = "7126119101"
        ls = [r for r in self.by_gp[gp] if r["field"] == "Loan Status"][0]
        outcome = corrections.generate_reviewed_workbook(
            open(SRF_PATH, "rb").read(), "july.xlsx", self.display,
            list(self.candidates.values()), [{"id": ls["id"], "decision": "approve"}],
            "Regression Test", "sha256-real", "2026-07")
        self.assertEqual(outcome["applied_count"], 1)
        ws = load_workbook(io.BytesIO(outcome["file_bytes"]), data_only=False)["Servicer Inputs"]
        cols = {}
        for col in range(1, (ws.max_column or 0) + 1):
            value = ws.cell(row=9, column=col).value
            if value:
                cols[str(value).strip()] = col
        row = ls["rows"][0]
        self.assertEqual(ws.cell(row=row, column=cols["Loan Status"]).value, "BANKRUPTCY")
        self.assertEqual(ws.cell(row=row, column=cols["Delinquency Status"]).value, "150 DAYS")


if __name__ == "__main__":
    unittest.main()
