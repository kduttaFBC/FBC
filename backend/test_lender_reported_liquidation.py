"""Mandatory Human Review of a lender-reported LIQUIDATION.

CLIENT-CONFIRMED (2026-08-14). Every in-scope loan whose CURRENT lender-reported
Loan Status is LIQUIDATION must appear in Human Review as Review Required, and the
CLIENT makes the Loan Status decision. Falcon Bridge does not attempt to decide
whether a liquidation is completed, in progress, or sufficiently supported.

This is a PRESENCE rule, which is what makes it unlike every other Loan Status
branch. It fires regardless of:

    * whether the report contradicts prior SOUP at all;
    * what the prior status is -- LIQUIDATION, ACTIVE, BANKRUPTCY, FORECLOSURE,
      SUBSTITUTION-IN, blank, anything;
    * reversal activity, which otherwise suppresses Loan Status recommendations;
    * WARNING_REASONS, which otherwise exclude the loan from every correction;
    * the Ending Balance, including $0.00 -- the balance does not determine Loan
      Status and there is deliberately NO zero-balance exception.

THE ONE EXCEPTION, unchanged: a loan the prior finalized SOUP already established
as PAID-IN-FULL at $0.00 is CLOSED and is never reopened by a later lender
submission (checks.loans_closed_at_zero). That suppression is upstream and this
branch never sees such a loan.

WHAT IT DOES NOT DO. It never manufactures a proposed status. Where the engine
already has a proposable expected status it is offered as a RECOMMENDATION; where
it does not, corrected_value is None and the reviewer must choose. Never defaulted
to ACTIVE, never to LIQUIDATION. Contradictions against prior status, balances,
terminal dates, payments, bankruptcy / foreclosure / REO information, proceeds,
losses, recoveries, reversals and memo evidence stay visible as EVIDENCE; they are
not additional Loan Status decision rules and are not restated as review reasons.

MEASURED CONTEXT. Closing the bypasses changed no published figure: June remains
127/77/50 and July Lendistry 2025 remains 91/57/34, because every in-scope
LIQUIDATION loan in the available data already surfaced through the older
terminal-contradiction branch. What changed is that it can no longer stop
surfacing.
"""

import io
import os
import re
import unittest

import pandas as pd
from openpyxl import load_workbook

import checks
import corrections


HEADER_ROW = 9
SERVICER = "Test Bank"
GP = "1000000001"
BAL = 80000.0

JSX = os.path.join(os.path.dirname(__file__), "..", "frontend", "src",
                   "components", "CorrectionReview.jsx")


def _soup(rows):
    """rows: (gp, month, loan_status, dq, ending_balance)"""
    return pd.DataFrame([{
        "SBA GP Number": gp, "Payment Month": m, "Loan Status": s,
        "Delinquency Status": dq, "Ending Actual Balance": bal,
        "Beginning Actual Balance": bal, "Deal": "SOUP2025_1",
    } for gp, m, s, dq, bal in rows])


def _row(status, dq, bal, received=0.0, scheduled=500.0, excel_row=HEADER_ROW + 1,
         instance=1, last="Y", **extra):
    row = {
        "SBA GP Number": GP, "Lender Loan Number": "L-1", "Loan Status": status,
        "Delinquency Status": dq, "Payment Received": received, "Principal": 0.0,
        "Net Interest": 0.0, "Servicing Fee": 0.0, "Scheduled Payment": scheduled,
        "Beginning Balance": bal, "Ending Balance": bal,
        "Loan Payment Instance": instance, "Last Payment Instance?": last,
        "_excel_row": excel_row, "_sheet_name": "Servicer Inputs",
        "_header_excel_row": HEADER_ROW,
    }
    row.update(extra)
    return row


def _srf(*rows):
    return pd.DataFrame(list(rows))


def _run(soup_rows, srf, memo=None):
    """Full pipeline -> (stat01, recommendations, candidates, skipped)."""
    results = checks.run_servicer_reconciliation(
        SERVICER, pd.DataFrame(), srf, _soup(soup_rows),
        memo_df=memo, memo_lender_key=SERVICER)
    stat01 = next(r for r in results if r["id"].split("::")[0] == "STAT-01")
    recs, cands, skipped = corrections.build_recommendations(stat01, SERVICER)
    return stat01, recs, cands, skipped


def _ls(recs):
    return [r for r in recs if r["field"] == "Loan Status"]


def _dq(recs):
    return [r for r in recs if r["field"] == "Delinquency Status"]


# ---------------------------------------------------------------------------
class EveryPriorStatusSurfaces(unittest.TestCase):
    """The bypasses. Each of these previously produced NO Loan Status
    recommendation and the loan disappeared from Human Review entirely."""

    def _one(self, prior_status, prior_dq="CURRENT"):
        _, recs, _, _ = _run([(GP, "2026-06-01", prior_status, prior_dq, BAL)],
                             _srf(_row("LIQUIDATION", "CURRENT", BAL)))
        ls = _ls(recs)
        self.assertEqual(len(ls), 1, f"prior {prior_status!r} produced {len(ls)} items")
        self.assertEqual(ls[0]["reported"], "LIQUIDATION")
        self.assertEqual(ls[0]["reviewFlag"],
                         corrections.REVIEW_REASON_LENDER_REPORTED_LIQUIDATION)
        self.assertFalse(ls[0]["standard"])
        self.assertEqual(ls[0]["category"], corrections.REVIEW_CATEGORY_LOAN_STATUS)
        return ls[0]

    def test_prior_active_still_surfaces_and_keeps_its_recommendation(self):
        """The one case that already worked. ACTIVE is proposable, so the audit's
        recommendation is still offered."""
        rec = self._one("ACTIVE")
        self.assertEqual(rec["calculated"], "ACTIVE")

    def test_prior_liquidation_surfaces_although_nothing_contradicts(self):
        """The primary bypass: statuses AGREE, so there is no LOAN STATUS MISMATCH
        and the old branch never ran. A loan reported LIQUIDATION two months
        running must appear in BOTH months."""
        rec = self._one("LIQUIDATION")
        self.assertIsNone(rec["calculated"])

    def test_prior_bankruptcy_surfaces(self):
        self.assertIsNone(self._one("BANKRUPTCY")["calculated"])

    def test_prior_foreclosure_surfaces(self):
        self.assertIsNone(self._one("FORECLOSURE")["calculated"])

    def test_prior_substitution_in_surfaces(self):
        self.assertIsNone(self._one("SUBSTITUTION-IN")["calculated"])

    def test_prior_substitution_out_surfaces(self):
        self.assertIsNone(self._one("SUBSTITUTION-OUT")["calculated"])

    def test_prior_reo_surfaces(self):
        self.assertIsNone(self._one("REO")["calculated"])

    def test_prior_blank_surfaces(self):
        """Expected resolves to UNVERIFIED, which is not proposable."""
        self.assertIsNone(self._one("")["calculated"])

    def test_prior_deferred_surfaces(self):
        self._one("DEFERRED")

    def test_the_fully_agreeing_case_surfaces(self):
        """The case that motivated this rule: prior LIQUIDATION, delinquency
        agrees, no reversal, no warning, nothing mismatched at all. STAT-01 carries
        only TERMINAL STATUS REVIEW, and previously nothing reached the reviewer."""
        stat01, recs, _, _ = _run(
            [(GP, "2026-05-01", "LIQUIDATION", "CURRENT", BAL),
             (GP, "2026-06-01", "LIQUIDATION", "CURRENT", BAL)],
            _srf(_row("LIQUIDATION", "CURRENT", BAL, received=600.0)))
        entry = next(c for c in stat01["correction_data"] if str(c["gp"]) == GP)
        self.assertTrue(entry["status_match"])
        self.assertIs(entry["dq_match"], True)
        self.assertEqual(sorted(set(entry["reasons"])), ["TERMINAL STATUS REVIEW"])
        self.assertEqual(len(_ls(recs)), 1)
        self.assertEqual(_dq(recs), [])          # nothing else to correct


# ---------------------------------------------------------------------------
class ZeroEndingBalanceIsStillInScope(unittest.TestCase):
    """CLIENT-CONFIRMED (2026-08-14): the balance does not determine Loan Status,
    and there is deliberately NO zero-balance exception. Only a prior finalized
    PAID-IN-FULL at $0.00 closes a loan."""

    def test_zero_ending_balance_liquidation_surfaces(self):
        _, recs, _, _ = _run([(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
                             _srf(_row("LIQUIDATION", "CURRENT", 0.0)))
        ls = _ls(recs)
        self.assertEqual(len(ls), 1)
        self.assertEqual(ls[0]["reviewFlag"],
                         corrections.REVIEW_REASON_LENDER_REPORTED_LIQUIDATION)

    def test_zero_balance_with_a_liquidation_date_surfaces(self):
        _, recs, _, _ = _run(
            [(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
            _srf(_row("LIQUIDATION", "CURRENT", 0.0, **{"Liquidation Date": "2026-07-20"})))
        self.assertEqual(len(_ls(recs)), 1)

    def test_a_zero_balance_liquidation_is_not_confused_with_a_closed_loan(self):
        """A CURRENT zero balance is not prior finalized closure. The distinction
        is the whole of the exception."""
        _, recs, _, _ = _run([(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
                             _srf(_row("LIQUIDATION", "CURRENT", 0.0)))
        self.assertEqual(len(_ls(recs)), 1)


# ---------------------------------------------------------------------------
class ReversalPrecedence(unittest.TestCase):
    """LIQUIDATION takes precedence over the 2026-08-03 reversal suppression, for
    LOAN STATUS ONLY. Everything else keeps the existing reversal treatment."""

    def _reversal_srf(self, status, first, second):
        return _srf(_row(status, "CURRENT", BAL, received=first),
                    _row(status, "CURRENT", BAL, received=second,
                         excel_row=HEADER_ROW + 2, instance=2, last=""))

    def test_net_negative_reversal_no_longer_suppresses_a_liquidation(self):
        _, recs, _, _ = _run([(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
                             self._reversal_srf("LIQUIDATION", 0.0, -900.0))
        ls = _ls(recs)
        self.assertEqual(len(ls), 1)
        self.assertEqual(ls[0]["reviewFlag"],
                         corrections.REVIEW_REASON_LENDER_REPORTED_LIQUIDATION)

    def test_offsetting_zero_net_reversal_no_longer_suppresses_a_liquidation(self):
        _, recs, _, _ = _run([(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
                             self._reversal_srf("LIQUIDATION", 500.0, -500.0))
        self.assertEqual(len(_ls(recs)), 1)

    def test_a_non_liquidation_reversal_still_gets_no_loan_status_recommendation(self):
        """The 2026-08-03 rule is UNCHANGED outside LIQUIDATION."""
        _, recs, _, _ = _run([(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
                             self._reversal_srf("ACTIVE", 500.0, -900.0))
        self.assertEqual(_ls(recs), [])

    def test_a_non_liquidation_reversal_still_reaches_review_via_delinquency(self):
        _, recs, _, _ = _run([(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
                             self._reversal_srf("ACTIVE", 500.0, -900.0))
        dq = _dq(recs)
        self.assertEqual(len(dq), 1)
        self.assertFalse(dq[0]["standard"])
        self.assertEqual(dq[0]["category"], corrections.REVIEW_CATEGORY_REVERSAL)

    def test_delinquency_processing_for_a_liquidation_keeps_reversal_treatment(self):
        """Precedence is scoped to Loan Status. The DQ recommendation on the same
        loan still carries the reversal reason and is still Review Required."""
        _, recs, _, _ = _run([(GP, "2026-06-01", "ACTIVE", "30 DAYS", BAL)],
                             self._reversal_srf("LIQUIDATION", 0.0, -900.0))
        dq = _dq(recs)
        if dq:                                   # only if a DQ mismatch arose
            self.assertFalse(dq[0]["standard"])
            self.assertEqual(dq[0]["category"], corrections.REVIEW_CATEGORY_REVERSAL)


# ---------------------------------------------------------------------------
class WarningReasonsScope(unittest.TestCase):
    """The relocated WARNING_REASONS gate. It no longer suppresses the mandatory
    LIQUIDATION item and still suppresses everything else, unchanged.

    Driven through build_eligible_corrections with `reasons` injected directly,
    because that is precisely the input the gate reads -- deriving a warning
    through STAT-01 would test the derivation, not the gate."""

    BASE = ["DELINQUENCY STATUS MISMATCH", "LOAN STATUS MISMATCH",
            "TERMINAL STATUS REVIEW"]

    def _entry(self, status, reasons, **over):
        e = {
            "gp": GP, "excel_rows": [HEADER_ROW + 1], "skipped_row_refs": [],
            "sheet_name": "Servicer Inputs", "header_excel_row": HEADER_ROW,
            "reasons": list(reasons), "dq_match": False, "calculated_dq": "30 DAYS",
            "srf_dq": "CURRENT", "srf_status": status, "soup_status": "ACTIVE",
            "expected_loan_status_effective": "ACTIVE",
            "expected_loan_status_raw": None, "ending_balance": BAL,
            "posting_count": 1, "has_negative_posting": False,
            "offsetting_zero_net_reversal": False, "reversal_equivalents": None,
        }
        e.update(over)
        return e

    def _fields(self, status, reasons, **over):
        out = corrections.build_eligible_corrections(
            [self._entry(status, reasons, **over)])
        return [e["field"] for e in out]

    def test_every_warning_still_lets_a_liquidation_through(self):
        for w in sorted(corrections.WARNING_REASONS):
            with self.subTest(warning=w):
                self.assertEqual(
                    self._fields("LIQUIDATION", self.BASE + [w]), ["Loan Status"])

    def test_every_warning_still_suppresses_the_delinquency_correction(self):
        for w in sorted(corrections.WARNING_REASONS):
            with self.subTest(warning=w):
                self.assertNotIn(
                    "Delinquency Status", self._fields("LIQUIDATION", self.BASE + [w]))

    def test_every_warning_still_suppresses_a_non_liquidation_loan_entirely(self):
        for w in sorted(corrections.WARNING_REASONS):
            with self.subTest(warning=w):
                self.assertEqual(
                    self._fields("ACTIVE", ["DELINQUENCY STATUS MISMATCH", w]), [])

    def test_every_warning_still_suppresses_foreclosure_entirely(self):
        """FORECLOSURE is a collection status but NOT LIQUIDATION, so the bypass
        must not reach it."""
        for w in sorted(corrections.WARNING_REASONS):
            with self.subTest(warning=w):
                self.assertEqual(self._fields("FORECLOSURE", self.BASE + [w]), [])

    def test_without_a_warning_the_delinquency_correction_is_unaffected(self):
        self.assertEqual(sorted(self._fields("LIQUIDATION", self.BASE)),
                         ["Delinquency Status", "Loan Status"])

    def test_the_warning_set_itself_is_unchanged(self):
        """A membership pin, so the LIQUIDATION bypass above is re-examined
        whenever the warning set grows.

        UPDATED with the row-scoped FORM-01 quarantine (see
        test_form01_row_quarantine.py), which deliberately adds a fourth
        member: a loan whose SOURCE workbook cells were overwritten with
        pasted literals must not become a correction candidate. The four
        iterating tests above already assert the new member suppresses
        corrections exactly like the other three -- including that it is NOT
        reached by the LIQUIDATION bypass -- so the guard this pin provides
        is intact rather than merely re-baselined."""
        self.assertEqual(corrections.WARNING_REASONS,
                         {"MEMO DATA QUALITY WARNING", "DEFERRAL DATA QUALITY WARNING",
                          "SCHEDULED PAYMENT AMBIGUOUS",
                          "SOURCE FORMULA INTEGRITY DEFECT"})


# ---------------------------------------------------------------------------
class OneItemNotTwo(unittest.TestCase):
    def test_a_liquidation_that_also_contradicts_yields_exactly_one(self):
        """Both the new branch and the terminal-contradiction branch would fire.
        The _emitted_fields guard is what keeps it to one -- two would share a
        single id and generate_reviewed_workbook would silently keep only one."""
        _, recs, cands, _ = _run([(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
                                 _srf(_row("LIQUIDATION", "CURRENT", BAL)))
        ls = _ls(recs)
        self.assertEqual(len(ls), 1)
        self.assertEqual(len({r["id"] for r in recs}), len(recs))
        self.assertEqual(len([c for c in cands if c["field"] == "Loan Status"]), 1)

    def test_the_single_item_carries_the_liquidation_reason_not_the_terminal_one(self):
        _, recs, _, _ = _run([(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
                             _srf(_row("LIQUIDATION", "CURRENT", BAL)))
        flag = _ls(recs)[0]["reviewFlag"]
        self.assertEqual(flag, corrections.REVIEW_REASON_LENDER_REPORTED_LIQUIDATION)
        self.assertNotIn(corrections.REVIEW_REASON_TERMINAL_STATUS, flag)

    def test_foreclosure_and_reo_keep_the_terminal_contradiction_branch(self):
        for status in ("FORECLOSURE", "REO"):
            with self.subTest(status=status):
                _, recs, _, _ = _run([(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
                                     _srf(_row(status, "CURRENT", BAL)))
                ls = _ls(recs)
                self.assertEqual(len(ls), 1)
                self.assertEqual(ls[0]["reviewFlag"],
                                 corrections.REVIEW_REASON_TERMINAL_STATUS)
                self.assertEqual(ls[0]["calculated"], "ACTIVE")


# ---------------------------------------------------------------------------
class NoManufacturedValue(unittest.TestCase):
    def test_no_proposable_status_means_no_proposed_value(self):
        for prior in ("LIQUIDATION", "BANKRUPTCY", "FORECLOSURE", "REO",
                      "SUBSTITUTION-IN", "SUBSTITUTION-OUT", ""):
            with self.subTest(prior=prior):
                _, recs, _, _ = _run([(GP, "2026-06-01", prior, "CURRENT", BAL)],
                                     _srf(_row("LIQUIDATION", "CURRENT", BAL)))
                rec = _ls(recs)[0]
                self.assertIsNone(rec["calculated"])
                self.assertNotEqual(rec["calculated"], "ACTIVE")
                self.assertNotEqual(rec["calculated"], "LIQUIDATION")

    def test_a_proposable_status_is_offered_as_a_recommendation(self):
        _, recs, _, _ = _run([(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
                             _srf(_row("LIQUIDATION", "CURRENT", BAL)))
        self.assertEqual(_ls(recs)[0]["calculated"], "ACTIVE")

    def test_the_application_still_cannot_generate_liquidation(self):
        ok, reason = corrections.validate_approved_value(
            "Loan Status", "LIQUIDATION", substituted=False)
        self.assertFalse(ok)
        self.assertIn("may generate", reason)

    def test_a_blank_approved_value_is_refused(self):
        """This is what makes "the reviewer must choose" enforceable server-side:
        approving without selecting cannot be applied."""
        for blank in (None, "", "   "):
            with self.subTest(value=blank):
                ok, reason = corrections.validate_approved_value(
                    "Loan Status", blank, substituted=False)
                self.assertFalse(ok)
                self.assertIn("blank", reason.lower())


# ---------------------------------------------------------------------------
class ReviewerMaySelectAnyPermittedStatus(unittest.TestCase):
    def test_all_nine_validate_as_a_reviewer_selection(self):
        self.assertEqual(len(corrections._PERMITTED_LOAN_STATUS), 9)
        for status in sorted(corrections._PERMITTED_LOAN_STATUS):
            with self.subTest(status=status):
                ok, value = corrections.validate_approved_value(
                    "Loan Status", status, substituted=True)
                self.assertTrue(ok)
                self.assertEqual(value, status)

    def test_liquidation_is_a_valid_explicit_reviewer_selection(self):
        ok, value = corrections.validate_approved_value(
            "Loan Status", "LIQUIDATION", substituted=True)
        self.assertTrue(ok)
        self.assertEqual(value, "LIQUIDATION")

    def test_cross_domain_and_invented_values_are_refused(self):
        for bad in ("CURRENT", "60 DAYS", "CHARGE-OFF", "LIQUIDATED", "", "  "):
            with self.subTest(value=bad):
                ok, _ = corrections.validate_approved_value(
                    "Loan Status", bad, substituted=True)
                self.assertFalse(ok)


# ---------------------------------------------------------------------------
class PriorPaidInFullAtZeroStaysClosed(unittest.TestCase):
    """The ONE exception, unchanged. A lender cannot reopen a finalized closure."""

    CLOSED = [(GP, "2026-05-01", "ACTIVE", "CURRENT", BAL),
              (GP, "2026-06-01", "PAID-IN-FULL", "CURRENT", 0.0)]

    def test_a_later_liquidation_report_produces_nothing(self):
        stat01, recs, _, _ = _run(self.CLOSED, _srf(_row("LIQUIDATION", "CURRENT", BAL)))
        self.assertEqual(_ls(recs), [])
        self.assertEqual(recs, [])
        trace = stat01.get("suppressed_closed_loans") or []
        self.assertEqual([t["SBA GP Number"] for t in trace], [GP])

    def test_the_suppression_survives_a_zero_balance_liquidation_report(self):
        _, recs, _, _ = _run(self.CLOSED, _srf(_row("LIQUIDATION", "CURRENT", 0.0)))
        self.assertEqual(recs, [])

    def test_the_suppression_survives_a_reversal(self):
        _, recs, _, _ = _run(self.CLOSED, _srf(
            _row("LIQUIDATION", "CURRENT", BAL, received=500.0),
            _row("LIQUIDATION", "CURRENT", BAL, received=-500.0,
                 excel_row=HEADER_ROW + 2, instance=2, last="")))
        self.assertEqual(recs, [])

    def test_a_prior_paid_in_full_at_a_NONZERO_balance_is_not_closed(self):
        """Both halves of the exception are required."""
        _, recs, _, _ = _run([(GP, "2026-06-01", "PAID-IN-FULL", "CURRENT", 4000.0)],
                             _srf(_row("LIQUIDATION", "CURRENT", BAL)))
        self.assertEqual(len(_ls(recs)), 1)

    def test_a_prior_zero_balance_while_active_is_not_closed(self):
        _, recs, _, _ = _run([(GP, "2026-06-01", "ACTIVE", "CURRENT", 0.0)],
                             _srf(_row("LIQUIDATION", "CURRENT", BAL)))
        self.assertEqual(len(_ls(recs)), 1)


# ---------------------------------------------------------------------------
class ReviewerDecisionWritesThrough(unittest.TestCase):
    """The reviewer's selection becomes the authoritative audited Loan Status and
    reaches every downstream surface. Uses a real workbook."""

    SHEET = "Servicer Inputs"
    HEADERS = ["SBA GP Number", "Lender Loan Number", "Loan Status",
               "Delinquency Status", "Payment Received", "Servicing Fee",
               "Scheduled Payment", "Beginning Balance", "Ending Balance",
               "Loan Payment Instance", "Last Payment Instance?"]

    def _workbook(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = self.SHEET
        for i, h in enumerate(self.HEADERS, start=1):
            ws.cell(row=HEADER_ROW, column=i, value=h)
        vals = [GP, "L-1", "LIQUIDATION", "CURRENT", 0.0, 0.0, 0.0, BAL, BAL, 1, "Y"]
        for i, v in enumerate(vals, start=1):
            ws.cell(row=HEADER_ROW + 1, column=i, value=v)
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def _candidate(self, calculated):
        return [{
            "id": f"{SERVICER}||{GP}||Loan Status", "gp": GP, "field": "Loan Status",
            "corrected_value": calculated, "reported_value": "LIQUIDATION",
            "excel_rows": [HEADER_ROW + 1], "skipped_row_refs": [],
            "sheet_name": self.SHEET, "header_excel_row": HEADER_ROW,
            "eligibility_rule": "Lender-reported LIQUIDATION requires a human Loan Status decision",
            "review_required": True,
            "review_reason": corrections.REVIEW_REASON_LENDER_REPORTED_LIQUIDATION,
            "evidence_marker": None,
        }]

    def _review(self, decisions, calculated=None):
        return corrections.generate_reviewed_workbook(
            self._workbook(), "T.xlsx", SERVICER, self._candidate(calculated),
            decisions, "Investigator", "sha", pd.Timestamp("2026-07-01"))

    def _log(self, file_bytes):
        wb = load_workbook(io.BytesIO(file_bytes))
        ws = wb[corrections.CORRECTION_LOG_SHEET_NAME]
        head = [c.value for c in ws[1]]
        return [dict(zip(head, [c.value for c in r])) for r in ws.iter_rows(min_row=2)]

    def test_an_explicit_liquidation_selection_is_written_and_logged(self):
        out = self._review([{"id": f"{SERVICER}||{GP}||Loan Status",
                             "decision": "approve", "approved_value": "LIQUIDATION"}])
        self.assertEqual(out["applied_count"], 1)
        ws = load_workbook(io.BytesIO(out["file_bytes"]))[self.SHEET]
        self.assertEqual(ws.cell(row=HEADER_ROW + 1, column=3).value, "LIQUIDATION")
        row = self._log(out["file_bytes"])[0]
        self.assertEqual(row["Original Value"], "LIQUIDATION")
        self.assertIsNone(row["Calculated Value"])       # nothing was proposed
        self.assertEqual(row["Reviewer Approved Value"], "LIQUIDATION")
        self.assertEqual(row["Corrected Value"], "LIQUIDATION")
        self.assertEqual(out["verified_overrides"], {(GP, "Loan Status"): "LIQUIDATION"})

    def test_a_different_selection_overrides_the_lender(self):
        out = self._review([{"id": f"{SERVICER}||{GP}||Loan Status",
                             "decision": "approve", "approved_value": "BANKRUPTCY"}])
        ws = load_workbook(io.BytesIO(out["file_bytes"]))[self.SHEET]
        self.assertEqual(ws.cell(row=HEADER_ROW + 1, column=3).value, "BANKRUPTCY")
        self.assertEqual(out["verified_overrides"], {(GP, "Loan Status"): "BANKRUPTCY"})

    def test_every_permitted_status_can_be_written(self):
        for status in sorted(corrections._PERMITTED_LOAN_STATUS):
            with self.subTest(status=status):
                out = self._review([{"id": f"{SERVICER}||{GP}||Loan Status",
                                     "decision": "approve", "approved_value": status}])
                self.assertEqual(out["applied_count"], 1, f"{status} not written")
                self.assertEqual(out["verified_overrides"],
                                 {(GP, "Loan Status"): status})

    def test_approving_without_selecting_cannot_be_applied(self):
        out = self._review([{"id": f"{SERVICER}||{GP}||Loan Status",
                             "decision": "approve"}])
        self.assertEqual(out["applied_count"], 0)
        self.assertEqual(out["approved_not_applied_count"], 1)
        self.assertEqual(out["verified_overrides"], {})
        ws = load_workbook(io.BytesIO(out["file_bytes"]))[self.SHEET]
        self.assertEqual(ws.cell(row=HEADER_ROW + 1, column=3).value, "LIQUIDATION")

    def test_rejecting_retains_liquidation_and_is_a_valid_final_decision(self):
        out = self._review([{"id": f"{SERVICER}||{GP}||Loan Status", "decision": "reject"}])
        self.assertEqual(out["rejected_count"], 1)
        self.assertEqual(out["verified_overrides"], {})
        row = self._log(out["file_bytes"])[0]
        self.assertEqual(row["Rows Updated"], 0)
        self.assertIn("Rejected", str(row["Decision"]))

    def test_the_selection_reaches_the_final_audited_summary_and_the_msr(self):
        import monthly_servicer_report as msr
        si = pd.DataFrame([_row("LIQUIDATION", "CURRENT", BAL)])
        override = {(GP, "Loan Status"): "ACTIVE"}
        lender = checks.verified_portfolio_summary_views(si, {}, SERVICER, None)
        reviewed = checks.verified_portfolio_summary_views(si, override, SERVICER, None)
        self.assertEqual(lender["counts"].get("Liquidation", 0), 1)
        self.assertEqual(reviewed["counts"].get("Liquidation", 0), 0)
        self.assertAlmostEqual(reviewed["dollars"]["Outstanding ($)"], BAL, places=2)

        def status_of(rows):
            for r in rows:
                for k, v in r.items():
                    if "Loan Status" in str(k):
                        return v
            return None
        si2 = si.assign(**{"Principal": 0.0, "Net Interest": 0.0})
        self.assertEqual(status_of(msr.build_servicer_rows(si2, SERVICER, {})), "LIQUIDATION")
        self.assertEqual(status_of(msr.build_servicer_rows(si2, SERVICER, override)), "ACTIVE")


# ---------------------------------------------------------------------------
class D4GatesTheMandatoryReview(unittest.TestCase):
    """The item is an ordinary candidate, so D4's existing arithmetic gates it."""

    def test_it_is_counted_and_blocks_while_pending(self):
        _, recs, cands, _ = _run([(GP, "2026-06-01", "LIQUIDATION", "CURRENT", BAL)],
                                 _srf(_row("LIQUIDATION", "CURRENT", BAL)))
        ls = _ls(recs)
        self.assertEqual(len(ls), 1)
        # present in the CANDIDATE set, which is D4's denominator
        self.assertIn(ls[0]["id"], [c["id"] for c in cands])

    def test_return_is_a_non_reportable_disposition(self):
        self.assertIn("return", corrections.VALID_DECISIONS)
        self.assertEqual(corrections.VALID_DECISIONS,
                         frozenset(("approve", "reject", "return")))

    def test_it_is_never_bulk_approvable(self):
        _, recs, _, _ = _run([(GP, "2026-06-01", "LIQUIDATION", "CURRENT", BAL)],
                             _srf(_row("LIQUIDATION", "CURRENT", BAL)))
        rec = _ls(recs)[0]
        self.assertFalse(rec["standard"])
        self.assertEqual(rec["category"], corrections.REVIEW_CATEGORY_LOAN_STATUS)
        cat, std = corrections.classify_recommendation(
            {"field": "Loan Status", "review_required": True,
             "review_reason": corrections.REVIEW_REASON_LENDER_REPORTED_LIQUIDATION,
             "reversal_equivalents": None, "offsetting_zero_net_reversal": False,
             "has_negative_posting": False})
        self.assertFalse(std)
        self.assertEqual(cat, corrections.REVIEW_CATEGORY_LOAN_STATUS)


# ---------------------------------------------------------------------------
class EvidenceRemainsEvidence(unittest.TestCase):
    """Contradictions stay visible as evidence and are NOT restated as reasons."""

    def test_the_reason_is_exactly_one_string_with_nothing_appended(self):
        _, recs, _, _ = _run(
            [(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
            _srf(_row("LIQUIDATION", "CURRENT", BAL,
                      **{"Liquidation Date": "2026-07-20",
                         "Bankruptcy Date": "2026-03-12",
                         "Foreclosure Date": "2026-02-01", "REO Date": "2026-01-15"})))
        flag = _ls(recs)[0]["reviewFlag"]
        self.assertEqual(flag, corrections.REVIEW_REASON_LENDER_REPORTED_LIQUIDATION)
        for other in (corrections.REVIEW_REASON_TERMINAL_STATUS,
                      corrections.REVIEW_REASON_CROSS_DOMAIN,
                      corrections.REVIEW_REASON_REVERSAL,
                      corrections.REVIEW_REASON_BANKRUPTCY_CURRENT_PERIOD):
            self.assertNotIn(other, flag)

    def test_the_evidence_fields_are_all_carried(self):
        _, recs, _, _ = _run(
            [(GP, "2026-06-01", "ACTIVE", "CURRENT", BAL)],
            _srf(_row("LIQUIDATION", "CURRENT", BAL,
                      **{"Liquidation Date": "2026-07-20",
                         "Liquidation Proceeds": 1000.0, "Liquidation Loss": 250.0,
                         "Subsequent Recoveries/(Losses)": -10.0,
                         "Foreclosure Date": "2026-02-01", "REO Date": "2026-01-15",
                         "Bankruptcy Date": "2026-03-12"})))
        rec = _ls(recs)[0]
        self.assertEqual(rec["reported"], "LIQUIDATION")
        self.assertEqual(rec["servicerLoanStatus"], "LIQUIDATION")
        self.assertEqual(rec["priorLoanStatus"], "ACTIVE")
        for key in ("priorEndingBalance", "balance", "paymentReceived",
                    "scheduledPayment", "liquidationDate", "liquidationProceeds",
                    "liquidationLoss", "subsequentRecoveries", "foreclosureDate",
                    "reoDate", "bankruptcyDate", "evidenceStatus",
                    "hasNegativePosting", "postings"):
            self.assertIn(key, rec)
        self.assertEqual(rec["liquidationProceeds"], 1000.0)
        self.assertEqual(rec["liquidationLoss"], 250.0)
        self.assertEqual(rec["subsequentRecoveries"], -10.0)


# ---------------------------------------------------------------------------
class FrontendContract(unittest.TestCase):
    """The review page must offer a placeholder and require a choice."""

    @classmethod
    def setUpClass(cls):
        with open(JSX, encoding="utf-8") as fh:
            cls.src = fh.read()

    def test_the_placeholder_exists_and_is_not_a_status(self):
        self.assertIn("— Select Loan Status —", self.src)
        self.assertIn('<option value="">— Select Loan Status —</option>', self.src)

    def test_the_placeholder_only_appears_without_a_recommendation(self):
        self.assertIn("{!rec.calculated && <option", self.src)

    def test_the_select_stays_controlled_when_there_is_no_recommendation(self):
        self.assertIn("d.approved_value || rec.calculated || ''", self.src)

    def test_the_domain_is_still_exactly_the_nine_permitted_values(self):
        dom = re.findall(r"'([^']+)'", re.search(
            r"const LOAN_STATUS_DOMAIN = \[(.*?)\]", self.src, re.S).group(1))
        self.assertEqual(set(dom), corrections._PERMITTED_LOAN_STATUS)
        self.assertEqual(len(dom), 9)

    def test_no_free_text_loan_status_input(self):
        self.assertNotRegex(
            self.src, r"type=\"text\"[^>]*approved_value|approved_value[^)]*type=\"text\"")

    def test_the_reviewer_instruction_is_present(self):
        self.assertIn("The lender reported this loan as LIQUIDATION.", self.src)
        self.assertIn("select the Loan Status you believe is appropriate", self.src)

    def test_the_jsx_reason_constant_matches_the_backend(self):
        m = re.search(r"const LENDER_LIQUIDATION_REASON = '([^']+)'", self.src)
        self.assertIsNotNone(m)
        self.assertTrue(
            corrections.REVIEW_REASON_LENDER_REPORTED_LIQUIDATION.startswith(m.group(1)))

    def test_an_approval_without_a_selection_blocks_submission(self):
        self.assertIn("approvalsMissingValue", self.src)
        self.assertIn("|| approvalsMissingValue.length > 0", self.src)


if __name__ == "__main__":
    unittest.main()
