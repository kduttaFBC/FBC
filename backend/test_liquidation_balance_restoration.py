"""LIQUIDATION -> ACTIVE requires an explicit reviewer balance decision.

THE DEFECT THIS CLOSES (June 2026, Lendistry SOUP 2025-1)

A reviewer approved Loan Status LIQUIDATION -> ACTIVE on two loans. The override
mechanism maps only Loan Status and Delinquency Status, so each loan kept the
lender's $0.00 Ending Balance. Both entered Outstanding by COUNT (1,228 -> 1,230)
but contributed $0, understating the Portfolio Summary Outstanding Balance by
exactly the two beginning balances, while a Liquidation Date and Liquidation Loss
survived beside a status of ACTIVE.

THE RULE (client-approved 2026-08-15)

When, and ONLY when, all three hold:

    reported Loan Status is LIQUIDATION
    the approved Loan Status is ACTIVE
    current-month Beginning Balance > 0

a SECOND, REQUIRED reviewer decision is raised with exactly two permitted
choices -- restore Ending Balance to the current-month Beginning Balance, or keep
it at $0.00. There is no free-form balance entry. Prior Month Ending Balance is
SUPPORTING EVIDENCE ONLY and is never selectable, even when it differs from the
Beginning Balance; a difference is displayed and warned about, and neither option
is pre-selected.

Restore, and only restore, also clears Liquidation Date and Liquidation Loss.
Keep $0.00 preserves them, and the resulting contradiction stays visible.

SCOPE. This must never become general reviewer editing of Ending Balance:
NoGeneralBalanceEditingTests below is the guard on that.
"""

import io
import os
import unittest

import pandas as pd
from openpyxl import Workbook, load_workbook

import checks
import corrections


HEADER_ROW = 9
SERVICER = "Test Bank"
SHEET = "Servicer Inputs"

GP_LIQ = "1000000001"      # LIQUIDATION, positive beginning balance
GP_OTHER = "2000000002"    # ordinary ACTIVE loan, must never be affected

BEGIN = 20088.08
LIQ_DATE = "2026-06-30"
LIQ_LOSS = 20088.08

SRF_COLUMNS = [
    "SBA GP Number", "Lender Loan Number", "Loan Status", "Delinquency Status",
    "Net Interest Rate", "Net Interest", "Principal", "Payment Received",
    "Servicing Fee", "Scheduled Payment", "Beginning Balance", "Ending Balance",
    "Loan Payment Instance", "Last Payment Instance?",
    "Liquidation Date", "Liquidation Loss", "Liquidation Proceeds",
    "Subsequent Recoveries/(Losses)", "Paid-In-Full Date", "Foreclosure Date",
    "REO Date", "Bankruptcy Date",
]


# ---------------------------------------------------------------------------
# fixtures
# ---------------------------------------------------------------------------

def _soup(rows):
    """rows: (gp, month, loan_status, dq, ending_balance)"""
    return pd.DataFrame([{
        "SBA GP Number": gp, "Payment Month": m, "Loan Status": s,
        "Delinquency Status": dq, "Ending Actual Balance": bal,
        "Beginning Actual Balance": bal, "Deal": "SOUP2025_1",
    } for gp, m, s, dq, bal in rows])


def _row(gp, status="ACTIVE", dq="CURRENT", beginning=BEGIN, ending=BEGIN,
         excel_row=HEADER_ROW + 1, instance=1, last="Y",
         liquidation_date=None, liquidation_loss=None, received=0.0,
         scheduled=500.0):
    return {
        "SBA GP Number": gp, "Lender Loan Number": "L-" + gp,
        "Loan Status": status, "Delinquency Status": dq,
        "Net Interest Rate": 0.075, "Net Interest": 0.0, "Principal": 0.0,
        "Payment Received": received, "Servicing Fee": 0.0,
        "Scheduled Payment": scheduled,
        "Beginning Balance": beginning, "Ending Balance": ending,
        "Loan Payment Instance": instance, "Last Payment Instance?": last,
        "Liquidation Date": liquidation_date, "Liquidation Loss": liquidation_loss,
        "Liquidation Proceeds": None, "Subsequent Recoveries/(Losses)": None,
        "Paid-In-Full Date": None, "Foreclosure Date": None, "REO Date": None,
        "Bankruptcy Date": None,
        "_excel_row": excel_row, "_sheet_name": SHEET,
        "_header_excel_row": HEADER_ROW,
    }


def _liquidated_row(gp=GP_LIQ, beginning=BEGIN, excel_row=HEADER_ROW + 1,
                    instance=1, last="Y", ending=0.0,
                    liquidation_date=LIQ_DATE, liquidation_loss=LIQ_LOSS):
    return _row(gp, status="LIQUIDATION", dq="90 DAYS", beginning=beginning,
                ending=ending, excel_row=excel_row, instance=instance, last=last,
                liquidation_date=liquidation_date, liquidation_loss=liquidation_loss)


def _srf(*rows):
    return pd.DataFrame(list(rows))


def _prior_soup(gp=GP_LIQ, prior_ending=BEGIN, status="ACTIVE"):
    return [(gp, "2026-05-01", status, "CURRENT", prior_ending)]


def _run(srf, soup_rows):
    """Full pipeline -> (stat01, recommendations, candidates)."""
    results = checks.run_servicer_reconciliation(
        SERVICER, pd.DataFrame(), srf, _soup(soup_rows),
        memo_df=None, memo_lender_key=SERVICER)
    stat01 = next(r for r in results if r["id"].split("::")[0] == "STAT-01")
    recs, cands, _skipped = corrections.build_recommendations(stat01, SERVICER)
    return stat01, recs, cands


def _of_field(recs, field):
    return [r for r in recs if r["field"] == field]


def _balance_recs(recs):
    return _of_field(recs, corrections.BALANCE_DECISION_FIELD)


def _workbook_bytes(rows):
    """A minimal SRF workbook _apply_one_correction can locate columns in."""
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    for offset, name in enumerate(SRF_COLUMNS, start=1):
        ws.cell(row=HEADER_ROW, column=offset, value=name)
    for index, row in enumerate(rows):
        excel_row = row.get("_excel_row") or HEADER_ROW + 1 + index
        for offset, name in enumerate(SRF_COLUMNS, start=1):
            ws.cell(row=excel_row, column=offset, value=row.get(name))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _review(srf_rows, soup_rows, decisions_for):
    """Runs the pipeline, then generate_reviewed_workbook with the decisions
    decisions_for(recs) returns. -> (result, recs, workbook)."""
    srf = _srf(*srf_rows)
    _stat01, recs, cands = _run(srf, soup_rows)
    result = corrections.generate_reviewed_workbook(
        _workbook_bytes(srf_rows), "test.xlsx", SERVICER, cands,
        decisions_for(recs), "tester", "sha-256-test", None)
    wb = load_workbook(io.BytesIO(result["file_bytes"]), data_only=False)
    return result, recs, wb


def _approve_status_and_balance(balance_value):
    def build(recs):
        out = []
        for rec in recs:
            if rec["field"] == "Loan Status":
                out.append({"id": rec["id"], "decision": "approve",
                            "approved_value": "ACTIVE"})
            elif rec["field"] == corrections.BALANCE_DECISION_FIELD:
                out.append({"id": rec["id"], "decision": "approve",
                            "approved_value": balance_value})
        return out
    return build


def _cell(wb, excel_row, column_name):
    ws = wb[SHEET]
    for c in range(1, ws.max_column + 1):
        if ws.cell(row=HEADER_ROW, column=c).value == column_name:
            return ws.cell(row=excel_row, column=c).value
    raise AssertionError("column %r not found" % column_name)


def _log_rows(wb):
    """The Correction Log sheet AS WRITTEN, one dict per row keyed by header.

    Read from the artifact rather than from an in-memory list: the sheet is what
    an auditor actually receives, so that is what these tests assert against."""
    ws = wb[corrections.CORRECTION_LOG_SHEET_NAME]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        record = {headers[c - 1]: ws.cell(row=r, column=c).value
                  for c in range(1, ws.max_column + 1)}
        if record.get("SBA GP Number") in (None, ""):
            continue
        rows.append(record)
    return rows


def _log_by_field(wb):
    out = {}
    for row in _log_rows(wb):
        out.setdefault(row.get("Field Updated"), []).append(row)
    return out


# ---------------------------------------------------------------------------
# 1 -- when the balance decision exists, and when it must not
# ---------------------------------------------------------------------------

class BalanceDecisionEmissionTests(unittest.TestCase):

    def test_positive_beginning_balance_raises_the_decision(self):
        _s, recs, _c = _run(_srf(_liquidated_row()), _prior_soup())
        balance = _balance_recs(recs)
        self.assertEqual(len(balance), 1)
        self.assertEqual(str(balance[0]["gp"]), GP_LIQ)
        self.assertTrue(balance[0]["reviewFlag"])
        self.assertIn("BALANCE", balance[0]["reviewFlag"].upper())

    def test_it_is_required_review_never_standard(self):
        _s, recs, _c = _run(_srf(_liquidated_row()), _prior_soup())
        rec = _balance_recs(recs)[0]
        self.assertFalse(rec["standard"])
        self.assertTrue(rec["requiresBalanceDecision"])

    def test_zero_beginning_balance_raises_no_decision(self):
        _s, recs, _c = _run(_srf(_liquidated_row(beginning=0.0)), _prior_soup())
        self.assertEqual(_balance_recs(recs), [])

    def test_blank_beginning_balance_raises_no_decision(self):
        _s, recs, _c = _run(_srf(_liquidated_row(beginning=None)), _prior_soup())
        self.assertEqual(_balance_recs(recs), [])

    def test_negative_beginning_balance_raises_no_decision(self):
        _s, recs, _c = _run(_srf(_liquidated_row(beginning=-5.0)), _prior_soup())
        self.assertEqual(_balance_recs(recs), [])

    def test_it_carries_no_default_so_the_reviewer_must_choose(self):
        _s, recs, _c = _run(_srf(_liquidated_row()), _prior_soup())
        self.assertIsNone(_balance_recs(recs)[0]["calculated"])

    def test_it_offers_exactly_two_choices(self):
        _s, recs, _c = _run(_srf(_liquidated_row()), _prior_soup())
        choices = _balance_recs(recs)[0]["balanceChoices"]
        self.assertEqual(len(choices), 2)
        self.assertEqual(sorted(round(float(c["value"]), 2) for c in choices),
                         [0.0, round(BEGIN, 2)])

    def test_the_restore_value_is_the_current_month_beginning_balance(self):
        _s, recs, _c = _run(_srf(_liquidated_row()), _prior_soup())
        self.assertEqual(_balance_recs(recs)[0]["balanceRestoreValue"], round(BEGIN, 2))

    def test_it_reports_the_lenders_ending_balance_as_reported(self):
        _s, recs, _c = _run(_srf(_liquidated_row()), _prior_soup())
        self.assertEqual(_balance_recs(recs)[0]["reported"], 0.0)

    def test_it_groups_with_its_loan_status_sibling(self):
        _s, recs, _c = _run(_srf(_liquidated_row()), _prior_soup())
        status = _of_field(recs, "Loan Status")[0]
        balance = _balance_recs(recs)[0]
        self.assertEqual(status["gp"], balance["gp"])
        self.assertEqual(status["category"], balance["category"])

    def test_ids_are_distinct_so_each_is_decided_independently(self):
        _s, recs, cands = _run(_srf(_liquidated_row()), _prior_soup())
        ids = [r["id"] for r in recs]
        self.assertEqual(len(ids), len(set(ids)))
        self.assertEqual(len(cands), len(recs))


class ReportedEndingBalanceMustBeZeroTests(unittest.TestCase):
    """CONDITION 4 (2026-08-15). The decision exists to undo a ZEROED balance.

    A LIQUIDATION -> ACTIVE reversal whose reported Ending Balance is already
    positive must NOT receive it, because neither offered choice is correct for
    such a loan: "restore to Beginning" moves a figure that was right, and "keep
    0.00" destroys it. Measured on Mission Valley GP 4778059109 -- reported ending
    276,702.10, beginning 279,150.10, and the client's own corrected file left the
    276,702.10 exactly as submitted."""

    def test_a_positive_reported_ending_balance_raises_no_decision(self):
        _s, recs, _c = _run(_srf(_liquidated_row(beginning=279150.10,
                                                 ending=276702.10)),
                            _prior_soup(prior_ending=279150.10))
        self.assertEqual(_balance_recs(recs), [])

    def test_the_loan_status_decision_is_still_offered(self):
        """Only the balance companion is withheld -- the reversal itself is still
        reviewed exactly as before this feature existed."""
        _s, recs, _c = _run(_srf(_liquidated_row(beginning=279150.10,
                                                 ending=276702.10)),
                            _prior_soup(prior_ending=279150.10))
        self.assertTrue(_of_field(recs, "Loan Status"))

    def test_a_small_positive_ending_balance_still_raises_no_decision(self):
        _s, recs, _c = _run(_srf(_liquidated_row(ending=0.01)), _prior_soup())
        self.assertEqual(_balance_recs(recs), [])

    def test_a_blank_reported_ending_balance_raises_no_decision(self):
        """Blank is not zero. An unknown closing balance must not be offered a
        restore decision."""
        _s, recs, _c = _run(_srf(_liquidated_row(ending=None)), _prior_soup())
        self.assertEqual(_balance_recs(recs), [])

    def test_a_negative_reported_ending_balance_raises_no_decision(self):
        _s, recs, _c = _run(_srf(_liquidated_row(ending=-1.0)), _prior_soup())
        self.assertEqual(_balance_recs(recs), [])

    def test_exactly_zero_does_raise_the_decision(self):
        _s, recs, _c = _run(_srf(_liquidated_row(ending=0.0)), _prior_soup())
        self.assertEqual(len(_balance_recs(recs)), 1)

    def test_the_predicate_itself(self):
        self.assertTrue(corrections._reported_balance_is_zero(0.0))
        self.assertTrue(corrections._reported_balance_is_zero(0))
        self.assertTrue(corrections._reported_balance_is_zero(0.001))   # rounds to 0.00
        self.assertFalse(corrections._reported_balance_is_zero(None))
        self.assertFalse(corrections._reported_balance_is_zero(""))
        self.assertFalse(corrections._reported_balance_is_zero(0.01))
        self.assertFalse(corrections._reported_balance_is_zero(276702.10))
        self.assertFalse(corrections._reported_balance_is_zero(float("nan")))
        self.assertFalse(corrections._reported_balance_is_zero(float("inf")))

    def test_a_positively_balanced_reversal_writes_no_financial_override(self):
        """End to end: the whole feature is inert on such a loan."""
        rows = [_liquidated_row(beginning=279150.10, ending=276702.10)]

        def decisions(recs):
            return [{"id": r["id"], "decision": "approve", "approved_value": "ACTIVE"}
                    for r in recs if r["field"] == "Loan Status"]

        result, _recs, wb = _review(rows, _prior_soup(prior_ending=279150.10), decisions)
        written = {f for (_gp, f) in result["verified_overrides"]}
        self.assertEqual(written, {"Loan Status"})
        self.assertEqual(_cell(wb, HEADER_ROW + 1, "Ending Balance"), 276702.10)
        self.assertEqual(_cell(wb, HEADER_ROW + 1, "Liquidation Date"), LIQ_DATE)


class ProposedStatusGatingTests(unittest.TestCase):
    """The decision belongs to LIQUIDATION -> ACTIVE only."""

    def test_a_liquidation_loan_the_reviewer_leaves_liquidated_is_untouched(self):
        """Approving LIQUIDATION (not ACTIVE) writes no balance override."""
        rows = [_liquidated_row()]

        def decisions(recs):
            return [{"id": r["id"], "decision": "approve",
                     "approved_value": "LIQUIDATION"}
                    for r in recs if r["field"] == "Loan Status"]

        result, _recs, wb = _review(rows, _prior_soup(), decisions)
        overrides = result["verified_overrides"]
        self.assertNotIn((GP_LIQ, corrections.BALANCE_DECISION_FIELD), overrides)
        self.assertEqual(_cell(wb, HEADER_ROW + 1, "Ending Balance"), 0.0)

    def test_an_ordinary_active_loan_raises_no_balance_decision(self):
        rows = [_row(GP_OTHER, status="ACTIVE", dq="CURRENT")]
        _s, recs, _c = _run(_srf(*rows), _prior_soup(GP_OTHER))
        self.assertEqual(_balance_recs(recs), [])

    def test_a_deferred_correction_raises_no_balance_decision(self):
        rows = [_row(GP_OTHER, status="ACTIVE", dq="CURRENT", received=0.0)]
        _s, recs, _c = _run(_srf(*rows),
                            [(GP_OTHER, "2026-05-01", "DEFERRED", "CURRENT", BEGIN)])
        self.assertEqual(_balance_recs(recs), [])


# ---------------------------------------------------------------------------
# 2 -- prior month evidence
# ---------------------------------------------------------------------------

class PriorEndingBalanceEvidenceTests(unittest.TestCase):

    def test_matching_prior_balance_is_shown_and_flagged_as_agreeing(self):
        _s, recs, _c = _run(_srf(_liquidated_row()), _prior_soup(prior_ending=BEGIN))
        rec = _balance_recs(recs)[0]
        self.assertEqual(rec["priorEndingBalance"], round(BEGIN, 2))
        self.assertFalse(rec["balanceEvidenceConflict"])

    def test_differing_prior_balance_is_shown_and_flagged_as_conflicting(self):
        _s, recs, _c = _run(_srf(_liquidated_row()),
                            _prior_soup(prior_ending=19000.00))
        rec = _balance_recs(recs)[0]
        self.assertEqual(rec["priorEndingBalance"], 19000.00)
        self.assertEqual(rec["beginningBalance"], round(BEGIN, 2))
        self.assertTrue(rec["balanceEvidenceConflict"])

    def test_a_conflict_does_not_change_the_two_permitted_choices(self):
        """Prior Month Ending Balance is EVIDENCE ONLY -- never selectable,
        even when it disagrees with the Beginning Balance."""
        _s, recs, _c = _run(_srf(_liquidated_row()),
                            _prior_soup(prior_ending=19000.00))
        choices = _balance_recs(recs)[0]["balanceChoices"]
        self.assertEqual(sorted(round(float(c["value"]), 2) for c in choices),
                         [0.0, round(BEGIN, 2)])
        self.assertNotIn(19000.00, [round(float(c["value"]), 2) for c in choices])

    def test_a_conflicting_prior_balance_is_not_writable(self):
        ok, _reason = corrections.validate_approved_value(
            corrections.BALANCE_DECISION_FIELD, 19000.00,
            balance_restore_value=round(BEGIN, 2))
        self.assertFalse(ok)

    def test_absent_prior_balance_still_raises_the_decision(self):
        """A SOUP row with no Ending Actual Balance: the decision must still be
        raised, with no conflict warning, because there is nothing to conflict."""
        _s, recs, _c = _run(_srf(_liquidated_row()), _prior_soup(prior_ending=None))
        rec = _balance_recs(recs)[0]
        self.assertIsNone(rec["priorEndingBalance"])
        self.assertFalse(rec["balanceEvidenceConflict"])
        self.assertEqual(rec["balanceRestoreValue"], round(BEGIN, 2))


# ---------------------------------------------------------------------------
# 3 -- the domain restriction: no free-form balance
# ---------------------------------------------------------------------------

class BalanceDomainRestrictionTests(unittest.TestCase):

    FIELD = None  # set in setUp so the constant is read from production code

    def setUp(self):
        self.FIELD = corrections.BALANCE_DECISION_FIELD

    def test_the_restore_value_is_permitted(self):
        ok, value = corrections.validate_approved_value(
            self.FIELD, BEGIN, balance_restore_value=round(BEGIN, 2))
        self.assertTrue(ok)
        self.assertEqual(value, round(BEGIN, 2))

    def test_zero_is_permitted(self):
        ok, value = corrections.validate_approved_value(
            self.FIELD, 0.0, balance_restore_value=round(BEGIN, 2))
        self.assertTrue(ok)
        self.assertEqual(value, 0.0)

    def test_an_arbitrary_amount_is_refused(self):
        ok, reason = corrections.validate_approved_value(
            self.FIELD, 12345.67, balance_restore_value=round(BEGIN, 2))
        self.assertFalse(ok)
        self.assertIn("permitted", reason.lower())

    def test_a_near_miss_amount_is_refused(self):
        ok, _reason = corrections.validate_approved_value(
            self.FIELD, round(BEGIN, 2) + 0.01, balance_restore_value=round(BEGIN, 2))
        self.assertFalse(ok)

    def test_a_negative_amount_is_refused(self):
        ok, _reason = corrections.validate_approved_value(
            self.FIELD, -BEGIN, balance_restore_value=round(BEGIN, 2))
        self.assertFalse(ok)

    def test_a_non_numeric_value_is_refused(self):
        ok, _reason = corrections.validate_approved_value(
            self.FIELD, "restore it", balance_restore_value=round(BEGIN, 2))
        self.assertFalse(ok)

    def test_without_a_server_side_restore_value_nothing_is_writable(self):
        """The allow-list is built from the SERVER-SIDE candidate. With no
        restore value there is no domain, so no write is possible -- a payload
        cannot mint one."""
        ok, _reason = corrections.validate_approved_value(
            self.FIELD, BEGIN, balance_restore_value=None)
        self.assertFalse(ok)

    def test_a_status_value_is_refused_for_the_balance_field(self):
        ok, _reason = corrections.validate_approved_value(
            self.FIELD, "ACTIVE", balance_restore_value=round(BEGIN, 2))
        self.assertFalse(ok)


class NoGeneralBalanceEditingTests(unittest.TestCase):
    """The guard on scope: Ending Balance must not become generally editable."""

    def test_ending_balance_is_not_offered_on_an_ordinary_loan(self):
        _s, recs, _c = _run(_srf(_row(GP_OTHER, dq="CURRENT")), _prior_soup(GP_OTHER))
        self.assertEqual(_balance_recs(recs), [])

    def test_a_balance_decision_cannot_be_forged_for_an_unlisted_loan(self):
        """A decision id that no candidate owns is ignored, exactly as any
        unknown id already is."""
        rows = [_liquidated_row()]
        srf = _srf(*rows)
        _stat01, _recs, cands = _run(srf, _prior_soup())
        forged = [{"id": "%s||%s||%s" % (SERVICER, GP_OTHER,
                                         corrections.BALANCE_DECISION_FIELD),
                   "decision": "approve", "approved_value": 99999.99}]
        result = corrections.generate_reviewed_workbook(
            _workbook_bytes(rows), "test.xlsx", SERVICER, cands, forged,
            "tester", "sha", None)
        self.assertNotIn((GP_OTHER, corrections.BALANCE_DECISION_FIELD),
                         result["verified_overrides"])

    def test_only_three_fields_can_carry_an_override(self):
        self.assertEqual(
            set(checks._VERIFIED_OVERRIDE_COLUMNS),
            {"Loan Status", "Delinquency Status", corrections.BALANCE_DECISION_FIELD,
             "Liquidation Date", "Liquidation Loss"})


# ---------------------------------------------------------------------------
# 4 -- restore
# ---------------------------------------------------------------------------

class ReviewerChoosesRestoreTests(unittest.TestCase):

    def setUp(self):
        self.rows = [_liquidated_row()]
        self.result, self.recs, self.wb = _review(
            self.rows, _prior_soup(), _approve_status_and_balance(BEGIN))
        self.overrides = self.result["verified_overrides"]

    def test_the_balance_override_is_the_beginning_balance(self):
        self.assertEqual(
            self.overrides[(GP_LIQ, corrections.BALANCE_DECISION_FIELD)],
            round(BEGIN, 2))

    def test_the_status_override_is_active(self):
        self.assertEqual(self.overrides[(GP_LIQ, "Loan Status")], "ACTIVE")

    def test_liquidation_date_is_cleared(self):
        self.assertIn((GP_LIQ, "Liquidation Date"), self.overrides)
        self.assertEqual(self.overrides[(GP_LIQ, "Liquidation Date")], "")

    def test_liquidation_loss_is_cleared(self):
        self.assertIn((GP_LIQ, "Liquidation Loss"), self.overrides)
        self.assertIsNone(self.overrides[(GP_LIQ, "Liquidation Loss")])

    def test_the_workbook_cell_carries_the_restored_balance(self):
        self.assertEqual(_cell(self.wb, HEADER_ROW + 1, "Ending Balance"),
                         round(BEGIN, 2))

    def test_the_workbook_liquidation_cells_are_cleared(self):
        self.assertIn(_cell(self.wb, HEADER_ROW + 1, "Liquidation Date"), (None, ""))
        self.assertIn(_cell(self.wb, HEADER_ROW + 1, "Liquidation Loss"), (None, ""))

    def test_the_beginning_balance_cell_is_never_touched(self):
        self.assertEqual(_cell(self.wb, HEADER_ROW + 1, "Beginning Balance"),
                         round(BEGIN, 2))


class ReviewerChoosesKeepZeroTests(unittest.TestCase):

    def setUp(self):
        self.rows = [_liquidated_row()]
        self.result, self.recs, self.wb = _review(
            self.rows, _prior_soup(), _approve_status_and_balance(0.0))
        self.overrides = self.result["verified_overrides"]

    def test_the_balance_stays_zero(self):
        self.assertEqual(
            self.overrides[(GP_LIQ, corrections.BALANCE_DECISION_FIELD)], 0.0)
        self.assertEqual(_cell(self.wb, HEADER_ROW + 1, "Ending Balance"), 0.0)

    def test_liquidation_date_is_PRESERVED(self):
        self.assertNotIn((GP_LIQ, "Liquidation Date"), self.overrides)
        self.assertEqual(_cell(self.wb, HEADER_ROW + 1, "Liquidation Date"), LIQ_DATE)

    def test_liquidation_loss_is_PRESERVED(self):
        self.assertNotIn((GP_LIQ, "Liquidation Loss"), self.overrides)
        self.assertEqual(_cell(self.wb, HEADER_ROW + 1, "Liquidation Loss"), LIQ_LOSS)

    def test_the_status_override_still_applies(self):
        self.assertEqual(self.overrides[(GP_LIQ, "Loan Status")], "ACTIVE")


class ClearingRequiresExplicitApprovalTests(unittest.TestCase):
    """Liquidation fields clear ONLY on an explicit restore -- never on the
    Loan Status approval alone, and never on reject/return."""

    def _overrides(self, decisions):
        rows = [_liquidated_row()]
        result, _recs, _wb = _review(rows, _prior_soup(), decisions)
        return result["verified_overrides"]

    def test_status_approved_but_balance_left_undecided_clears_nothing(self):
        def decisions(recs):
            return [{"id": r["id"], "decision": "approve", "approved_value": "ACTIVE"}
                    for r in recs if r["field"] == "Loan Status"]
        overrides = self._overrides(decisions)
        self.assertNotIn((GP_LIQ, "Liquidation Date"), overrides)
        self.assertNotIn((GP_LIQ, "Liquidation Loss"), overrides)
        self.assertNotIn((GP_LIQ, corrections.BALANCE_DECISION_FIELD), overrides)

    def test_a_rejected_balance_decision_clears_nothing(self):
        def decisions(recs):
            out = []
            for r in recs:
                if r["field"] == "Loan Status":
                    out.append({"id": r["id"], "decision": "approve",
                                "approved_value": "ACTIVE"})
                elif r["field"] == corrections.BALANCE_DECISION_FIELD:
                    out.append({"id": r["id"], "decision": "reject"})
            return out
        overrides = self._overrides(decisions)
        self.assertNotIn((GP_LIQ, "Liquidation Date"), overrides)
        self.assertNotIn((GP_LIQ, "Liquidation Loss"), overrides)

    def test_a_returned_balance_decision_clears_nothing(self):
        def decisions(recs):
            out = []
            for r in recs:
                if r["field"] == "Loan Status":
                    out.append({"id": r["id"], "decision": "approve",
                                "approved_value": "ACTIVE"})
                elif r["field"] == corrections.BALANCE_DECISION_FIELD:
                    out.append({"id": r["id"], "decision": "return"})
            return out
        overrides = self._overrides(decisions)
        self.assertNotIn((GP_LIQ, "Liquidation Date"), overrides)
        self.assertNotIn((GP_LIQ, "Liquidation Loss"), overrides)

    def test_a_refused_balance_value_clears_nothing(self):
        def decisions(recs):
            out = []
            for r in recs:
                if r["field"] == "Loan Status":
                    out.append({"id": r["id"], "decision": "approve",
                                "approved_value": "ACTIVE"})
                elif r["field"] == corrections.BALANCE_DECISION_FIELD:
                    out.append({"id": r["id"], "decision": "approve",
                                "approved_value": 12345.67})
            return out
        overrides = self._overrides(decisions)
        self.assertNotIn((GP_LIQ, "Liquidation Date"), overrides)
        self.assertNotIn((GP_LIQ, corrections.BALANCE_DECISION_FIELD), overrides)


# ---------------------------------------------------------------------------
# 5 -- the write target: LAST posting row only
# ---------------------------------------------------------------------------

class MultiPostingWriteTargetTests(unittest.TestCase):
    """A balance is NOT a loan-level value. Writing it to every posting row
    would corrupt the balance chain -- measured at 59 of 1,234 Lendistry 2025
    loans (4.8%) carrying more than one posting."""

    def setUp(self):
        # three postings; the cycle closes on the third
        self.rows = [
            _liquidated_row(excel_row=HEADER_ROW + 1, instance=1, last="N",
                            beginning=BEGIN, ending=15000.00,
                            liquidation_date=None, liquidation_loss=None),
            _liquidated_row(excel_row=HEADER_ROW + 2, instance=2, last="N",
                            beginning=15000.00, ending=10000.00,
                            liquidation_date=None, liquidation_loss=None),
            _liquidated_row(excel_row=HEADER_ROW + 3, instance=3, last="Y",
                            beginning=10000.00, ending=0.0),
        ]
        self.result, self.recs, self.wb = _review(
            self.rows, _prior_soup(), _approve_status_and_balance(BEGIN))

    def test_the_balance_is_written_only_to_the_last_posting_row(self):
        self.assertEqual(_cell(self.wb, HEADER_ROW + 3, "Ending Balance"),
                         round(BEGIN, 2))

    def test_earlier_posting_balances_are_untouched(self):
        self.assertEqual(_cell(self.wb, HEADER_ROW + 1, "Ending Balance"), 15000.00)
        self.assertEqual(_cell(self.wb, HEADER_ROW + 2, "Ending Balance"), 10000.00)

    def test_loan_status_still_writes_to_every_posting_row(self):
        """Loan Status IS loan-level; that behaviour must not change."""
        for offset in (1, 2, 3):
            self.assertEqual(_cell(self.wb, HEADER_ROW + offset, "Loan Status"),
                             "ACTIVE")

    def test_the_log_records_one_row_updated_for_the_balance(self):
        rows = _log_by_field(self.wb)[corrections.BALANCE_DECISION_FIELD]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Rows Updated"], 1)
        self.assertEqual(str(rows[0]["Excel Row Numbers"]), str(HEADER_ROW + 3))

    def test_the_log_shows_loan_status_updating_all_three_rows(self):
        rows = _log_by_field(self.wb)["Loan Status"]
        self.assertEqual(rows[0]["Rows Updated"], 3)


# ---------------------------------------------------------------------------
# 6 -- audit trail
# ---------------------------------------------------------------------------

class CorrectionLogTests(unittest.TestCase):

    def setUp(self):
        self.rows = [_liquidated_row()]
        self.result, self.recs, self.wb = _review(
            self.rows, _prior_soup(prior_ending=19000.00),
            _approve_status_and_balance(BEGIN))
        self.by_field = _log_by_field(self.wb)

    def test_every_changed_field_gets_its_own_log_row(self):
        for field in ("Loan Status", corrections.BALANCE_DECISION_FIELD,
                      "Liquidation Date", "Liquidation Loss"):
            self.assertIn(field, self.by_field, "missing log row for %s" % field)

    def test_the_balance_row_records_original_and_approved(self):
        row = self.by_field[corrections.BALANCE_DECISION_FIELD][0]
        self.assertEqual(row["Original Ending Balance"], 0.0)
        self.assertEqual(row["Approved Ending Balance"], round(BEGIN, 2))

    def test_the_balance_row_records_both_evidence_balances(self):
        row = self.by_field[corrections.BALANCE_DECISION_FIELD][0]
        self.assertEqual(row["Beginning Balance (Evidence)"], round(BEGIN, 2))
        self.assertEqual(row["Prior Month Ending Balance (Evidence)"], 19000.00)

    def test_the_liquidation_rows_record_original_and_approved(self):
        date_row = self.by_field["Liquidation Date"][0]
        self.assertEqual(str(date_row["Original Liquidation Date"]), LIQ_DATE)
        self.assertIn(date_row["Approved Liquidation Date"], (None, ""))
        loss_row = self.by_field["Liquidation Loss"][0]
        self.assertEqual(loss_row["Original Liquidation Loss"], LIQ_LOSS)
        self.assertIn(loss_row["Approved Liquidation Loss"], (None, ""))

    def test_the_loan_status_row_records_original_and_approved_status(self):
        row = self.by_field["Loan Status"][0]
        self.assertEqual(str(row["Original Value"]).upper(), "LIQUIDATION")
        self.assertEqual(row["Reviewer Approved Value"], "ACTIVE")

    def test_every_row_is_attributed_to_a_reviewer_not_the_system(self):
        for field in (corrections.BALANCE_DECISION_FIELD, "Liquidation Date",
                      "Liquidation Loss"):
            row = self.by_field[field][0]
            self.assertEqual(row["Decision"],
                             corrections._DECISION_LABELS[corrections._DECISION_APPROVE])
            self.assertTrue(row["Eligibility Rule"])
            self.assertIn("review", str(row["Eligibility Rule"]).lower())
            self.assertTrue(row["Reviewed By (self-declared)"])
            self.assertTrue(row["Reviewed On"])

    def test_a_keep_zero_decision_is_recorded_as_zero_not_blanked(self):
        """"Keep 0.00" IS the reviewer's answer -- the log must not erase it."""
        _r, _recs, wb = _review([_liquidated_row()], _prior_soup(),
                                _approve_status_and_balance(0.0))
        row = _log_by_field(wb)[corrections.BALANCE_DECISION_FIELD][0]
        self.assertEqual(row["Approved Ending Balance"], 0.0)
        self.assertNotIn("Liquidation Date", _log_by_field(wb))

    def test_the_new_headers_exist_and_extend_rather_than_reorder(self):
        headers = corrections.CORRECTION_LOG_HEADERS
        for name in ("Beginning Balance (Evidence)",
                     "Prior Month Ending Balance (Evidence)",
                     "Original Ending Balance", "Approved Ending Balance",
                     "Original Liquidation Date", "Approved Liquidation Date",
                     "Original Liquidation Loss", "Approved Liquidation Loss"):
            self.assertIn(name, headers)
        self.assertEqual(headers[:22][5], "Field Updated")
        self.assertEqual(headers[:22][6], "Original Value")
        self.assertEqual(headers[:22][7], "Corrected Value")

    def test_the_written_log_sheet_carries_the_values(self):
        ws = self.wb[corrections.CORRECTION_LOG_SHEET_NAME]
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        self.assertEqual(headers, corrections.CORRECTION_LOG_HEADERS)
        col = headers.index("Approved Ending Balance") + 1
        field_col = headers.index("Field Updated") + 1
        found = [ws.cell(row=r, column=col).value
                 for r in range(2, ws.max_row + 1)
                 if ws.cell(row=r, column=field_col).value
                 == corrections.BALANCE_DECISION_FIELD]
        self.assertEqual(found, [round(BEGIN, 2)])

    def test_unrelated_corrections_leave_the_new_columns_blank(self):
        """A Delinquency Status correction must not acquire balance columns."""
        rows = [_row(GP_OTHER, dq="CURRENT", received=0.0, excel_row=HEADER_ROW + 1)]
        srf = _srf(*rows)
        _s, recs, cands = _run(srf, _prior_soup(GP_OTHER))
        dq = [r for r in recs if r["field"] == "Delinquency Status"]
        if not dq:
            self.skipTest("fixture produced no Delinquency Status recommendation")
        result = corrections.generate_reviewed_workbook(
            _workbook_bytes(rows), "t.xlsx", SERVICER, cands,
            [{"id": dq[0]["id"], "decision": "approve"}], "tester", "sha", None)
        wb = load_workbook(io.BytesIO(result["file_bytes"]))
        for row in _log_rows(wb):
            for header in ("Beginning Balance (Evidence)", "Original Ending Balance",
                           "Approved Ending Balance", "Original Liquidation Date"):
                self.assertIn(row[header], (None, ""))


# ---------------------------------------------------------------------------
# 7 -- downstream consumption
# ---------------------------------------------------------------------------

class DownstreamPortfolioSummaryTests(unittest.TestCase):

    def _outstanding(self, overrides):
        srf = _srf(_liquidated_row(), _row(GP_OTHER, ending=100000.0,
                                          excel_row=HEADER_ROW + 2))
        views = checks.verified_portfolio_summary_views(
            srf, overrides, servicer_name=SERVICER, excluded_gps=None)
        return views["dollars"]["Outstanding ($)"], views["counts"]["Outstanding"]

    def test_status_only_override_leaves_the_balance_out_of_outstanding(self):
        """The defect, pinned: ACTIVE by count, $0 by dollars."""
        dollars, count = self._outstanding({(GP_LIQ, "Loan Status"): "ACTIVE"})
        self.assertEqual(count, 2)
        self.assertEqual(dollars, 100000.0)

    def test_the_approved_balance_reaches_outstanding_dollars(self):
        dollars, count = self._outstanding({
            (GP_LIQ, "Loan Status"): "ACTIVE",
            (GP_LIQ, corrections.BALANCE_DECISION_FIELD): round(BEGIN, 2)})
        self.assertEqual(count, 2)
        self.assertEqual(dollars, round(100000.0 + BEGIN, 2))

    def test_keep_zero_leaves_outstanding_dollars_unchanged(self):
        dollars, _count = self._outstanding({
            (GP_LIQ, "Loan Status"): "ACTIVE",
            (GP_LIQ, corrections.BALANCE_DECISION_FIELD): 0.0})
        self.assertEqual(dollars, 100000.0)

    def test_an_unrelated_loan_is_never_affected(self):
        dollars, _count = self._outstanding({
            (GP_LIQ, "Loan Status"): "ACTIVE",
            (GP_LIQ, corrections.BALANCE_DECISION_FIELD): round(BEGIN, 2)})
        base, _c = self._outstanding({})
        self.assertEqual(dollars - round(BEGIN, 2), base)


class DownstreamMonthlyServicerReportTests(unittest.TestCase):

    def _rows(self, overrides):
        import monthly_servicer_report as msr
        srf = _srf(_liquidated_row())
        rows = msr.build_servicer_rows(srf, SERVICER, verified_overrides=overrides)
        return {str(r["SBA GP Number"]): r for r in rows}

    def test_the_msr_receives_the_approved_ending_balance(self):
        row = self._rows({
            (GP_LIQ, "Loan Status"): "ACTIVE",
            (GP_LIQ, corrections.BALANCE_DECISION_FIELD): round(BEGIN, 2)})[GP_LIQ]
        self.assertEqual(row["Ending Actual Balance"], round(BEGIN, 2))
        self.assertEqual(row["Loan Status"], "ACTIVE")

    def test_the_msr_cannot_show_active_with_surviving_liquidation_fields(self):
        """The whole point: no ACTIVE + restored balance + live liquidation."""
        row = self._rows({
            (GP_LIQ, "Loan Status"): "ACTIVE",
            (GP_LIQ, corrections.BALANCE_DECISION_FIELD): round(BEGIN, 2),
            (GP_LIQ, "Liquidation Date"): "",
            (GP_LIQ, "Liquidation Loss"): None})[GP_LIQ]
        self.assertEqual(row["Loan Status"], "ACTIVE")
        self.assertEqual(row["Ending Actual Balance"], round(BEGIN, 2))
        self.assertIn(row["Liquidation Date"], (None, ""))
        self.assertIn(row["Liquidation Loss"], (None, ""))

    def test_keep_zero_leaves_the_liquidation_fields_visible(self):
        row = self._rows({
            (GP_LIQ, "Loan Status"): "ACTIVE",
            (GP_LIQ, corrections.BALANCE_DECISION_FIELD): 0.0})[GP_LIQ]
        self.assertEqual(row["Ending Actual Balance"], 0.0)
        self.assertIsNotNone(row["Liquidation Date"])
        self.assertIsNotNone(row["Liquidation Loss"])

    def test_no_override_leaves_the_lender_values_untouched(self):
        row = self._rows({})[GP_LIQ]
        self.assertEqual(row["Loan Status"], "LIQUIDATION")
        self.assertEqual(row["Ending Actual Balance"], 0.0)


# ---------------------------------------------------------------------------
# 8 -- D4: an undecided balance decision blocks reporting
# ---------------------------------------------------------------------------

class ReportingReadinessTests(unittest.TestCase):
    """The balance decision is a CANDIDATE, so D4's existing pending count
    blocks reporting with no change to the D4 code."""

    def test_the_balance_decision_is_a_candidate(self):
        _s, recs, cands = _run(_srf(_liquidated_row()), _prior_soup())
        ids = {c["id"] for c in cands}
        for rec in _balance_recs(recs):
            self.assertIn(rec["id"], ids)

    def test_two_candidates_exist_for_one_liquidation_loan(self):
        _s, _recs, cands = _run(_srf(_liquidated_row()), _prior_soup())
        fields = sorted(c["field"] for c in cands)
        self.assertIn(corrections.BALANCE_DECISION_FIELD, fields)
        self.assertIn("Loan Status", fields)


# ---------------------------------------------------------------------------
# 9 -- no hardcoded regression data in production code
# ---------------------------------------------------------------------------

class NoHardcodedRegressionDataTests(unittest.TestCase):

    PRODUCTION = ("checks.py", "corrections.py", "monthly_servicer_report.py",
                  "main.py", "report.py")
    FORBIDDEN = ("5886179101", "5991379104", "20088.08", "20203.66", "40291.74",
                 "45489960.19", "45449668.45")

    def test_no_june_gp_number_or_amount_appears_in_production_code(self):
        here = os.path.dirname(os.path.abspath(__file__))
        for name in self.PRODUCTION:
            path = os.path.join(here, name)
            if not os.path.exists(path):
                continue
            text = open(path, encoding="utf-8").read()
            for token in self.FORBIDDEN:
                with self.subTest(file=name, token=token):
                    self.assertNotIn(token, text)

    def test_the_frontend_carries_no_regression_data_either(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                            "frontend", "src", "components", "CorrectionReview.jsx")
        if not os.path.exists(path):
            self.skipTest("CorrectionReview.jsx not found")
        text = open(path, encoding="utf-8").read()
        for token in self.FORBIDDEN:
            with self.subTest(token=token):
                self.assertNotIn(token, text)


if __name__ == "__main__":
    unittest.main()
