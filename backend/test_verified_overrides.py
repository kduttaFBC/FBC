"""generate_reviewed_workbook returns the VERIFIED layer
(client-confirmed 2026-08-10).

Three layers:

    Original   what the lender submitted
    Audited    what the audit calculated / recommended
    Verified   the final value after Manual Review

verified_overrides is the third. It carries {(gp, field): approved_value} for
every correction that was approved, validated AND successfully written -- and
nothing else. Absence of an entry means the lender's original value stands,
which is equally true of the corrected workbook, since nothing was written to
those cells either.

The key property, asserted below several ways: the override map and the
workbook are populated from the SAME branch, so they cannot disagree. That is
what lets the corrected workbook and the Final Audited Portfolio Summary be
generated from one source.

Nothing consumes verified_overrides yet -- this commit is additive.

Run with:
    python -m pytest test_verified_overrides.py
"""

import io
import unittest

import pandas as pd
from openpyxl import Workbook, load_workbook

import corrections


GP = "1234567890"
SHEET = "Servicer Inputs"
HEADER_ROW = 9
DATA_ROW = 10


def _workbook_bytes(loan_status="ACTIVE", dq_status="CURRENT", gp=GP):
    """A minimal workbook shaped like the real SRF: header on row 9, one loan."""
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    headers = ["SBA GP Number", "Lender Loan Number", "Loan Status",
               "Delinquency Status", "Ending Balance"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=HEADER_ROW, column=col, value=h)
    ws.cell(row=DATA_ROW, column=1, value=gp)
    ws.cell(row=DATA_ROW, column=2, value=f"L-{gp}")
    ws.cell(row=DATA_ROW, column=3, value=loan_status)
    ws.cell(row=DATA_ROW, column=4, value=dq_status)
    ws.cell(row=DATA_ROW, column=5, value=100000.0)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _candidate(field, corrected_value, gp=GP, rec_id="c1"):
    return {
        "id": rec_id, "gp": gp, "field": field,
        "corrected_value": corrected_value,
        "reported_value": "ACTIVE" if field == "Loan Status" else "CURRENT",
        "excel_rows": [DATA_ROW], "skipped_row_refs": [],
        "sheet_name": SHEET, "header_excel_row": HEADER_ROW,
        "eligibility_rule": "test", "review_required": False, "review_reason": None,
    }


def _run(candidates, decisions, file_bytes=None):
    return corrections.generate_reviewed_workbook(
        file_bytes if file_bytes is not None else _workbook_bytes(),
        "original.xlsx", "Test Bank", candidates, decisions,
        "Test Reviewer", "abc123", pd.Timestamp("2026-06-01"),
    )


def _cell(outcome, column):
    ws = load_workbook(io.BytesIO(outcome["file_bytes"]))[SHEET]
    return ws.cell(row=DATA_ROW, column=column).value


LOAN_STATUS_COL = 3
DQ_COL = 4


class ApproveRecordsTheVerifiedValueTests(unittest.TestCase):

    def test_approve_without_substitution_records_the_audited_value(self):
        out = _run([_candidate("Delinquency Status", "60 DAYS")],
                   [{"id": "c1", "decision": "approve"}])
        self.assertEqual(out["verified_overrides"], {(GP, "Delinquency Status"): "60 DAYS"})
        self.assertEqual(out["applied_count"], 1)

    def test_approve_with_substitution_records_the_CLIENTS_value(self):
        # The whole point of the three-layer model: when the client modifies
        # the recommendation, verified is the CLIENT's value, not the audit's.
        out = _run([_candidate("Delinquency Status", "60 DAYS")],
                   [{"id": "c1", "decision": "approve", "approved_value": "30 DAYS"}])
        self.assertEqual(out["verified_overrides"], {(GP, "Delinquency Status"): "30 DAYS"})
        self.assertNotIn("60 DAYS", out["verified_overrides"].values())

    def test_loan_status_approval_is_recorded(self):
        out = _run([_candidate("Loan Status", "DEFERRED")],
                   [{"id": "c1", "decision": "approve"}])
        self.assertEqual(out["verified_overrides"], {(GP, "Loan Status"): "DEFERRED"})

    def test_a_reviewer_selected_terminal_status_is_recorded(self):
        # A human may select any status in the client's authoritative domain.
        out = _run([_candidate("Loan Status", "ACTIVE")],
                   [{"id": "c1", "decision": "approve", "approved_value": "LIQUIDATION"}])
        self.assertEqual(out["verified_overrides"], {(GP, "Loan Status"): "LIQUIDATION"})

    def test_the_recorded_value_is_the_NORMALISED_one(self):
        # validate_approved_value canonicalises; the override must carry the
        # workbook's own spelling, not whatever casing was submitted.
        out = _run([_candidate("Loan Status", "ACTIVE")],
                   [{"id": "c1", "decision": "approve", "approved_value": "deferred"}])
        self.assertEqual(out["verified_overrides"], {(GP, "Loan Status"): "DEFERRED"})


class NonApprovalsRecordNothingTests(unittest.TestCase):
    """Every path that does not write must leave no entry -- absence is what
    'the lender's original value stands' is expressed as."""

    def test_reject_records_nothing(self):
        out = _run([_candidate("Delinquency Status", "60 DAYS")],
                   [{"id": "c1", "decision": "reject"}])
        self.assertEqual(out["verified_overrides"], {})
        self.assertEqual(out["rejected_count"], 1)

    def test_return_to_servicer_records_nothing(self):
        out = _run([_candidate("Delinquency Status", "60 DAYS")],
                   [{"id": "c1", "decision": "return"}])
        self.assertEqual(out["verified_overrides"], {})
        self.assertEqual(out["returned_count"], 1)

    def test_an_undecided_candidate_records_nothing(self):
        out = _run([_candidate("Delinquency Status", "60 DAYS", rec_id="c1"),
                    _candidate("Loan Status", "DEFERRED", rec_id="c2")],
                   [{"id": "c1", "decision": "approve"}])
        self.assertEqual(out["verified_overrides"], {(GP, "Delinquency Status"): "60 DAYS"})
        self.assertNotIn((GP, "Loan Status"), out["verified_overrides"])

    def test_a_domain_refused_substitution_records_nothing(self):
        # Cross-domain: a Delinquency Status value approved into Loan Status.
        out = _run([_candidate("Loan Status", "DEFERRED")],
                   [{"id": "c1", "decision": "approve", "approved_value": "60 DAYS"}])
        self.assertEqual(out["verified_overrides"], {})
        self.assertEqual(out["applied_count"], 0)
        self.assertTrue(out["corrections_not_applied"])

    def test_an_unrecognised_decision_records_nothing(self):
        out = _run([_candidate("Delinquency Status", "60 DAYS")],
                   [{"id": "c1", "decision": "maybe"}])
        self.assertEqual(out["verified_overrides"], {})

    def test_a_failed_write_records_nothing(self):
        # GP in the workbook does not match the candidate's GP, so
        # _apply_one_correction refuses to write. Approved, but not verified.
        out = _run([_candidate("Delinquency Status", "60 DAYS")],
                   [{"id": "c1", "decision": "approve"}],
                   file_bytes=_workbook_bytes(gp="9999999999"))
        self.assertEqual(out["verified_overrides"], {})
        self.assertEqual(out["applied_count"], 0)

    def test_an_early_return_still_carries_the_key(self):
        # One contract for every caller -- never "does the key exist?".
        out = _run([_candidate("Delinquency Status", "60 DAYS")],
                   [{"id": "unknown-id", "decision": "approve"}])
        self.assertIn("verified_overrides", out)
        self.assertEqual(out["verified_overrides"], {})


class OverridesAndWorkbookCannotDisagreeTests(unittest.TestCase):
    """The two artifacts come from one branch. Asserted against the real
    writer, not in theory."""

    def test_an_approved_value_is_in_BOTH_the_workbook_and_the_map(self):
        out = _run([_candidate("Delinquency Status", "60 DAYS")],
                   [{"id": "c1", "decision": "approve"}])
        self.assertEqual(_cell(out, DQ_COL), "60 DAYS")
        self.assertEqual(out["verified_overrides"][(GP, "Delinquency Status")], "60 DAYS")

    def test_a_rejected_value_is_in_NEITHER(self):
        out = _run([_candidate("Delinquency Status", "60 DAYS")],
                   [{"id": "c1", "decision": "reject"}])
        self.assertEqual(_cell(out, DQ_COL), "CURRENT")      # untouched
        self.assertEqual(out["verified_overrides"], {})

    def test_a_substituted_value_agrees_in_both(self):
        out = _run([_candidate("Loan Status", "DEFERRED")],
                   [{"id": "c1", "decision": "approve", "approved_value": "PAID-IN-FULL"}])
        self.assertEqual(_cell(out, LOAN_STATUS_COL), "PAID-IN-FULL")
        self.assertEqual(out["verified_overrides"][(GP, "Loan Status")], "PAID-IN-FULL")

    def test_the_key_set_equals_the_applied_log_rows(self):
        # applied_count and the override map are driven by the same condition.
        out = _run([_candidate("Delinquency Status", "60 DAYS", rec_id="c1"),
                    _candidate("Loan Status", "DEFERRED", rec_id="c2")],
                   [{"id": "c1", "decision": "approve"},
                    {"id": "c2", "decision": "reject"}])
        self.assertEqual(len(out["verified_overrides"]), out["applied_count"])

    def test_multiple_approvals_are_all_recorded(self):
        out = _run([_candidate("Delinquency Status", "60 DAYS", rec_id="c1"),
                    _candidate("Loan Status", "DEFERRED", rec_id="c2")],
                   [{"id": "c1", "decision": "approve"},
                    {"id": "c2", "decision": "approve"}])
        self.assertEqual(out["verified_overrides"],
                         {(GP, "Delinquency Status"): "60 DAYS",
                          (GP, "Loan Status"): "DEFERRED"})
        self.assertEqual(_cell(out, DQ_COL), "60 DAYS")
        self.assertEqual(_cell(out, LOAN_STATUS_COL), "DEFERRED")


class FinancialFieldsAreNotReachableTests(unittest.TestCase):
    """CLIENT-CONFIRMED (2026-08-10): Manual Review stays limited to status
    corrections. A financial value can only ever enter the verified layer from
    the lender's own submission."""

    def test_a_financial_field_is_refused_and_never_recorded(self):
        for field in ("Ending Balance", "Beginning Balance", "Payment Received"):
            with self.subTest(field=field):
                out = _run([_candidate(field, "12345.67")],
                           [{"id": "c1", "decision": "approve"}])
                self.assertEqual(out["verified_overrides"], {})
                self.assertEqual(out["applied_count"], 0)

    def test_the_correction_allow_list_never_proposes_a_financial_field(self):
        # The upstream gate: no financial field is ever a candidate at all.
        eligible = corrections.build_eligible_corrections([{
            "gp": GP, "reasons": ["DELINQUENCY STATUS MISMATCH"], "dq_match": False,
            "calculated_dq": "60 DAYS", "srf_dq": "CURRENT", "srf_status": "ACTIVE",
            "excel_rows": [DATA_ROW], "sheet_name": SHEET, "header_excel_row": HEADER_ROW,
        }])
        self.assertTrue(all(e["field"] in ("Loan Status", "Delinquency Status")
                            for e in eligible))


if __name__ == "__main__":
    unittest.main()
