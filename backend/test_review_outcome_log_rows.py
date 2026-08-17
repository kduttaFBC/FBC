"""generate_reviewed_workbook exposes its per-decision review record.

`log_rows` is already assembled to write the Correction Log sheet. Returning it
avoids a second, independent reconstruction of the same decision provenance
elsewhere -- it carries the submitted value, the engine's value, what the
reviewer authorised, the decision, its review reason and the reviewer's note,
for APPROVE, REJECT and RETURN alike. `verified_overrides` cannot express that:
it holds only successfully APPLIED approvals.

THIS COMMIT IS PURELY ADDITIVE and nothing reads the new key yet. These tests
exist to prove exactly that:

  * the key is on the success return AND on every early return
  * the workbook bytes, the Correction Log sheet, verified_overrides and every
    existing outcome key are unchanged by its presence
  * no existing caller can see the difference, because every access is by
    explicit named key
"""

import ast
import inspect
import io
import unittest

import openpyxl

import corrections


EXISTING_KEYS = {
    "file_bytes", "filename", "applied_count", "rejected_count",
    "returned_count", "approved_not_applied_count", "corrections_not_applied",
    "unknown_ids", "skip_reason", "verified_overrides",
}


def _candidate(gp="1000000001", field="Delinquency Status", rec_id="a1",
               reported="60 DAYS", calculated="30 DAYS"):
    return {
        "id": rec_id, "gp": gp, "field": field, "excel_rows": [10],
        "sheet_name": "Servicer Inputs", "header_excel_row": 9,
        "reported_value": reported, "corrected_value": calculated,
        "eligibility_rule": "STAT-01 deterministic calculation",
        "review_reason": None, "review_required": False,
    }


def _workbook_bytes():
    """A minimal Servicer Inputs workbook the correction writer can reopen."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Servicer Inputs"
    headers = ["SBA GP Number", "Lender Loan Number", "Loan Status",
               "Delinquency Status", "Beginning Balance", "Ending Balance",
               "Payment Received"]
    for col, name in enumerate(headers, start=1):
        ws.cell(row=9, column=col, value=name)
    ws.cell(row=10, column=1, value="1000000001")
    ws.cell(row=10, column=2, value="L-1")
    ws.cell(row=10, column=3, value="ACTIVE")
    ws.cell(row=10, column=4, value="60 DAYS")
    ws.cell(row=10, column=5, value=1000.0)
    ws.cell(row=10, column=6, value=900.0)
    ws.cell(row=10, column=7, value=100.0)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _run(decisions, candidates=None, file_bytes=None):
    return corrections.generate_reviewed_workbook(
        file_bytes if file_bytes is not None else _workbook_bytes(),
        "servicer.xlsx", "Test Bank",
        candidates if candidates is not None else [_candidate()],
        decisions, "K. Dutta", "abc123", None,
    )


class TheKeyIsAlwaysPresentTests(unittest.TestCase):
    """One contract, whichever path the function takes."""

    def test_the_success_return_carries_log_rows(self):
        outcome = _run([{"id": "a1", "decision": "approve"}])
        self.assertIn("log_rows", outcome)
        self.assertIsInstance(outcome["log_rows"], list)
        self.assertTrue(outcome["log_rows"])

    def test_no_decisions_early_return_carries_log_rows(self):
        outcome = _run([])
        self.assertIn("log_rows", outcome)
        self.assertEqual(outcome["log_rows"], [])
        self.assertIsNotNone(outcome["skip_reason"])

    def test_unreadable_workbook_early_return_carries_log_rows(self):
        outcome = _run([{"id": "a1", "decision": "approve"}], file_bytes=b"not a workbook")
        self.assertIn("log_rows", outcome)
        self.assertEqual(outcome["log_rows"], [])

    def test_no_candidates_early_return_carries_log_rows(self):
        outcome = _run([{"id": "a1", "decision": "approve"}], candidates=[])
        self.assertIn("log_rows", outcome)
        self.assertEqual(outcome["log_rows"], [])

    def test_every_early_return_uses_the_shared_empty_template(self):
        """So a future key cannot be added to one path and missed on another."""
        source = inspect.getsource(corrections.generate_reviewed_workbook)
        returns = [n for n in ast.walk(ast.parse(source.lstrip()))
                   if isinstance(n, ast.Return)]
        self.assertGreaterEqual(len(returns), 5)
        starred = sum(1 for n in returns if isinstance(n.value, ast.Dict)
                      and any(k is None for k in n.value.keys))
        self.assertEqual(starred, len(returns) - 1)


class TheRecordIsCompleteTests(unittest.TestCase):
    """What log_rows carries that verified_overrides cannot."""

    def test_an_approval_records_submitted_engine_and_approved_values(self):
        outcome = _run([{"id": "a1", "decision": "approve"}])
        row = outcome["log_rows"][0]
        self.assertEqual(row["gp"], "1000000001")
        self.assertEqual(row["field"], "Delinquency Status")
        self.assertIn("60 DAYS", row["original_values"])
        self.assertEqual(row["calculated_value"], "30 DAYS")
        self.assertEqual(row["approved_value"], "30 DAYS")

    def test_a_rejection_is_recorded_though_it_writes_nothing(self):
        outcome = _run([{"id": "a1", "decision": "reject"}])
        self.assertEqual(outcome["verified_overrides"], {})
        self.assertEqual(len(outcome["log_rows"]), 1)
        self.assertEqual(outcome["log_rows"][0]["decision"], "Rejected")
        self.assertIn("60 DAYS", outcome["log_rows"][0]["original_values"])

    def test_a_return_is_recorded_though_it_writes_nothing(self):
        outcome = _run([{"id": "a1", "decision": "return"}])
        self.assertEqual(outcome["verified_overrides"], {})
        self.assertEqual(outcome["log_rows"][0]["decision"], "Returned to servicer")

    def test_the_reviewer_note_survives(self):
        outcome = _run([{"id": "a1", "decision": "reject", "reviewer_note": "kept as filed"}])
        self.assertEqual(outcome["log_rows"][0]["reviewer_note"], "kept as filed")


class NonInterferenceTests(unittest.TestCase):
    """Nothing that existed before is changed by the new key."""

    def setUp(self):
        self.outcome = _run([{"id": "a1", "decision": "approve"}])

    def test_every_pre_existing_key_is_still_present(self):
        self.assertTrue(EXISTING_KEYS <= set(self.outcome))

    def test_log_rows_is_the_only_new_key(self):
        self.assertEqual(set(self.outcome) - EXISTING_KEYS, {"log_rows"})

    def test_verified_overrides_is_unchanged(self):
        self.assertEqual(self.outcome["verified_overrides"],
                         {("1000000001", "Delinquency Status"): "30 DAYS"})

    def test_the_counts_are_unchanged(self):
        self.assertEqual(self.outcome["applied_count"], 1)
        self.assertEqual(self.outcome["rejected_count"], 0)
        self.assertEqual(self.outcome["returned_count"], 0)
        self.assertEqual(self.outcome["approved_not_applied_count"], 0)

    def test_the_workbook_still_contains_the_correction_log(self):
        wb = openpyxl.load_workbook(io.BytesIO(self.outcome["file_bytes"]))
        self.assertIn(corrections.CORRECTION_LOG_SHEET_NAME, wb.sheetnames)

    def test_the_correction_log_still_holds_one_row_per_decision(self):
        wb = openpyxl.load_workbook(io.BytesIO(self.outcome["file_bytes"]))
        ws = wb[corrections.CORRECTION_LOG_SHEET_NAME]
        written = [r for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=1).value]
        self.assertEqual(len(written), len(self.outcome["log_rows"]))

    def test_the_corrected_value_was_still_written_to_the_sheet(self):
        wb = openpyxl.load_workbook(io.BytesIO(self.outcome["file_bytes"]))
        ws = wb["Servicer Inputs"]
        self.assertEqual(ws.cell(row=10, column=4).value, "30 DAYS")

    def test_exposing_the_list_cannot_alter_the_already_written_log(self):
        """The dict is constructed after wb.save(). Mutating the returned list
        afterwards must leave the serialised sheet untouched."""
        before = openpyxl.load_workbook(io.BytesIO(self.outcome["file_bytes"]))
        rows_before = before[corrections.CORRECTION_LOG_SHEET_NAME].max_row
        self.outcome["log_rows"].append({"gp": "tampered"})
        after = openpyxl.load_workbook(io.BytesIO(self.outcome["file_bytes"]))
        self.assertEqual(after[corrections.CORRECTION_LOG_SHEET_NAME].max_row, rows_before)

    def test_the_log_sheet_writer_only_reads_its_rows(self):
        source = inspect.getsource(corrections._add_correction_log_sheet)
        # No assignment into a row dict, and no in-place dict mutation.
        self.assertNotIn(".pop(", source)
        self.assertNotIn(".update(", source)
        self.assertNotRegex(source, r"\br\[[^\]]+\]\s*=")


class NoCallerCanSeeTheDifferenceTests(unittest.TestCase):
    """Every consumer reads the outcome by explicit named key."""

    def test_no_module_iterates_the_outcome_dict(self):
        import main
        for module in (main, corrections):
            source = inspect.getsource(module)
            for pattern in ("outcome.keys()", "**outcome", "for k in outcome",
                            "len(outcome)", "sorted(outcome"):
                with self.subTest(module=module.__name__, pattern=pattern):
                    self.assertNotIn(pattern, source)

    def test_the_consumer_reads_log_rows_by_explicit_named_key(self):
        """main.py IS the consumer now -- the mismatch report quotes the same
        review_reason the servicer's own Correction Log carries.

        The assertion this replaces ("nothing reads the new key yet") pinned
        Commit 1's additive-only claim and expired when that consumer arrived.
        The invariant that actually matters outlives it and is asserted here:
        the key is read by explicit name, exactly once, off `outcome` -- so the
        dict is still never iterated or unpacked (see the test above), and a
        future key still cannot be seen by any existing caller."""
        import main
        source = inspect.getsource(main)
        self.assertIn('outcome.get("log_rows")', source)
        self.assertEqual(source.count('outcome.get("log_rows")'), 1)


if __name__ == "__main__":
    unittest.main()
