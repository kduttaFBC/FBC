"""Step 2 — the Delinquency Information Request Form wired into
POST /api/corrections/review.

CLIENT-CONFIRMED (2026-08-18). The endpoint integration only: the gate, the
shared post-review frames, per-lender collection in upload order, the fail-safe
blocking policy, the minimal response surface, INDEPENDENCE FROM THE MISMATCH
REPORT, and non-interference with every artifact that already existed.

The form's own semantics -- eligibility, column sources, the receipt-date
mapping, formatting -- are Step 1's and live in test_delinquency_request_form.py.
Nothing here re-tests them.

DELIBERATELY NOT MOCKED: consolidate_cycle_postings, _apply_verified_overrides
and build_lender_rows. The point of Step 2 is that main.py builds the shared
frames once and feeds the real builder, so stubbing them would prove nothing.
"""

import asyncio
import base64
import io
import json
import unittest
from unittest import mock

import pandas as pd
from openpyxl import load_workbook

import checks
import delinquency_request_form as drf
import main


D2024 = "SOUP2024_1"
D2025 = "SOUP2025_1"
HEADER_ROW = 9


def _posting(gp, **over):
    row = {
        "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}",
        "Loan Payment Instance": 1,
        "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
        "Payment Received": 500.0, "Principal": 100.0, "Net Interest": 400.0,
        "Servicing Fee": 10.0, "Scheduled Payment": 500.0,
        "Beginning Balance": 100000.0, "Ending Balance": 99900.0,
        "Liquidation Date": None, "Liquidation Loss": None,
        "Payment Receipt Date": "2026-07-05",
        "_excel_row": HEADER_ROW + 1, "_sheet_name": "Servicer Inputs",
        "_header_excel_row": HEADER_ROW,
    }
    row.update(over)
    return row


def _cand(servicer, gp, field="Delinquency Status", reported="CURRENT",
          corrected="60 DAYS"):
    return {
        "id": f"{servicer}::{gp}::{field}", "gp": gp, "field": field,
        "corrected_value": corrected, "reported_value": reported,
        "excel_rows": [HEADER_ROW + 1], "skipped_row_refs": [],
        "sheet_name": "Servicer Inputs", "header_excel_row": HEADER_ROW,
        "eligibility_rule": "test", "review_required": False, "review_reason": None,
    }


class _Harness(unittest.TestCase):
    """Drives the real endpoint end to end."""

    def run_review(self, servicers, decisions=None, reviewer="Investigator",
                   read_failures=(), extra_patches=()):
        state = {}
        for name, spec in servicers.items():
            state[name] = {
                "original_bytes": f"BYTES::{name}".encode(),
                "original_filename": f"{name}.xlsx",
                "source_sha256": f"sha-{name}", "institution": name,
                "cycle_month": pd.Timestamp("2026-07-01"), "deal": spec["deal"],
                "candidates": list(spec.get("candidates") or []),
                "excluded_gps": list(spec.get("excluded_gps") or []),
            }

        if decisions is None:
            decisions = [{"id": c["id"], "decision": "approve",
                          "approved_value": c["corrected_value"]}
                         for spec in servicers.values()
                         for c in (spec.get("candidates") or [])]

        def fake_workbook(file_bytes, original_filename, institution, candidates,
                          decs, reviewer_name, source_sha256, cycle_month,
                          reviewed_on=None):
            by_id = {c["id"]: c for c in candidates}
            applied = rejected = returned = 0
            for d in decs:
                if d.get("id") not in by_id:
                    continue
                kind = d.get("decision")
                applied += kind == "approve"
                rejected += kind == "reject"
                returned += kind == "return"
            return {
                "file_bytes": b"CORRECTED" if decs else None,
                "filename": f"{institution}_CORRECTED.xlsx",
                "applied_count": applied, "rejected_count": rejected,
                "returned_count": returned, "approved_not_applied_count": 0,
                "corrections_not_applied": [], "unknown_ids": [],
                "skip_reason": None if decs else "nothing",
                "verified_overrides": dict(servicers[institution].get("overrides") or {})
                                      if applied else {},
                "log_rows": [],
            }

        def fake_read(byts, filename=None):
            name = byts.decode().split("::", 1)[1]
            if name in read_failures:
                raise ValueError(f"unreadable: {name}")
            return (name, name, pd.DataFrame(),
                    pd.DataFrame(servicers[name]["postings"]),
                    servicers[name]["deal"])

        def fake_summary(blocks, buffer, cycle_month=None, title=None,
                         blocked_deals=None):
            buffer.write(b"SUMMARY")

        sid = "delinquency-session"
        main.SESSION_CACHE[sid] = {"soup_df": None, "created": 9e18,
                                   "corrections": state}
        payload = {"session_id": sid, "reviewed_by": reviewer, "decisions": decisions}

        patches = [
            mock.patch.object(main.corrections_mod, "generate_reviewed_workbook",
                              fake_workbook),
            mock.patch.object(main.file_readers, "read_servicer_file", fake_read),
            mock.patch.object(main.msr_mod, "build_servicer_rows",
                              lambda si, sk, **kw: [{"SBA GP Number": "1000000001",
                                                     "Servicer Name": sk,
                                                     "Ending Balance": 99900.0}]),
            mock.patch.object(main.file_readers, "read_note_series_label",
                              lambda b: "2025 FBC-1"),
            mock.patch.object(main.file_readers, "read_remittance_date",
                              lambda b: pd.Timestamp("2026-08-23")),
            mock.patch.object(main.file_readers, "read_collection_period_calendar",
                              lambda b: None),
            mock.patch.object(main.msr_mod, "generate_monthly_servicer_report_bytes",
                              lambda r, buf: buf.write(b"MSR")),
            mock.patch.object(main.report_mod, "generate_final_portfolio_summary_bytes",
                              fake_summary),
        ]
        patches.extend(extra_patches)
        for p in patches:
            p.start()
        try:
            resp = asyncio.new_event_loop().run_until_complete(
                main.corrections_review(payload))
        finally:
            for p in reversed(patches):
                p.stop()
        body = json.loads(bytes(resp.body).decode()) if resp.body else {}
        return resp.status_code, body

    # -- helpers ----------------------------------------------------------
    @staticmethod
    def status(body):
        return body["delinquency_request_form_status"]

    @staticmethod
    def workbook(body):
        return load_workbook(io.BytesIO(
            base64.b64decode(body["delinquency_request_form_b64"])))

    def data_rows(self, body, sheet):
        ws = self.workbook(body)[sheet]
        out = []
        for r in range(drf.FIRST_DATA_ROW, ws.max_row + 1):
            first = ws.cell(row=r, column=drf.FIRST_COLUMN).value
            if first and first != drf.NO_QUALIFYING_LOANS_TEXT:
                out.append([ws.cell(row=r, column=c).value
                            for c in range(drf.FIRST_COLUMN, drf.FIRST_COLUMN + 7)])
        return out


def _lender(deal, gp, overrides=None, postings=None, candidates=None,
            excluded_gps=None, name=None):
    return {
        "deal": deal,
        "postings": postings or [_posting(gp)],
        "overrides": {} if overrides is None else overrides,
        "candidates": candidates if candidates is not None
                      else [_cand(name or deal, gp)],
        "excluded_gps": excluded_gps or [],
    }


ONE = {"CBT 2025": _lender(D2025, "1000000001", name="CBT 2025")}


class GateTests(_Harness):

    def test_an_incomplete_review_generates_nothing(self):
        spec = {"CBT 2025": _lender(D2025, "1000000001", name="CBT 2025")}
        spec["CBT 2025"]["candidates"].append(
            _cand("CBT 2025", "1000000002", field="Loan Status",
                  reported="ACTIVE", corrected="DEFERRED"))
        decisions = [{"id": spec["CBT 2025"]["candidates"][0]["id"],
                      "decision": "approve", "approved_value": "60 DAYS"}]
        _, body = self.run_review(spec, decisions=decisions)
        self.assertFalse(self.status(body)["generated"])
        self.assertIsNone(body["delinquency_request_form_b64"])
        self.assertIsNone(body["delinquency_request_form_filename"])
        self.assertIn("not yet complete", self.status(body)["blocked_reason"])
        self.assertIn("1 recommendation", self.status(body)["blocked_reason"])

    def test_a_complete_review_generates_the_form(self):
        _, body = self.run_review(ONE)
        self.assertTrue(self.status(body)["generated"])
        self.assertTrue(body["delinquency_request_form_b64"])
        self.assertIsNone(self.status(body)["blocked_reason"])
        self.assertEqual(body["delinquency_request_form_filename"],
                         "Delinquency_Information_Request_Form_July_2026.xlsx")

    def test_complete_but_not_reporting_ready_still_generates(self):
        """A returned item withholds the MSR and the Final Audited Portfolio
        Summary. It must NOT withhold this form."""
        spec = {"CBT 2025": _lender(D2025, "1000000001", name="CBT 2025")}
        spec["CBT 2025"]["candidates"].append(
            _cand("CBT 2025", "1000000002", field="Loan Status",
                  reported="ACTIVE", corrected="DEFERRED"))
        decisions = [
            {"id": spec["CBT 2025"]["candidates"][0]["id"], "decision": "approve",
             "approved_value": "60 DAYS"},
            {"id": spec["CBT 2025"]["candidates"][1]["id"], "decision": "return"},
        ]
        _, body = self.run_review(spec, decisions=decisions)
        self.assertTrue(body["review_completeness"]["review_complete"])
        self.assertFalse(body["reporting_readiness"]["reporting_ready"])
        self.assertIsNone(body["portfolio_summary_final_b64"])
        self.assertEqual(body["monthly_servicer_reports"], [])
        self.assertTrue(self.status(body)["generated"])

    def test_the_gate_reads_the_response_dict_not_a_second_local(self):
        import inspect
        src = inspect.getsource(main.corrections_review)
        self.assertEqual(src.count('review_completeness["review_complete"]'), 2)

    def test_no_payload_field_can_assert_the_gate(self):
        spec = {"CBT 2025": _lender(D2025, "1000000001", name="CBT 2025")}
        spec["CBT 2025"]["candidates"].append(
            _cand("CBT 2025", "1000000002", field="Loan Status"))
        decisions = [{"id": spec["CBT 2025"]["candidates"][0]["id"],
                      "decision": "approve", "approved_value": "60 DAYS"},
                     {"review_complete": True}]
        _, body = self.run_review(spec, decisions=decisions)
        self.assertFalse(self.status(body)["generated"])


class WiringTests(_Harness):
    """The shared final frame really drives eligibility."""

    def test_a_correction_brings_a_loan_into_the_form(self):
        gp = "1000000001"
        spec = {"CBT 2025": _lender(D2025, gp, name="CBT 2025",
                                    overrides={(gp, "Delinquency Status"): "60 DAYS"})}
        _, body = self.run_review(spec)
        rows = self.data_rows(body, "CBT 2025")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][0], gp)
        self.assertEqual(rows[0][4], "60 DAYS")

    def test_a_correction_out_of_scope_removes_a_loan(self):
        gp = "1000000001"
        spec = {"CBT 2025": _lender(
            D2025, gp, name="CBT 2025",
            postings=[_posting(gp, **{"Delinquency Status": "60 DAYS"})],
            overrides={(gp, "Delinquency Status"): "CURRENT"})}
        _, body = self.run_review(spec)
        self.assertEqual(self.data_rows(body, "CBT 2025"), [])
        self.assertEqual(self.status(body)["loan_count"], 0)

    def test_a_loan_status_correction_brings_a_loan_in(self):
        gp = "1000000001"
        spec = {"CBT 2025": _lender(D2025, gp, name="CBT 2025",
                                    overrides={(gp, "Loan Status"): "DEFERRED"})}
        _, body = self.run_review(spec)
        rows = self.data_rows(body, "CBT 2025")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][3], "DEFERRED")

    def test_the_balance_is_the_final_overridden_value(self):
        gp = "1000000001"
        spec = {"CBT 2025": _lender(
            D2025, gp, name="CBT 2025",
            overrides={(gp, "Loan Status"): "REO", (gp, "Ending Balance"): 145729.87})}
        _, body = self.run_review(spec)
        self.assertAlmostEqual(self.data_rows(body, "CBT 2025")[0][2], 145729.87, 2)

    def test_excluded_gps_suppression_is_applied(self):
        gp = "1000000001"
        spec = {"CBT 2025": _lender(D2025, gp, name="CBT 2025",
                                    overrides={(gp, "Loan Status"): "REO"},
                                    excluded_gps=[gp])}
        _, body = self.run_review(spec)
        self.assertEqual(self.data_rows(body, "CBT 2025"), [])

    def test_the_receipt_date_survives_end_to_end(self):
        gp = "1000000001"
        spec = {"CBT 2025": _lender(D2025, gp, name="CBT 2025",
                                    overrides={(gp, "Loan Status"): "REO"})}
        _, body = self.run_review(spec)
        self.assertEqual(str(self.data_rows(body, "CBT 2025")[0][5])[:10], "2026-07-05")

    def test_the_balance_heading_carries_the_reporting_month_end(self):
        _, body = self.run_review(ONE)
        ws = self.workbook(body)["CBT 2025"]
        self.assertEqual(ws.cell(row=drf.HEADER_ROW, column=drf.FIRST_COLUMN + 2).value,
                         "Securitization Balance as of 7/31/26")


class StructureTests(_Harness):

    ROSTER = {
        "Community Bank & Trust 2025": _lender(D2025, "1000000001",
                                               name="Community Bank & Trust 2025"),
        "Lendistry 2024": _lender(D2024, "2000000002", name="Lendistry 2024"),
        "Lendistry 2025": _lender(D2025, "3000000003", name="Lendistry 2025"),
        "Millennium 2024": _lender(D2024, "4000000004", name="Millennium 2024"),
        "Millennium 2025": _lender(D2025, "5000000005", name="Millennium 2025"),
        "Mission Valley Bank": _lender(D2024, "6000000006",
                                       name="Mission Valley Bank"),
        "Peoples Bank": _lender(D2024, "7000000007", name="Peoples Bank"),
        "Poppy Bank": _lender(D2024, "8000000008", name="Poppy Bank"),
    }

    def test_eight_lender_and_deal_tabs_in_upload_order(self):
        _, body = self.run_review(self.ROSTER)
        self.assertEqual(self.workbook(body).sheetnames, list(self.ROSTER))
        self.assertEqual(self.status(body)["lender_count"], 8)

    def test_one_lender_in_two_deals_keeps_two_separate_sheets(self):
        gp_a, gp_b = "2000000002", "3000000003"
        spec = {
            "Lendistry 2024": _lender(D2024, gp_a, name="Lendistry 2024",
                                      overrides={(gp_a, "Loan Status"): "REO"}),
            "Lendistry 2025": _lender(D2025, gp_b, name="Lendistry 2025",
                                      overrides={(gp_b, "Loan Status"): "REO"}),
        }
        _, body = self.run_review(spec)
        self.assertEqual([r[0] for r in self.data_rows(body, "Lendistry 2024")], [gp_a])
        self.assertEqual([r[0] for r in self.data_rows(body, "Lendistry 2025")], [gp_b])

    def test_a_lender_with_no_qualifying_loans_keeps_its_sheet(self):
        """A clean lender is told so, never omitted."""
        _, body = self.run_review(ONE)
        ws = self.workbook(body)["CBT 2025"]
        self.assertEqual(ws.cell(row=drf.FIRST_DATA_ROW,
                                 column=drf.FIRST_COLUMN).value,
                         "No loans requiring additional information.")
        self.assertEqual(self.status(body)["loan_count"], 0)
        self.assertEqual(self.status(body)["lender_count"], 1)

    def test_loan_count_is_the_total_written_across_every_sheet(self):
        gp_a, gp_b = "2000000002", "3000000003"
        spec = {
            "Lendistry 2024": _lender(D2024, gp_a, name="Lendistry 2024",
                                      overrides={(gp_a, "Loan Status"): "REO"}),
            "Lendistry 2025": _lender(D2025, gp_b, name="Lendistry 2025",
                                      overrides={(gp_b, "Loan Status"): "REO"}),
        }
        _, body = self.run_review(spec)
        self.assertEqual(self.status(body)["loan_count"], 2)
        self.assertEqual(self.status(body)["lender_count"], 2)

    def test_the_notes_column_is_blank_on_every_generated_row(self):
        """THE invariant. It is the lender's column to complete."""
        spec = {name: _lender(v["deal"], f"{i+1}00000000{i+1}", name=name,
                              overrides={(f"{i+1}00000000{i+1}", "Loan Status"): "REO"},
                              postings=[_posting(f"{i+1}00000000{i+1}")])
                for i, (name, v) in enumerate(self.ROSTER.items())}
        _, body = self.run_review(spec)
        wb = self.workbook(body)
        written = 0
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for r in range(drf.FIRST_DATA_ROW, ws.max_row + 1):
                if ws.cell(row=r, column=drf.FIRST_COLUMN).value:
                    written += 1
                    self.assertIsNone(
                        ws.cell(row=r, column=drf.FIRST_COLUMN + 6).value,
                        f"{sheet} row {r}")
        self.assertEqual(written, 8)


class BlockingPolicyTests(_Harness):
    """This workbook leaves the building addressed to lenders."""

    TWO = {"CBT 2025": _lender(D2025, "1000000001", name="CBT 2025"),
           "Lendistry 2025": _lender(D2025, "2000000002", name="Lendistry 2025")}

    def test_an_unreadable_servicer_blocks_and_is_named(self):
        _, body = self.run_review(dict(self.TWO), read_failures={"Lendistry 2025"})
        self.assertFalse(self.status(body)["generated"])
        self.assertIsNone(body["delinquency_request_form_b64"])
        self.assertIn("Lendistry 2025", self.status(body)["blocked_reason"])
        self.assertIn("could not be read", self.status(body)["blocked_reason"])

    def test_an_evaluation_failure_blocks_with_distinct_wording(self):
        boom = mock.patch.object(main.drf_mod, "build_lender_rows",
                                 mock.Mock(side_effect=RuntimeError("boom")))
        _, body = self.run_review(dict(self.TWO), extra_patches=[boom])
        reason = self.status(body)["blocked_reason"]
        self.assertFalse(self.status(body)["generated"])
        self.assertIn("could not be evaluated for additional-information requests",
                      reason)
        self.assertNotIn("could not be read", reason)

    def test_a_blocked_reason_is_never_absent_when_generation_failed(self):
        _, body = self.run_review(dict(self.TWO), read_failures={"Lendistry 2025"})
        self.assertFalse(self.status(body)["generated"])
        self.assertTrue(self.status(body)["blocked_reason"])

    def test_generated_and_bytes_never_disagree(self):
        for spec, kw in ((ONE, {}),
                         (dict(self.TWO), {"read_failures": {"Lendistry 2025"}})):
            _, body = self.run_review(spec, **kw)
            self.assertEqual(self.status(body)["generated"],
                             body["delinquency_request_form_b64"] is not None)

    def test_no_sba_gp_number_appears_in_a_blocked_reason(self):
        gp = "2000000002"
        _, body = self.run_review(dict(self.TWO), read_failures={"Lendistry 2025"})
        self.assertNotIn(gp, self.status(body)["blocked_reason"])


class IndependenceTests(_Harness):
    """Either report must be able to fail without stopping the other."""

    def test_a_mismatch_failure_leaves_the_form_generated(self):
        gp = "1000000001"
        spec = {"CBT 2025": _lender(D2025, gp, name="CBT 2025",
                                    overrides={(gp, "Loan Status"): "REO"})}
        boom = mock.patch.object(main.mismatch_mod, "build_mismatch_rows",
                                 mock.Mock(side_effect=RuntimeError("boom")))
        _, body = self.run_review(spec, extra_patches=[boom])
        self.assertFalse(body["mismatch_report_status"]["generated"])
        self.assertTrue(self.status(body)["generated"])
        self.assertEqual(self.status(body)["loan_count"], 1)

    def test_a_form_failure_leaves_the_mismatch_report_generated(self):
        gp = "1000000001"
        spec = {"CBT 2025": _lender(D2025, gp, name="CBT 2025",
                                    overrides={(gp, "Loan Status"): "REO"})}
        boom = mock.patch.object(main.drf_mod, "build_lender_rows",
                                 mock.Mock(side_effect=RuntimeError("boom")))
        _, body = self.run_review(spec, extra_patches=[boom])
        self.assertFalse(self.status(body)["generated"])
        self.assertTrue(body["mismatch_report_status"]["generated"])

    def test_neither_status_object_references_the_other(self):
        _, body = self.run_review(ONE)
        self.assertEqual(set(self.status(body)),
                         {"generated", "blocked_reason", "loan_count", "lender_count"})
        self.assertEqual(set(body["mismatch_report_status"]),
                         {"generated", "blocked_reason", "mismatch_count",
                          "provenance_gap_count"})

    def test_the_form_generation_block_reads_no_mismatch_state(self):
        """Bounded to the form's OWN generation block. Slicing to the end of the
        function instead would sweep in the Final Audited Portfolio Summary and
        the response dict, where mismatch names legitimately appear."""
        import inspect
        src = inspect.getsource(main.corrections_review)
        start = src.index("---- Delinquency Information Request Form ---",
                          src.index("_complete_deals = "))
        end = src.index("---- Final Audited Portfolio Summary", start)
        block = src[start:end]
        self.assertIn("delinquency_generated = True", block)
        for leaked in ("mismatch_servicers", "mismatch_provenance_gaps",
                       "mismatch_blocked_reason", "mismatch_rows_by_servicer",
                       "mismatch_generated", "mismatch_report_b64"):
            with self.subTest(name=leaked):
                self.assertNotIn(leaked, block)

    def test_the_two_reports_use_separate_accumulators(self):
        import inspect
        src = inspect.getsource(main.corrections_review)
        for pair in (("mismatch_rows_by_servicer", "delinquency_blocks"),
                     ("mismatch_servicers", "delinquency_servicers"),
                     ("mismatch_failures", "delinquency_failures")):
            with self.subTest(pair=pair):
                for name in pair:
                    self.assertIn(f"{name} = []", src)

    def test_the_shared_frames_are_built_once_and_owned_by_neither(self):
        import inspect
        src = inspect.getsource(main.corrections_review)
        self.assertEqual(src.count("report_submitted = checks_mod.consolidate"), 1)
        self.assertEqual(src.count("report_final = checks_mod._apply_verified"), 1)
        self.assertNotIn("mismatch_submitted", src)
        self.assertNotIn("mismatch_final", src)


class NonInterferenceTests(_Harness):
    """Everything that existed before Step 2 is byte-for-byte unchanged."""

    SPEC = {"CBT 2025": _lender(D2025, "1000000001", name="CBT 2025",
                                overrides={("1000000001", "Loan Status"): "REO"}),
            "Millennium 2024": _lender(D2024, "3000000003",
                                       name="Millennium 2024")}

    def setUp(self):
        _, self.normal = self.run_review(dict(self.SPEC))
        boom = mock.patch.object(main.drf_mod, "build_lender_rows",
                                 mock.Mock(side_effect=RuntimeError("boom")))
        _, self.broken = self.run_review(dict(self.SPEC), extra_patches=[boom])

    def test_the_mismatch_report_bytes_are_identical(self):
        self.assertEqual(self.normal["mismatch_report_b64"],
                         self.broken["mismatch_report_b64"])
        self.assertIsNotNone(self.normal["mismatch_report_b64"])
        self.assertEqual(self.normal["mismatch_report_status"],
                         self.broken["mismatch_report_status"])

    def test_corrected_workbook_bytes_are_identical(self):
        self.assertEqual(self.normal["corrected_workbooks"],
                         self.broken["corrected_workbooks"])
        self.assertTrue(self.normal["corrected_workbooks"])

    def test_monthly_servicer_reports_are_identical(self):
        self.assertEqual(self.normal["monthly_servicer_reports"],
                         self.broken["monthly_servicer_reports"])
        self.assertTrue(self.normal["monthly_servicer_reports"])

    def test_the_final_audited_portfolio_summary_is_identical(self):
        self.assertEqual(self.normal["portfolio_summary_final_b64"],
                         self.broken["portfolio_summary_final_b64"])
        self.assertIsNotNone(self.normal["portfolio_summary_final_b64"])

    def test_review_completeness_and_reporting_readiness_are_identical(self):
        self.assertEqual(self.normal["review_completeness"],
                         self.broken["review_completeness"])
        self.assertEqual(self.normal["reporting_readiness"],
                         self.broken["reporting_readiness"])

    def test_every_other_response_key_is_identical(self):
        for key in ("deals", "status", "summary", "corrections_not_applied",
                    "unknown_ids", "monthly_servicer_report_validation_issues"):
            self.assertEqual(self.normal[key], self.broken[key], key)

    def test_a_form_failure_still_returns_http_200_with_the_artifacts(self):
        boom = mock.patch.object(main.drf_mod, "build_lender_rows",
                                 mock.Mock(side_effect=RuntimeError("boom")))
        code, body = self.run_review(dict(self.SPEC), extra_patches=[boom])
        self.assertEqual(code, 200)
        self.assertTrue(body["corrected_workbooks"])
        self.assertTrue(body["monthly_servicer_reports"])
        self.assertIsNotNone(body["portfolio_summary_final_b64"])

    def test_neither_shared_frame_is_mutated(self):
        seen = []
        real = checks.consolidate_cycle_postings

        def spy(df):
            out = real(df)
            seen.append((out, out.copy(deep=True)))
            return out

        with mock.patch.object(main.checks_mod, "consolidate_cycle_postings", spy):
            self.run_review(dict(self.SPEC))
        self.assertTrue(seen)
        for after, before in seen:
            pd.testing.assert_frame_equal(after, before)

    def test_the_verified_override_map_is_not_mutated(self):
        gp = "1000000001"
        overrides = {(gp, "Loan Status"): "REO"}
        spec = {"CBT 2025": _lender(D2025, gp, name="CBT 2025", overrides=overrides)}
        self.run_review(spec)
        self.assertEqual(overrides, {(gp, "Loan Status"): "REO"})


class ResponseSurfaceTests(_Harness):

    PRE_EXISTING = {
        "corrected_workbooks", "monthly_servicer_reports",
        "monthly_servicer_report_validation_issues", "status", "deals",
        "review_completeness", "reporting_readiness",
        "portfolio_summary_final_b64", "portfolio_summary_final_filename",
        "summary", "corrections_not_applied", "unknown_ids",
        "mismatch_report_b64", "mismatch_report_filename", "mismatch_report_status",
    }
    NEW_KEYS = {"delinquency_request_form_b64",
                "delinquency_request_form_filename",
                "delinquency_request_form_status"}

    def setUp(self):
        _, self.body = self.run_review(ONE)

    def test_every_pre_existing_key_survives(self):
        self.assertTrue(self.PRE_EXISTING <= set(self.body))

    def test_exactly_three_keys_were_added(self):
        self.assertEqual(set(self.body) - self.PRE_EXISTING, self.NEW_KEYS)

    def test_the_status_object_carries_exactly_four_fields(self):
        self.assertEqual(set(self.status(self.body)),
                         {"generated", "blocked_reason", "loan_count",
                          "lender_count"})

    def test_no_loan_level_diagnostic_array_is_exposed(self):
        for key in ("delinquency_blocks", "delinquency_failures",
                    "delinquency_servicers", "rows"):
            self.assertNotIn(key, self.body)
            self.assertNotIn(key, self.status(self.body))

    def test_the_whole_response_is_json_serialisable(self):
        json.dumps(self.body)


class BoundaryTests(unittest.TestCase):
    """Step 2 is backend-only and touches nothing it was told not to."""

    def test_main_imports_the_form_module(self):
        import inspect
        self.assertIn("import delinquency_request_form as drf_mod",
                      inspect.getsource(main))

    def test_the_form_module_still_imports_no_main(self):
        import ast
        import inspect
        names = set()
        for node in ast.walk(ast.parse(inspect.getsource(drf))):
            if isinstance(node, ast.Import):
                names.update(a.name for a in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                names.add(node.module)
        self.assertNotIn("main", names)

    def test_main_reaches_the_module_only_through_public_entry_points(self):
        import inspect
        import re
        used = set(re.findall(r"drf_mod\.(\w+)", inspect.getsource(main)))
        self.assertTrue(used)
        self.assertEqual(used - {"build_lender_rows",
                                 "generate_delinquency_request_form_bytes",
                                 "delinquency_request_form_filename"}, set())
        self.assertEqual([n for n in used if n.startswith("_")], [])


if __name__ == "__main__":
    unittest.main()
