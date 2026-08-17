"""Commit 3 — the mismatch report wired into POST /api/corrections/review.

WHAT THIS COVERS
----------------
CLIENT-CONFIRMED (2026-08-17). The endpoint integration only: the gate, the two
frames, multi-servicer collection, the fail-safe blocking policy, the minimal
response surface, and non-interference with every artifact that already existed.

The report's own semantics -- detection, normalisation, comments, worksheet
structure -- are Commit 2's and live in test_mismatch_report.py. Nothing here
re-tests them.

THE GATE IS review_complete, NOT reporting_ready. This report states no figure
of its own, only what was submitted and what FBC determined, and both are known
once every recommendation has a decision. So "complete but NOT ready" -- which
withholds the Monthly Servicer Report and the Final Audited Portfolio Summary --
still produces this one. That asymmetry is asserted directly, twice.

DELIBERATELY NOT MOCKED: consolidate_cycle_postings and
_apply_verified_overrides. The whole point of Commit 3 is that main.py now calls
those two real functions, so a test that stubbed them would prove nothing about
the wiring.
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
import main
import mismatch_report as mr


D2024 = "SOUP2024_1"
D2025 = "SOUP2025_1"
HEADER_ROW = 9

SHEET_2024 = mr.WORKSHEET_2024
SHEET_2025 = mr.WORKSHEET_2025


def _posting(gp, **over):
    row = {
        "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}",
        "Loan Payment Instance": 1,
        "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
        "Payment Received": 500.0, "Principal": 100.0, "Net Interest": 400.0,
        "Servicing Fee": 10.0, "Scheduled Payment": 500.0,
        "Beginning Balance": 100000.0, "Ending Balance": 99900.0,
        "Liquidation Date": None, "Liquidation Loss": None,
        "_excel_row": HEADER_ROW + 1, "_sheet_name": "Servicer Inputs",
        "_header_excel_row": HEADER_ROW,
    }
    row.update(over)
    return row


def _cand(servicer, gp, field="Delinquency Status", reported="CURRENT",
          corrected="30 DAYS", review_reason=None):
    return {
        "id": f"{servicer}::{gp}::{field}", "gp": gp, "field": field,
        "corrected_value": corrected, "reported_value": reported,
        "excel_rows": [HEADER_ROW + 1], "skipped_row_refs": [],
        "sheet_name": "Servicer Inputs", "header_excel_row": HEADER_ROW,
        "eligibility_rule": "test", "review_required": False,
        "review_reason": review_reason,
    }


class _Harness(unittest.TestCase):
    """Drives the real endpoint end to end.

    generate_reviewed_workbook is stubbed so a decision's OUTCOME can be chosen
    per candidate (approved-but-not-applied is a property of the write attempt,
    not of the payload) and so verified_overrides / log_rows can be handed back
    exactly as the real one would. Everything downstream of it -- consolidation,
    override application, grouping, workbook writing -- is the real code."""

    def run_review(self, servicers, decisions=None, reviewer="Investigator",
                   read_failures=(), extra_patches=(), log_rows_for=None):
        """servicers: {name: {"deal":, "postings": [...], "overrides": {...},
                              "candidates": [...]}}"""
        state = {}
        for name, spec in servicers.items():
            state[name] = {
                "original_bytes": f"BYTES::{name}".encode(),
                "original_filename": f"{name}.xlsx",
                "source_sha256": f"sha-{name}", "institution": name,
                "cycle_month": pd.Timestamp("2026-06-01"), "deal": spec["deal"],
                "candidates": list(spec.get("candidates") or []),
                # Non-empty on purpose in one test: D2 says this must NOT
                # suppress a mismatch.
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
            log_rows = []
            for d in decs:
                cand = by_id.get(d.get("id"))
                if cand is None:
                    continue
                kind = d.get("decision")
                if kind == "approve":
                    applied += 1
                elif kind == "reject":
                    rejected += 1
                elif kind == "return":
                    returned += 1
                log_rows.append({
                    "gp": cand["gp"], "field": cand["field"],
                    "original_values": cand["reported_value"],
                    "calculated_value": cand["corrected_value"],
                    "approved_value": d.get("approved_value"),
                    "decision": kind,
                    "review_reason": (log_rows_for or {}).get(cand["id"]),
                    "reviewer_note": "INTERNAL NOTE MUST NOT APPEAR",
                })
            return {
                "file_bytes": b"CORRECTED" if decs else None,
                "filename": f"{institution}_CORRECTED.xlsx",
                "applied_count": applied, "rejected_count": rejected,
                "returned_count": returned, "approved_not_applied_count": 0,
                "corrections_not_applied": [], "unknown_ids": [],
                "skip_reason": None if decs else "nothing",
                # Only APPROVALS reach the verified layer, exactly as the real
                # function behaves: a rejection writes nothing.
                "verified_overrides": dict(servicers[institution].get("overrides") or {})
                                      if applied else {},
                "log_rows": log_rows,
            }

        def fake_read(byts, filename=None):
            name = byts.decode().split("::", 1)[1]
            if name in read_failures:
                raise ValueError(f"unreadable: {name}")
            return (name, name, pd.DataFrame(),
                    pd.DataFrame(servicers[name]["postings"]),
                    servicers[name]["deal"])

        captured = {}

        def fake_summary(blocks, buffer, cycle_month=None, title=None, blocked_deals=None):
            captured["summary"] = True
            buffer.write(b"SUMMARY")

        def fake_msr(deal_report, buffer):
            captured.setdefault("msr", []).append(deal_report)
            buffer.write(b"MSR")

        sid = "mismatch-session"
        main.SESSION_CACHE[sid] = {"soup_df": None, "created": 9e18, "corrections": state}
        payload = {"session_id": sid, "reviewed_by": reviewer, "decisions": decisions}

        patches = [
            mock.patch.object(main.corrections_mod, "generate_reviewed_workbook", fake_workbook),
            mock.patch.object(main.file_readers, "read_servicer_file", fake_read),
            mock.patch.object(main.msr_mod, "build_servicer_rows",
                              lambda si, sk, **kw: [{"SBA GP Number": "1000000001",
                                                     "Servicer Name": sk,
                                                     "Ending Balance": 99900.0}]),
            mock.patch.object(main.file_readers, "read_note_series_label", lambda b: "2025 FBC-1"),
            mock.patch.object(main.file_readers, "read_remittance_date",
                              lambda b: pd.Timestamp("2026-07-23")),
            mock.patch.object(main.file_readers, "read_collection_period_calendar", lambda b: None),
            mock.patch.object(main.msr_mod, "generate_monthly_servicer_report_bytes", fake_msr),
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
        return resp.status_code, body, captured

    # -- helpers ----------------------------------------------------------
    @staticmethod
    def status(body):
        return body["mismatch_report_status"]

    @staticmethod
    def sheets(body):
        wb = load_workbook(io.BytesIO(base64.b64decode(body["mismatch_report_b64"])))
        out = {}
        for name in wb.sheetnames:
            ws = wb[name]
            out[name] = [[c.value for c in row] for row in ws.iter_rows()]
        return out

    def data_rows(self, body, sheet):
        rows = self.sheets(body)[sheet][1:]
        return [r for r in rows if r[0] and r[0] != mr.NO_MISMATCHES_TEXT]


# One servicer per deal, one loan each, one approvable Delinquency Status.
def _one(deal, gp, overrides=None, field="Delinquency Status",
         corrected="30 DAYS", postings=None, name=None):
    return {
        "deal": deal,
        "postings": postings or [_posting(gp)],
        "overrides": overrides if overrides is not None
                     else {(gp, field): corrected},
        "candidates": [_cand(name or deal, gp, field=field, corrected=corrected)],
    }


class GateTests(_Harness):
    """review_complete, and nothing else, decides whether the report exists."""

    def test_an_incomplete_review_generates_no_report(self):
        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")}
        spec["CBT 2025"]["candidates"].append(
            _cand("CBT 2025", "1000000001", field="Loan Status",
                  reported="ACTIVE", corrected="DEFERRED"))
        # Decide only the first of the two.
        decisions = [{"id": spec["CBT 2025"]["candidates"][0]["id"],
                      "decision": "approve", "approved_value": "30 DAYS"}]
        _, body, _ = self.run_review(spec, decisions=decisions)
        self.assertFalse(self.status(body)["generated"])
        self.assertIsNone(body["mismatch_report_b64"])
        self.assertIsNone(body["mismatch_report_filename"])
        self.assertIn("not yet complete", self.status(body)["blocked_reason"])

    def test_the_blocked_reason_states_how_many_are_pending(self):
        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")}
        spec["CBT 2025"]["candidates"].append(
            _cand("CBT 2025", "1000000002", field="Loan Status"))
        decisions = [{"id": spec["CBT 2025"]["candidates"][0]["id"],
                      "decision": "approve", "approved_value": "30 DAYS"}]
        _, body, _ = self.run_review(spec, decisions=decisions)
        self.assertEqual(body["review_completeness"]["pending"], 1)
        self.assertIn("1 recommendation", self.status(body)["blocked_reason"])

    def test_a_complete_review_generates_the_report(self):
        _, body, _ = self.run_review({"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")})
        self.assertTrue(self.status(body)["generated"])
        self.assertTrue(body["mismatch_report_b64"])
        self.assertIsNone(self.status(body)["blocked_reason"])

    def test_complete_but_not_reporting_ready_still_generates(self):
        """THE ASYMMETRY. A returned item blocks the MSR and the Final Audited
        Portfolio Summary and must NOT block this report."""
        gp = "1000000001"
        spec = {"CBT 2025": _one(D2025, gp, name="CBT 2025")}
        spec["CBT 2025"]["candidates"].append(
            _cand("CBT 2025", "1000000002", field="Loan Status",
                  reported="ACTIVE", corrected="DEFERRED"))
        decisions = [
            {"id": spec["CBT 2025"]["candidates"][0]["id"], "decision": "approve",
             "approved_value": "30 DAYS"},
            {"id": spec["CBT 2025"]["candidates"][1]["id"], "decision": "return"},
        ]
        _, body, cap = self.run_review(spec, decisions=decisions)
        self.assertTrue(body["review_completeness"]["review_complete"])
        self.assertFalse(body["reporting_readiness"]["reporting_ready"])
        # Withheld ...
        self.assertIsNone(body["portfolio_summary_final_b64"])
        self.assertEqual(body["monthly_servicer_reports"], [])
        # ... and this one is not.
        self.assertTrue(self.status(body)["generated"])
        self.assertTrue(body["mismatch_report_b64"])

    def test_the_gate_reads_the_response_dict_not_a_second_local(self):
        import inspect
        src = inspect.getsource(main.corrections_review)
        self.assertIn('review_completeness["review_complete"]', src)

    def test_no_payload_field_can_assert_the_gate(self):
        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")}
        spec["CBT 2025"]["candidates"].append(
            _cand("CBT 2025", "1000000002", field="Loan Status"))
        decisions = [{"id": spec["CBT 2025"]["candidates"][0]["id"],
                      "decision": "approve", "approved_value": "30 DAYS"}]
        _, body, _ = self.run_review(
            {**spec}, decisions=decisions + [{"review_complete": True}])
        self.assertFalse(self.status(body)["generated"])


class WiringTests(_Harness):
    """The two frames really are consolidate + _apply_verified_overrides."""

    def test_an_approved_correction_becomes_a_row(self):
        gp = "1000000001"
        _, body, _ = self.run_review({"CBT 2025": _one(D2025, gp, name="CBT 2025")})
        rows = self.data_rows(body, SHEET_2025)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][1], gp)                 # SBA GP Number
        self.assertEqual(rows[0][3], "CURRENT")          # Submitted Information
        self.assertEqual(rows[0][4], "30 DAYS")          # FBC Determination
        self.assertEqual(rows[0][5], "Delinquency Status")

    def test_the_lender_column_is_the_soup_key(self):
        _, body, _ = self.run_review({"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")})
        self.assertEqual(self.data_rows(body, SHEET_2025)[0][0], "CBT 2025")

    def test_a_rejected_recommendation_produces_no_row(self):
        gp = "1000000001"
        spec = {"CBT 2025": _one(D2025, gp, name="CBT 2025")}
        decisions = [{"id": spec["CBT 2025"]["candidates"][0]["id"], "decision": "reject"}]
        _, body, _ = self.run_review(spec, decisions=decisions)
        self.assertTrue(self.status(body)["generated"])
        self.assertEqual(self.status(body)["mismatch_count"], 0)
        self.assertEqual(self.data_rows(body, SHEET_2025), [])

    def test_the_submitted_column_is_the_consolidated_value_not_a_raw_posting(self):
        """Two postings in one cycle: the LAST status is what was submitted."""
        gp = "1000000001"
        postings = [_posting(gp, **{"Loan Payment Instance": 1,
                                    "Delinquency Status": "CURRENT"}),
                    _posting(gp, **{"Loan Payment Instance": 2,
                                    "Delinquency Status": "60 DAYS"})]
        spec = {"CBT 2025": _one(D2025, gp, postings=postings, name="CBT 2025")}
        _, body, _ = self.run_review(spec)
        self.assertEqual(self.data_rows(body, SHEET_2025)[0][3], "60 DAYS")

    def test_the_review_reason_reaches_the_comments(self):
        gp = "1000000001"
        spec = {"CBT 2025": _one(D2025, gp, name="CBT 2025")}
        cid = spec["CBT 2025"]["candidates"][0]["id"]
        _, body, _ = self.run_review(spec, log_rows_for={cid: "Deferral confirmed by memo"})
        self.assertIn("Deferral confirmed by memo",
                      self.data_rows(body, SHEET_2025)[0][6])

    def test_the_reviewer_note_never_reaches_the_comments(self):
        _, body, _ = self.run_review({"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")})
        self.assertNotIn("INTERNAL NOTE", self.data_rows(body, SHEET_2025)[0][6])

    def test_a_servicer_with_no_decisions_contributes_nothing(self):
        gp_a, gp_b = "1000000001", "2000000002"
        spec = {"CBT 2025": _one(D2025, gp_a, name="CBT 2025"),
                "Lendistry 2025": {"deal": D2025, "postings": [_posting(gp_b)],
                                   "overrides": {}, "candidates": []}}
        _, body, _ = self.run_review(spec)
        rows = self.data_rows(body, SHEET_2025)
        self.assertEqual([r[0] for r in rows], ["CBT 2025"])
        self.assertEqual(self.status(body)["provenance_gap_count"], 0)

    def test_excluded_gps_does_not_suppress_a_mismatch(self):
        """D2 (CLIENT-CONFIRMED 2026-08-17). Closed-at-zero suppression is a
        portfolio measurement rule; the corrected workbook still carries the
        change, so the mismatch stays visible."""
        gp = "1000000001"
        spec = {"CBT 2025": _one(D2025, gp, name="CBT 2025")}
        spec["CBT 2025"]["excluded_gps"] = [gp]
        _, body, _ = self.run_review(spec)
        self.assertEqual(len(self.data_rows(body, SHEET_2025)), 1)


class MultiServicerTests(_Harness):
    """Collection across servicers and deals."""

    def test_two_servicers_on_one_deal_share_a_worksheet_in_upload_order(self):
        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025"),
                "Lendistry 2025": _one(D2025, "2000000002", name="Lendistry 2025")}
        _, body, _ = self.run_review(spec)
        rows = self.data_rows(body, SHEET_2025)
        self.assertEqual([r[0] for r in rows], ["CBT 2025", "Lendistry 2025"])
        self.assertEqual(self.data_rows(body, SHEET_2024), [])

    def test_the_two_deals_are_kept_apart(self):
        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025"),
                "Millennium 2024": _one(D2024, "3000000003", name="Millennium 2024")}
        _, body, _ = self.run_review(spec)
        self.assertEqual([r[0] for r in self.data_rows(body, SHEET_2025)], ["CBT 2025"])
        self.assertEqual([r[0] for r in self.data_rows(body, SHEET_2024)], ["Millennium 2024"])

    def test_both_worksheets_always_exist(self):
        _, body, _ = self.run_review({"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")})
        self.assertEqual(list(self.sheets(body)), [SHEET_2024, SHEET_2025])

    def test_mismatch_count_is_the_total_across_both_worksheets(self):
        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025"),
                "Millennium 2024": _one(D2024, "3000000003", name="Millennium 2024")}
        _, body, _ = self.run_review(spec)
        self.assertEqual(self.status(body)["mismatch_count"], 2)

    def test_a_clean_month_still_produces_the_workbook(self):
        """Zero mismatches is a FINDING and must be stated, not implied by an
        absent file."""
        gp = "1000000001"
        spec = {"CBT 2025": _one(D2025, gp, name="CBT 2025")}
        decisions = [{"id": spec["CBT 2025"]["candidates"][0]["id"], "decision": "reject"}]
        _, body, _ = self.run_review(spec, decisions=decisions)
        sheets = self.sheets(body)
        self.assertTrue(self.status(body)["generated"])
        self.assertEqual(self.status(body)["mismatch_count"], 0)
        for name in (SHEET_2024, SHEET_2025):
            self.assertEqual(sheets[name][1][0], mr.NO_MISMATCHES_TEXT)

    def test_the_filename_carries_the_reporting_period(self):
        _, body, _ = self.run_review({"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")})
        self.assertEqual(body["mismatch_report_filename"],
                         "Submitted_vs_FBC_Mismatch_Report_June_2026.xlsx")


class BlockingPolicyTests(_Harness):
    """D3 fail-safe: block only where a real mismatch would be omitted."""

    def test_an_unreadable_servicer_blocks_and_is_named(self):
        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025"),
                "Lendistry 2025": _one(D2025, "2000000002", name="Lendistry 2025")}
        _, body, _ = self.run_review(spec, read_failures={"Lendistry 2025"})
        self.assertFalse(self.status(body)["generated"])
        self.assertIsNone(body["mismatch_report_b64"])
        self.assertIn("Lendistry 2025", self.status(body)["blocked_reason"])
        self.assertIn("could not be read", self.status(body)["blocked_reason"])

    def test_a_failure_building_rows_also_blocks_rather_than_shipping_short(self):
        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")}
        boom = mock.patch.object(
            main.mismatch_mod, "build_mismatch_rows",
            mock.Mock(side_effect=RuntimeError("boom")))
        _, body, _ = self.run_review(spec, extra_patches=[boom])
        self.assertFalse(self.status(body)["generated"])
        self.assertIsNotNone(self.status(body)["blocked_reason"])

    def test_the_two_failure_kinds_are_described_differently(self):
        """An unreadable file and a file that broke while being COMPARED are
        different problems; one message covering both would misdescribe one."""
        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")}
        _, unreadable, _ = self.run_review(spec, read_failures={"CBT 2025"})
        boom = mock.patch.object(main.mismatch_mod, "build_mismatch_rows",
                                 mock.Mock(side_effect=RuntimeError("boom")))
        _, uncomparable, _ = self.run_review(spec, extra_patches=[boom])
        self.assertIn("could not be read or consolidated",
                      self.status(unreadable)["blocked_reason"])
        self.assertIn("could not be compared",
                      self.status(uncomparable)["blocked_reason"])
        self.assertNotIn("could not be read",
                         self.status(uncomparable)["blocked_reason"])

    def test_an_unassignable_servicer_carrying_rows_blocks(self):
        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025"),
                "Legacy 2023": _one("SOUP2023_1", "4000000004", name="Legacy 2023")}
        _, body, _ = self.run_review(spec)
        self.assertFalse(self.status(body)["generated"])
        self.assertIn("Legacy 2023", self.status(body)["blocked_reason"])
        self.assertIn("could not be assigned", self.status(body)["blocked_reason"])

    def test_an_unassignable_servicer_carrying_no_rows_does_not_block(self):
        """No coverage is lost, so a parsing quirk must not withhold a correct
        report. group_by_deal only records rows, so this holds structurally."""
        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025"),
                "Legacy 2023": {"deal": "SOUP2023_1", "postings": [_posting("4000000004")],
                                "overrides": {}, "candidates": []}}
        _, body, _ = self.run_review(spec)
        self.assertTrue(self.status(body)["generated"])
        self.assertEqual([r[0] for r in self.data_rows(body, SHEET_2025)], ["CBT 2025"])

    def test_a_provenance_gap_warns_but_never_blocks(self):
        """An override with no matching review record. The row is real and
        stays; only its provenance sentence degrades."""
        gp = "1000000001"
        spec = {"CBT 2025": _one(D2025, gp, name="CBT 2025")}
        # Override a field the reviewer has no log row for.
        spec["CBT 2025"]["overrides"] = {(gp, "Loan Status"): "DEFERRED"}
        _, body, _ = self.run_review(spec)
        self.assertTrue(self.status(body)["generated"])
        self.assertGreaterEqual(self.status(body)["provenance_gap_count"], 1)
        rows = self.data_rows(body, SHEET_2025)
        self.assertTrue(any(r[5] == "Loan Status" for r in rows))
        self.assertTrue(any(mr.NO_REVIEW_RECORD in (r[6] or "") for r in rows))

    def test_a_blocked_reason_is_never_absent_when_generation_failed(self):
        for kwargs in ({"read_failures": {"Lendistry 2025"}},):
            spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025"),
                    "Lendistry 2025": _one(D2025, "2000000002", name="Lendistry 2025")}
            _, body, _ = self.run_review(spec, **kwargs)
            st = self.status(body)
            self.assertFalse(st["generated"])
            self.assertTrue(st["blocked_reason"])

    def test_generated_and_bytes_never_disagree(self):
        for spec, kw in (
            ({"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")}, {}),
            ({"CBT 2025": _one(D2025, "1000000001", name="CBT 2025"),
              "Lendistry 2025": _one(D2025, "2000000002", name="Lendistry 2025")},
             {"read_failures": {"Lendistry 2025"}}),
        ):
            _, body, _ = self.run_review(spec, **kw)
            self.assertEqual(self.status(body)["generated"],
                             body["mismatch_report_b64"] is not None)


class NonInterferenceTests(_Harness):
    """Everything that existed before Commit 3 is byte-for-byte unchanged."""

    def _run_pair(self, spec):
        """The same review, once normally and once with the mismatch build
        forced to fail. Any difference between the two is interference."""
        _, normal, cap_n = self.run_review(spec)
        boom = mock.patch.object(main.mismatch_mod, "build_mismatch_rows",
                                 mock.Mock(side_effect=RuntimeError("boom")))
        _, broken, cap_b = self.run_review(spec, extra_patches=[boom])
        return normal, broken, cap_n, cap_b

    def setUp(self):
        self.spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025"),
                     "Millennium 2024": _one(D2024, "3000000003", name="Millennium 2024")}
        self.normal, self.broken, _, _ = self._run_pair(self.spec)

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

    def test_deals_status_and_summary_are_identical(self):
        for key in ("deals", "status", "summary", "corrections_not_applied",
                    "unknown_ids", "monthly_servicer_report_validation_issues"):
            self.assertEqual(self.normal[key], self.broken[key], key)

    def test_a_mismatch_failure_still_returns_http_200_with_the_workbooks(self):
        boom = mock.patch.object(main.mismatch_mod, "build_mismatch_rows",
                                 mock.Mock(side_effect=RuntimeError("boom")))
        code, body, _ = self.run_review(self.spec, extra_patches=[boom])
        self.assertEqual(code, 200)
        self.assertTrue(body["corrected_workbooks"])
        self.assertTrue(body["monthly_servicer_reports"])

    def test_neither_input_frame_is_mutated(self):
        seen = {}
        real_consolidate = checks.consolidate_cycle_postings

        def spy(df):
            out = real_consolidate(df)
            seen.setdefault("frames", []).append((out, out.copy(deep=True)))
            return out

        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")}
        with mock.patch.object(main.checks_mod, "consolidate_cycle_postings", spy):
            self.run_review(spec)
        self.assertTrue(seen["frames"])
        for after, before in seen["frames"]:
            pd.testing.assert_frame_equal(after, before)

    def test_the_verified_override_map_is_not_mutated(self):
        gp = "1000000001"
        overrides = {(gp, "Delinquency Status"): "30 DAYS"}
        spec = {"CBT 2025": _one(D2025, gp, overrides=overrides, name="CBT 2025")}
        self.run_review(spec)
        self.assertEqual(overrides, {(gp, "Delinquency Status"): "30 DAYS"})


class ResponseSurfaceTests(_Harness):
    """Minimal by design; additive by construction."""

    EXISTING_KEYS = {
        "corrected_workbooks", "monthly_servicer_reports",
        "monthly_servicer_report_validation_issues", "status", "deals",
        "review_completeness", "reporting_readiness",
        "portfolio_summary_final_b64", "portfolio_summary_final_filename",
        "summary", "corrections_not_applied", "unknown_ids",
    }
    NEW_KEYS = {"mismatch_report_b64", "mismatch_report_filename",
                "mismatch_report_status"}

    def setUp(self):
        _, self.body, _ = self.run_review(
            {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025")})

    def test_every_pre_existing_key_survives(self):
        self.assertTrue(self.EXISTING_KEYS <= set(self.body))

    def test_exactly_three_keys_were_added(self):
        self.assertEqual(set(self.body) - self.EXISTING_KEYS, self.NEW_KEYS)

    def test_the_status_object_carries_exactly_four_fields(self):
        self.assertEqual(set(self.status(self.body)),
                         {"generated", "blocked_reason", "mismatch_count",
                          "provenance_gap_count"})

    def test_no_loan_level_diagnostic_array_is_exposed(self):
        """provenance_gaps and unassigned stay server-side: they carry SBA GP
        Numbers and whole rows the client has no use for."""
        for key in ("provenance_gaps", "unassigned", "mismatch_rows",
                    "mismatch_failures"):
            self.assertNotIn(key, self.body)
            self.assertNotIn(key, self.status(self.body))

    def test_no_sba_gp_number_appears_in_a_blocked_reason(self):
        gp = "4000000004"
        spec = {"CBT 2025": _one(D2025, "1000000001", name="CBT 2025"),
                "Legacy 2023": _one("SOUP2023_1", gp, name="Legacy 2023")}
        _, body, _ = self.run_review(spec)
        self.assertFalse(self.status(body)["generated"])
        self.assertNotIn(gp, self.status(body)["blocked_reason"])

    def test_the_whole_response_is_json_serialisable(self):
        json.dumps(self.body)


class BoundaryTests(unittest.TestCase):
    """Commit 3 is backend-only and touches nothing it was told not to."""

    def test_main_imports_the_report_module(self):
        import inspect
        self.assertIn("import mismatch_report as mismatch_mod", inspect.getsource(main))

    def test_the_report_module_still_imports_no_main(self):
        import ast
        import inspect
        tree = ast.parse(inspect.getsource(mr))
        names = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                names.update(a.name for a in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                names.add(node.module)
        self.assertNotIn("main", names)

    def test_the_five_field_scope_is_still_pinned_to_the_override_columns(self):
        self.assertEqual(set(mr.MISMATCH_FIELDS), set(checks._VERIFIED_OVERRIDE_COLUMNS))


if __name__ == "__main__":
    unittest.main()
