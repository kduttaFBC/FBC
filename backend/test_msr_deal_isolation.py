"""D3 — silent servicer dropout, gated PER SOUP DEAL.

WHAT THIS COVERS
----------------
CLIENT-CONFIRMED (2026-08-13). One loop in /api/corrections/review built both
post-review reports and contained four distinct failure behaviours, only one of
them deliberate:

  A  read fails             -> silent skip: no summary row, no MSR rows, no
                               remittance row, yet its corrected workbook was
                               already generated and still returned
  B  consolidate/build_rows -> those calls sat OUTSIDE the try, so the exception
                               escaped to the outer handler and returned HTTP
                               500, discarding corrected workbooks whose
                               decisions were already applied and whose
                               Correction Logs were already written
  C  deal produces 0 rows   -> silent `continue`, indistinguishable from a bug
                               that emptied the deal
  D  every servicer fails   -> empty list + None, rendering nothing at all:
                               total failure looked identical to no findings

THE CONTROL. One try around the whole per-servicer body; structured failure
records naming the stage; expected-vs-reported reconciliation against
correction_state (the authoritative roster since 2026-08-10); explicit per-deal
complete/blocked status; corrected workbooks always returned; never a 500.

GATED PER DEAL, which is the whole point of these tests. A failure in one SOUP
deal must never suppress a valid report for another: blocked SOUP 2024 + complete
SOUP 2025 still yields the SOUP 2025 workbook, and vice versa. A blocked deal
produces NO workbook -- never a partial or degraded one, because the MSR is a
population-level consolidation and a partial consolidation is not a degraded
report, it is a wrong one.

The partition key is the deal resolved at /api/audit and persisted on the
correction_state entry. It HAS to be persisted: the review endpoint re-reads the
bytes and the deal is a RESULT of that read, so a servicer whose bytes fail there
would otherwise have no knowable deal, and one unreadable file would block every
deal -- the opposite of the isolation required.

NOT COVERED HERE. D4 (Human Review completeness) is out of scope: this suite
decides only whether a report is produced, never whether a review was finished.
Business rules are untouched and not re-tested -- Servicing Population,
Outstanding, PIF, Liquidation, Not Making Payment, the delinquency ladder,
reversal handling, the correction write path and closed-at-zero all keep their
own suites.
"""

import asyncio
import json
import unittest
from unittest import mock

import pandas as pd

import checks
import main
import monthly_servicer_report as msr


D2024 = "SOUP2024_1"
D2025 = "SOUP2025_1"
HEADER_ROW = 9

# servicer -> deal. Two servicers per deal so "one fails" and "all fail" are
# different scenarios within the same deal.
ROSTER = {
    "CBT 2025": D2025,
    "Lendistry 2025": D2025,
    "Millennium 2024": D2024,
    "Poppy 2024": D2024,
}


def _si(gp):
    return pd.DataFrame([{
        "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}",
        "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
        "Payment Received": 500.0, "Principal": 100.0, "Net Interest": 400.0,
        "Servicing Fee": 10.0, "Scheduled Payment": 500.0,
        "Beginning Balance": 100000.0, "Ending Balance": 99900.0,
        "_excel_row": HEADER_ROW + 1, "_sheet_name": "Servicer Inputs",
        "_header_excel_row": HEADER_ROW,
    }])


def _candidate(servicer, gp):
    return {
        "id": f"{servicer}::{gp}::Delinquency Status", "gp": gp,
        "field": "Delinquency Status", "corrected_value": "30 DAYS",
        "reported_value": "CURRENT", "excel_rows": [HEADER_ROW + 1],
        "skipped_row_refs": [], "sheet_name": "Servicer Inputs",
        "header_excel_row": HEADER_ROW, "eligibility_rule": "test",
        "review_required": False, "review_reason": None,
    }


def _fake_workbook(file_bytes, original_filename, institution, candidates,
                   decisions, reviewer, source_sha256, cycle_month,
                   reviewed_on=None):
    """Stands in for the real generator. Corrected workbooks and their
    Correction Logs are produced BEFORE any D3 gate and must survive it."""
    return {
        "file_bytes": b"CORRECTED-WORKBOOK-WITH-LOG",
        "filename": f"{institution}_CORRECTED.xlsx",
        "applied_count": len(decisions), "rejected_count": 0, "returned_count": 0,
        "corrections_not_applied": [], "unknown_ids": [], "skip_reason": None,
        "verified_overrides": {},
    }


class _Harness(unittest.TestCase):
    """Drives the real endpoint with the read / consolidate / build_rows calls
    patched so a failure can be injected at a chosen servicer and stage."""

    def run_review(self, fail=None, empty_rows=(), roster=None,
                   deal_override=None):
        fail = dict(fail or {})
        empty_rows = set(empty_rows)
        roster = dict(roster or ROSTER)
        gp_of = {s: f"10000000{i:02d}" for i, s in enumerate(roster, start=1)}

        state = {}
        for servicer, deal in roster.items():
            state[servicer] = {
                "original_bytes": f"BYTES::{servicer}".encode(),
                "original_filename": f"{servicer}.xlsx",
                "source_sha256": f"sha-{servicer}",
                "institution": servicer,
                "cycle_month": pd.Timestamp("2026-06-01"),
                "deal": deal,
                "candidates": [_candidate(servicer, gp_of[servicer])],
                "excluded_gps": [],
            }

        def fake_read(byts, filename=None):
            servicer = byts.decode().split("::", 1)[1]
            if fail.get(servicer) == "read":
                raise ValueError(f"could not read workbook for {servicer}")
            deal = (deal_override or {}).get(servicer, roster[servicer])
            return (servicer, servicer, pd.DataFrame(), _si(gp_of[servicer]), deal)

        real_views = checks.verified_portfolio_summary_views

        def fake_views(si, overrides, servicer_name=None, excluded_gps=None):
            if fail.get(servicer_name) == "consolidate":
                raise RuntimeError(f"consolidation failed for {servicer_name}")
            return real_views(si, overrides or {}, servicer_name,
                              excluded_gps=excluded_gps)

        def fake_rows(si, soup_key, **kw):
            if fail.get(soup_key) == "build_rows":
                raise RuntimeError(f"row build failed for {soup_key}")
            if soup_key in empty_rows:
                return []
            return [{"SBA GP Number": gp_of[soup_key], "Servicer Name": soup_key,
                     "Ending Balance": 99900.0}]

        captured = {}

        def fake_summary(blocks, buffer, cycle_month=None,
                         title=None, blocked_deals=None):
            captured["blocks"] = blocks
            captured["blocked_deals"] = list(blocked_deals or [])
            buffer.write(b"SUMMARY")

        sid = "d3-test-session"
        main.SESSION_CACHE[sid] = {"soup_df": None, "created": 9e18,
                                   "corrections": state}
        payload = {
            "session_id": sid, "reviewed_by": "Investigator",
            "decisions": [{"id": c["id"], "decision": "reject"}
                          for e in state.values() for c in e["candidates"]],
        }

        with mock.patch.object(main.file_readers, "read_servicer_file", fake_read), \
             mock.patch.object(main.checks_mod, "verified_portfolio_summary_views", fake_views), \
             mock.patch.object(main.msr_mod, "build_servicer_rows", fake_rows), \
             mock.patch.object(main.file_readers, "read_note_series_label",
                               lambda b: "2025 FBC-1"), \
             mock.patch.object(main.file_readers, "read_remittance_date",
                               lambda b: pd.Timestamp("2026-07-23")), \
             mock.patch.object(main.msr_mod, "generate_monthly_servicer_report_bytes",
                               lambda report, buf: buf.write(b"MSR-XLSX")), \
             mock.patch.object(main.report_mod, "generate_final_portfolio_summary_bytes",
                               fake_summary), \
             mock.patch.object(main.corrections_mod, "generate_reviewed_workbook",
                               _fake_workbook):
            resp = asyncio.new_event_loop().run_until_complete(
                main.corrections_review(payload))
        body = json.loads(bytes(resp.body).decode())
        return resp.status_code, body, captured

    # --- helpers -----------------------------------------------------------
    @staticmethod
    def deals(body):
        return {d["deal"]: d for d in body.get("deals") or []}

    @staticmethod
    def msr_deals(body):
        return {r["deal"] for r in body.get("monthly_servicer_reports") or []}


class DealIsolationTests(_Harness):
    """The six scenarios that prove one deal cannot suppress another."""

    def test_a_2024_failure_leaves_2025_complete_and_its_msr_generated(self):
        status, body, _cap = self.run_review(fail={"Millennium 2024": "read"})
        self.assertEqual(status, 200)
        d = self.deals(body)
        self.assertEqual(d[D2024]["status"], "blocked")
        self.assertEqual(d[D2025]["status"], "complete")
        self.assertFalse(d[D2024]["report_available"])
        self.assertTrue(d[D2025]["report_available"])
        self.assertEqual(self.msr_deals(body), {D2025})
        self.assertEqual(body["status"], "blocked")   # any blocked deal

    def test_a_2025_failure_leaves_2024_complete_and_its_msr_generated(self):
        status, body, _cap = self.run_review(fail={"CBT 2025": "read"})
        self.assertEqual(status, 200)
        d = self.deals(body)
        self.assertEqual(d[D2025]["status"], "blocked")
        self.assertEqual(d[D2024]["status"], "complete")
        self.assertEqual(self.msr_deals(body), {D2024})

    def test_a_zero_row_2025_deal_blocks_2025_only(self):
        status, body, _cap = self.run_review(
            empty_rows={"CBT 2025", "Lendistry 2025"})
        self.assertEqual(status, 200)
        d = self.deals(body)
        self.assertEqual(d[D2025]["status"], "blocked")
        self.assertEqual(d[D2024]["status"], "complete")
        self.assertEqual(self.msr_deals(body), {D2024})
        codes = {i.get("code") for i in body["monthly_servicer_report_validation_issues"]}
        self.assertIn(msr.ISSUE_DEAL_EMPTY, codes)
        # Every expected servicer DID report -- the deal is blocked on rows, not
        # on roster, so no failure record is invented.
        self.assertEqual(d[D2025]["failures"], [])
        self.assertEqual(len(d[D2025]["servicers_reported"]), 2)

    def test_all_servicers_in_one_deal_failing_blocks_that_deal_only(self):
        status, body, _cap = self.run_review(
            fail={"Millennium 2024": "read", "Poppy 2024": "consolidate"})
        self.assertEqual(status, 200)
        d = self.deals(body)
        self.assertEqual(d[D2024]["status"], "blocked")
        self.assertEqual(d[D2024]["servicers_reported"], [])
        self.assertEqual(len(d[D2024]["failures"]), 2)
        self.assertEqual({f["stage"] for f in d[D2024]["failures"]},
                         {"read", "consolidate"})
        self.assertEqual(d[D2025]["status"], "complete")
        self.assertEqual(self.msr_deals(body), {D2025})

    def test_both_deals_failing_produces_no_msr_but_keeps_workbooks(self):
        status, body, _cap = self.run_review(
            fail={"Millennium 2024": "read", "Poppy 2024": "read",
                  "CBT 2025": "build_rows", "Lendistry 2025": "consolidate"})
        self.assertEqual(status, 200)             # never a 500
        d = self.deals(body)
        self.assertEqual(d[D2024]["status"], "blocked")
        self.assertEqual(d[D2025]["status"], "blocked")
        self.assertEqual(body["monthly_servicer_reports"], [])
        self.assertEqual(body["status"], "blocked")
        # Behaviour D: total failure must not look like "nothing to report".
        self.assertEqual(len(body["deals"]), 2)
        for entry in body["deals"]:
            self.assertTrue(entry["failures"])
        # Corrected workbooks survive the gate, with their Correction Logs.
        self.assertEqual(len(body["corrected_workbooks"]), 4)

    def test_happy_path_produces_both_complete_deal_reports(self):
        status, body, _cap = self.run_review()
        self.assertEqual(status, 200)
        d = self.deals(body)
        self.assertEqual(d[D2024]["status"], "complete")
        self.assertEqual(d[D2025]["status"], "complete")
        self.assertEqual(self.msr_deals(body), {D2024, D2025})
        self.assertEqual(body["status"], "complete")
        for entry in body["deals"]:
            self.assertEqual(entry["failures"], [])
            self.assertTrue(entry["report_available"])
            self.assertEqual(sorted(entry["servicers_expected"]),
                             sorted(entry["servicers_reported"]))


class FailureRecordTests(_Harness):
    """Behaviours A and B collapse into one named outcome."""

    def test_a_consolidate_failure_is_blocked_not_a_500(self):
        status, body, _cap = self.run_review(fail={"CBT 2025": "consolidate"})
        self.assertEqual(status, 200)
        f = self.deals(body)[D2025]["failures"][0]
        self.assertEqual(f["stage"], "consolidate")
        self.assertEqual(f["error_type"], "RuntimeError")

    def test_a_build_rows_failure_is_blocked_not_a_500(self):
        status, body, _cap = self.run_review(fail={"CBT 2025": "build_rows"})
        self.assertEqual(status, 200)
        self.assertEqual(self.deals(body)[D2025]["failures"][0]["stage"], "build_rows")

    def test_a_read_failure_is_attributed_to_its_persisted_deal(self):
        # The reason the deal is persisted at /api/audit: the read that would
        # have told us the deal is the read that failed.
        _s, body, _cap = self.run_review(fail={"Lendistry 2025": "read"})
        f = self.deals(body)[D2025]["failures"][0]
        self.assertEqual(f["servicer"], "Lendistry 2025")
        self.assertEqual(f["deal"], D2025)
        self.assertEqual(f["stage"], "read")

    def test_a_failure_record_names_everything_a_reviewer_needs(self):
        _s, body, _cap = self.run_review(fail={"CBT 2025": "read"})
        f = self.deals(body)[D2025]["failures"][0]
        for key in ("servicer", "filename", "deal", "stage", "error_type",
                    "error_detail", "source_sha256"):
            self.assertIn(key, f)
        self.assertEqual(f["filename"], "CBT 2025.xlsx")
        self.assertEqual(f["source_sha256"], "sha-CBT 2025")

    def test_a_deal_mismatch_fails_closed(self):
        # Persisted and freshly-read deal cannot legitimately differ (same bytes,
        # same parser). If they ever do, reconciliation would compare the wrong
        # sets, so it blocks and names it rather than picking one.
        _s, body, _cap = self.run_review(
            deal_override={"CBT 2025": D2024})
        d = self.deals(body)
        self.assertEqual(d[D2025]["status"], "blocked")
        self.assertEqual(d[D2025]["failures"][0]["stage"], "deal_mismatch")

    def test_reconciliation_is_a_set_comparison_not_a_count(self):
        _s, body, _cap = self.run_review(fail={"CBT 2025": "read"})
        d = self.deals(body)[D2025]
        self.assertEqual(sorted(d["servicers_expected"]),
                         ["CBT 2025", "Lendistry 2025"])
        self.assertEqual(d["servicers_reported"], ["Lendistry 2025"])


class FinalAuditedSummaryTests(_Harness):
    """Blocked deals are named, never zero-filled, never silently omitted."""

    def test_a_complete_deals_block_is_preserved(self):
        _s, _body, cap = self.run_review(fail={"Millennium 2024": "read"})
        labels = [b["totals"]["counts"].get("Servicer") for b in cap["blocks"]]
        self.assertTrue(any("2025" in str(l) for l in labels))

    def test_a_blocked_deal_is_not_in_the_figure_blocks(self):
        # Its partial servicer figures must not be presented -- that would be a
        # degraded summary, the same defect as a partial MSR.
        _s, _body, cap = self.run_review(fail={"Millennium 2024": "read"})
        for block in cap["blocks"]:
            for record in block["counts"]:
                self.assertNotEqual(record.get("Servicer"), "Poppy 2024")

    def test_a_blocked_deal_is_named_with_a_reason(self):
        _s, _body, cap = self.run_review(fail={"Millennium 2024": "read"})
        blocked = cap["blocked_deals"]
        self.assertEqual([b["deal"] for b in blocked], [D2024])
        self.assertIn("Millennium 2024", blocked[0]["reason"])
        self.assertIn("read", blocked[0]["reason"])

    def test_a_zero_row_deal_is_named_in_the_summary_and_keeps_no_figures(self):
        """INVERTED 2026-08-13. This previously asserted blocked_deals == [] for
        the zero-row case, which pinned an ORDERING FLAW: the summary was built
        before the MSR pass, so a deal found zero-row afterwards was reported
        blocked in the response while its figures were already written into the
        workbook as complete. Two artifacts from one run disagreeing.

        CLIENT-CONFIRMED: once a deal is blocked its figures are not FINAL
        AUDITED figures, whenever they happened to be computed. It gets a named
        marker and no numeric block."""
        _s, body, cap = self.run_review(empty_rows={"CBT 2025", "Lendistry 2025"})
        self.assertEqual([b["deal"] for b in cap["blocked_deals"]], [D2025])
        self.assertIn("zero rows", cap["blocked_deals"][0]["reason"])
        # ...and no numeric block survives for it.
        for block in cap["blocks"]:
            for record in block["counts"]:
                with self.subTest(servicer=record.get("Servicer")):
                    self.assertNotIn(record.get("Servicer"),
                                     ("CBT 2025", "Lendistry 2025"))
        # The complete deal is untouched.
        self.assertEqual(self.msr_deals(body), {D2024})

    def test_no_blocked_markers_on_the_happy_path(self):
        _s, _body, cap = self.run_review()
        self.assertEqual(cap["blocked_deals"], [])


class VerdictInvariantTests(_Harness):
    """The machine-checkable form of "the two artifacts cannot disagree".

    status == "complete"  <=>  report_available is True
                          <=>  the deal contributes numeric summary blocks

    It holds because the MSR pass is the last place a verdict may change: the
    only two ways a provisionally-complete deal can fail to produce a workbook
    (absent from msr_servicers_by_deal, or zero rows) both set status to blocked
    in the same pass that would have set report_available."""

    SCENARIOS = [
        ("happy path", {}, ()),
        ("2024 read failure", {"Millennium 2024": "read"}, ()),
        ("2025 read failure", {"CBT 2025": "read"}, ()),
        ("2025 consolidate failure", {"CBT 2025": "consolidate"}, ()),
        ("2025 build_rows failure", {"CBT 2025": "build_rows"}, ()),
        ("2025 zero rows", {}, ("CBT 2025", "Lendistry 2025")),
        ("all of 2024 fails", {"Millennium 2024": "read", "Poppy 2024": "consolidate"}, ()),
        ("both deals fail", {"Millennium 2024": "read", "Poppy 2024": "read",
                             "CBT 2025": "build_rows", "Lendistry 2025": "consolidate"}, ()),
    ]

    def test_complete_iff_report_available(self):
        for name, fail, empty in self.SCENARIOS:
            _s, body, _cap = self.run_review(fail=fail, empty_rows=empty)
            for d in body["deals"]:
                with self.subTest(scenario=name, deal=d["deal"]):
                    self.assertEqual(d["status"] == "complete", d["report_available"])

    def test_complete_iff_an_msr_workbook_exists(self):
        for name, fail, empty in self.SCENARIOS:
            _s, body, _cap = self.run_review(fail=fail, empty_rows=empty)
            produced = self.msr_deals(body)
            for d in body["deals"]:
                with self.subTest(scenario=name, deal=d["deal"]):
                    self.assertEqual(d["status"] == "complete", d["deal"] in produced)

    def test_only_complete_deals_contribute_numeric_summary_blocks(self):
        for name, fail, empty in self.SCENARIOS:
            _s, body, cap = self.run_review(fail=fail, empty_rows=empty)
            complete = {d["deal"] for d in body["deals"] if d["status"] == "complete"}
            blocked = {d["deal"] for d in body["deals"] if d["status"] == "blocked"}
            # Every blocked deal is named, none is omitted.
            with self.subTest(scenario=name):
                self.assertEqual({b["deal"] for b in cap["blocked_deals"]}, blocked)
            # No servicer belonging to a blocked deal appears in a figure block.
            blocked_servicers = {s for d in body["deals"] if d["status"] == "blocked"
                                 for s in d["servicers_reported"]}
            for block in cap["blocks"]:
                for record in block["counts"]:
                    with self.subTest(scenario=name, servicer=record.get("Servicer")):
                        self.assertNotIn(record.get("Servicer"), blocked_servicers)

    def test_corrected_workbooks_survive_every_scenario(self):
        for name, fail, empty in self.SCENARIOS:
            _s, body, _cap = self.run_review(fail=fail, empty_rows=empty)
            with self.subTest(scenario=name):
                self.assertEqual(len(body["corrected_workbooks"]), len(ROSTER))

    def test_never_a_500_in_any_scenario(self):
        for name, fail, empty in self.SCENARIOS:
            status, _body, _cap = self.run_review(fail=fail, empty_rows=empty)
            with self.subTest(scenario=name):
                self.assertEqual(status, 200)


class ReportRenderingTests(unittest.TestCase):
    """report.py renders the marker without zero-filling."""

    def _render(self, blocked_deals):
        import io
        from openpyxl import load_workbook
        import report as report_mod
        si = _si("1000000001")
        v = checks.verified_portfolio_summary_views(si, {}, "S")
        blocks = checks.group_portfolio_summary_by_deal(
            [v["dollars"]], [v["counts"]], [D2025])
        buf = io.BytesIO()
        report_mod.generate_final_portfolio_summary_bytes(
            blocks, buf, blocked_deals=blocked_deals)
        buf.seek(0)
        ws = load_workbook(buf)[report_mod.FINAL_SUMMARY_DEFAULT_TITLE]
        return ws, report_mod

    def test_the_marker_names_the_deal_and_the_reason(self):
        ws, report_mod = self._render([{"deal": D2024, "reason": "two servicers failed"}])
        col_a = [str(ws.cell(row=r, column=1).value or "")
                 for r in range(1, ws.max_row + 1)]
        self.assertIn(report_mod.BLOCKED_SECTION_TITLE, col_a)
        self.assertTrue(any(D2024 in v and "BLOCKED" in v and "two servicers failed" in v
                            for v in col_a))

    def test_the_marker_row_is_empty_never_zero_filled(self):
        ws, report_mod = self._render([{"deal": D2024, "reason": "failed"}])
        target = None
        for r in range(1, ws.max_row + 1):
            if "BLOCKED" in str(ws.cell(row=r, column=1).value or ""):
                target = r
                break
        self.assertIsNotNone(target)
        for c in range(2, len(report_mod.FINAL_SUMMARY_CATEGORIES) + 2):
            with self.subTest(col=c):
                self.assertIsNone(ws.cell(row=target, column=c).value)

    def test_the_caveat_precedes_the_figures(self):
        ws, report_mod = self._render([{"deal": D2024, "reason": "failed"}])
        col_a = [str(ws.cell(row=r, column=1).value or "")
                 for r in range(1, ws.max_row + 1)]
        self.assertLess(col_a.index(report_mod.BLOCKED_SECTION_TITLE),
                        col_a.index("COUNT"))

    def test_no_band_when_nothing_is_blocked(self):
        ws, report_mod = self._render(None)
        col_a = [str(ws.cell(row=r, column=1).value or "")
                 for r in range(1, ws.max_row + 1)]
        self.assertNotIn(report_mod.BLOCKED_SECTION_TITLE, col_a)


class IssueCodeTests(unittest.TestCase):
    def test_the_new_codes_exist_beside_the_existing_vocabulary(self):
        self.assertEqual(msr.ISSUE_SERVICER_FAILED, "servicer_failed")
        self.assertEqual(msr.ISSUE_DEAL_EMPTY, "deal_empty")

    def test_the_msr_module_does_not_raise_them_itself(self):
        # Raised from main.py's reconciliation only; build_servicer_rows and
        # build_deal_report are unchanged and must not emit them.
        import inspect
        for fn in (msr.build_servicer_rows, msr.build_deal_report):
            src = inspect.getsource(fn)
            with self.subTest(fn=fn.__name__):
                self.assertNotIn("ISSUE_SERVICER_FAILED", src)
                self.assertNotIn("ISSUE_DEAL_EMPTY", src)


if __name__ == "__main__":
    unittest.main()
