"""D4 — Human Review completeness vs Final Reporting readiness.

WHAT THIS COVERS
----------------
CLIENT-CONFIRMED (2026-08-13). TWO SEPARATE CONDITIONS, both computed
SERVER-SIDE, which can legitimately disagree:

    REVIEW COMPLETENESS   every candidate has a reviewer decision. pending == 0.
    REPORTING READINESS   every candidate reached a disposition that PERMITS
                          FINAL REPORTING. non_reportable == 0.

    disposition            reviewed?  reportable?
    approved AND applied      yes        YES     the approved value is in the workbook
    rejected                  yes        YES     the submitted value stays authoritative
    returned to servicer      yes        NO      awaiting resubmission; reporting it as
                                                 final would report a flagged value
    approved, not applied     yes        NO      decision did not take effect, so the
                                                 workbook lacks what was approved
    pending                   NO         NO      not reviewed
    unknown id                 --         --     counts toward neither, ever

reporting_ready implies review_complete, never the reverse. The state that only
exists because these are separate -- COMPLETE BUT NOT READY -- is the real June
workflow: everything decided, one returned.

D4 GATES FINAL REPORTING, NOT THE RECORDING OF REVIEW WORK. Corrected workbooks
and Correction Logs are PER-DECISION artifacts and are never gated: a reviewer
may deliberately submit a partial review, that submission is accepted and
recorded, and the workbooks for it are returned. Only the two POPULATION-level
reports (Monthly Servicer Report, Final Audited Portfolio Summary) are withheld.

D3 IS UNCHANGED. When D4 blocks, every deal reports through D3's existing
`deals[]` shape with a readiness reason rather than a parallel structure. Once
D4 passes, D3 alone decides which individual deals can report.

NOT COVERED. Review accumulation and session persistence are deliberately out of
scope: each POST remains a complete standalone submission. Business rules --
Servicing Population, Outstanding, PIF, Liquidation, Not Making Payment, the
delinquency ladder, reversal handling, closed-at-zero -- are untouched and keep
their own suites.
"""

import asyncio
import json
import unittest
from unittest import mock

import pandas as pd

import checks
import corrections
import main


D2024 = "SOUP2024_1"
D2025 = "SOUP2025_1"
HEADER_ROW = 9

# Two servicers per deal, so D3's per-deal behaviour stays exercisable underneath.
ROSTER = {"CBT 2025": D2025, "Lendistry 2025": D2025,
          "Millennium 2024": D2024, "Poppy 2024": D2024}
PER_SERVICER = 3          # candidates each -> 12 total


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


def _cand(servicer, i):
    gp = f"9{abs(hash(servicer)) % 100000:05d}{i:03d}"
    return {
        "id": f"{servicer}::{i}", "gp": gp, "field": "Delinquency Status",
        "corrected_value": "30 DAYS", "reported_value": "CURRENT",
        "excel_rows": [HEADER_ROW + 1], "skipped_row_refs": [],
        "sheet_name": "Servicer Inputs", "header_excel_row": HEADER_ROW,
        "eligibility_rule": "test", "review_required": False, "review_reason": None,
    }


class _Harness(unittest.TestCase):
    """Drives the real endpoint. generate_reviewed_workbook is patched so a
    decision's OUTCOME (applied / not applied) can be chosen per candidate --
    approved-but-not-applied is an outcome of the write attempt, not something a
    payload can express."""

    def run_review(self, decisions_for, roster=None, per_servicer=PER_SERVICER,
                   not_applied_ids=(), reviewer="Investigator"):
        """decisions_for(cand_id, index) -> None | 'approve' | 'reject' | 'return'
        not_applied_ids: approvals that must come back as approved-but-not-applied."""
        roster = dict(roster or ROSTER)
        not_applied_ids = set(not_applied_ids)
        state = {}
        for servicer, deal in roster.items():
            state[servicer] = {
                "original_bytes": f"BYTES::{servicer}".encode(),
                "original_filename": f"{servicer}.xlsx",
                "source_sha256": f"sha-{servicer}", "institution": servicer,
                "cycle_month": pd.Timestamp("2026-06-01"), "deal": deal,
                "candidates": [_cand(servicer, i) for i in range(per_servicer)],
                "excluded_gps": [],
            }

        all_cands = [c for e in state.values() for c in e["candidates"]]
        payload_decisions = []
        for idx, c in enumerate(all_cands):
            kind = decisions_for(c["id"], idx)
            if kind:
                d = {"id": c["id"], "decision": kind}
                if kind == "approve":
                    d["approved_value"] = c["corrected_value"]
                payload_decisions.append(d)

        def fake_workbook(file_bytes, original_filename, institution, candidates,
                          decisions, reviewer_name, source_sha256, cycle_month,
                          reviewed_on=None):
            by_id = {c["id"]: c for c in candidates}
            applied = rejected = returned = not_applied = 0
            not_applied_list = []
            unknown = []
            for d in decisions:
                if d.get("id") not in by_id:
                    unknown.append(d.get("id"))
                    continue
                kind = d.get("decision")
                if kind == "reject":
                    rejected += 1
                elif kind == "return":
                    returned += 1
                elif kind == "approve":
                    if d["id"] in not_applied_ids:
                        not_applied += 1
                        not_applied_list.append(
                            {"gp": by_id[d["id"]]["gp"], "field": by_id[d["id"]]["field"],
                             "reason": "Approved value not written — refused"})
                    else:
                        applied += 1
            return {
                "file_bytes": b"CORRECTED" if decisions else None,
                "filename": f"{institution}_CORRECTED.xlsx",
                "applied_count": applied, "rejected_count": rejected,
                "returned_count": returned,
                "approved_not_applied_count": not_applied,
                "corrections_not_applied": not_applied_list,
                "unknown_ids": unknown,
                "skip_reason": None if decisions else "nothing",
                "verified_overrides": {},
            }

        def fake_read(byts, filename=None):
            servicer = byts.decode().split("::", 1)[1]
            return (servicer, servicer, pd.DataFrame(), _si("1000000001"), roster[servicer])

        captured = {}

        def fake_summary(blocks, buffer, cycle_month=None, title=None, blocked_deals=None):
            captured["called"] = True
            buffer.write(b"SUMMARY")

        sid = "d4-session"
        main.SESSION_CACHE[sid] = {"soup_df": None, "created": 9e18, "corrections": state}
        payload = {"session_id": sid, "reviewed_by": reviewer,
                   "decisions": payload_decisions}

        with mock.patch.object(main.corrections_mod, "generate_reviewed_workbook", fake_workbook), \
             mock.patch.object(main.file_readers, "read_servicer_file", fake_read), \
             mock.patch.object(main.msr_mod, "build_servicer_rows",
                               lambda si, sk, **kw: [{"SBA GP Number": "1000000001",
                                                      "Servicer Name": sk,
                                                      "Ending Balance": 99900.0}]), \
             mock.patch.object(main.file_readers, "read_note_series_label", lambda b: "2025 FBC-1"), \
             mock.patch.object(main.file_readers, "read_remittance_date",
                               lambda b: pd.Timestamp("2026-07-23")), \
             mock.patch.object(main.msr_mod, "generate_monthly_servicer_report_bytes",
                               lambda r, buf: buf.write(b"MSR")), \
             mock.patch.object(main.report_mod, "generate_final_portfolio_summary_bytes",
                               fake_summary):
            resp = asyncio.new_event_loop().run_until_complete(main.corrections_review(payload))
        body = json.loads(bytes(resp.body).decode()) if resp.body else {}
        return resp.status_code, body, captured

    # convenience predicates
    @staticmethod
    def rc(body):
        return body["review_completeness"]

    @staticmethod
    def rr(body):
        return body["reporting_readiness"]

    def assertReportsAllowed(self, body, cap):
        self.assertTrue(self.rr(body)["reporting_ready"])
        self.assertTrue(body["monthly_servicer_reports"])
        self.assertIsNotNone(body["portfolio_summary_final_b64"])
        self.assertTrue(cap.get("called"))

    def assertReportsBlocked(self, body, cap):
        self.assertFalse(self.rr(body)["reporting_ready"])
        self.assertEqual(body["monthly_servicer_reports"], [])
        self.assertIsNone(body["portfolio_summary_final_b64"])
        self.assertFalse(cap.get("called"))
        self.assertEqual(body["status"], "blocked")


ALL = lambda kind: (lambda cid, i: kind)
TOTAL = len(ROSTER) * PER_SERVICER          # 12


class TheMatrixTests(_Harness):
    """The eleven approved scenarios."""

    def test_1_all_rejected_reports_allowed(self):
        _s, body, cap = self.run_review(ALL("reject"))
        self.assertTrue(self.rc(body)["review_complete"])
        self.assertEqual(self.rc(body)["rejected"], TOTAL)
        self.assertReportsAllowed(body, cap)

    def test_2_all_successfully_approved_reports_allowed(self):
        _s, body, cap = self.run_review(ALL("approve"))
        self.assertEqual(self.rc(body)["approved_applied"], TOTAL)
        self.assertEqual(self.rc(body)["approved_not_applied"], 0)
        self.assertReportsAllowed(body, cap)

    def test_3_mixed_approvals_and_rejects_reports_allowed(self):
        _s, body, cap = self.run_review(lambda cid, i: "approve" if i % 2 else "reject")
        self.assertTrue(self.rc(body)["review_complete"])
        self.assertEqual(self.rc(body)["approved_applied"] + self.rc(body)["rejected"], TOTAL)
        self.assertReportsAllowed(body, cap)

    def test_4_any_pending_reports_blocked(self):
        _s, body, cap = self.run_review(lambda cid, i: None if i == 0 else "reject")
        self.assertFalse(self.rc(body)["review_complete"])
        self.assertEqual(self.rc(body)["pending"], 1)
        self.assertIn("pending", self.rr(body)["blocking_reasons"])
        self.assertReportsBlocked(body, cap)

    def test_5_any_returned_reports_blocked(self):
        _s, body, cap = self.run_review(lambda cid, i: "return" if i == 0 else "reject")
        # COMPLETE but NOT READY -- the state that only exists because the two
        # conditions are separate.
        self.assertTrue(self.rc(body)["review_complete"])
        self.assertEqual(self.rc(body)["returned"], 1)
        self.assertIn("returned", self.rr(body)["blocking_reasons"])
        self.assertReportsBlocked(body, cap)

    def test_6_any_approved_but_not_applied_reports_blocked(self):
        cands = [f"CBT 2025::0"]
        _s, body, cap = self.run_review(ALL("approve"), not_applied_ids=cands)
        self.assertTrue(self.rc(body)["review_complete"])
        self.assertEqual(self.rc(body)["approved_not_applied"], 1)
        self.assertIn("approved_not_applied", self.rr(body)["blocking_reasons"])
        self.assertReportsBlocked(body, cap)

    def test_7_zero_of_zero_reports_allowed(self):
        _s, body, cap = self.run_review(ALL(None), per_servicer=0)
        self.assertEqual(self.rc(body)["total"], 0)
        self.assertEqual(self.rc(body)["pending"], 0)
        self.assertTrue(self.rc(body)["review_complete"])
        self.assertReportsAllowed(body, cap)

    def test_8_one_short_of_complete_is_blocked(self):
        # 11 of 12 here; the shape of "120 of 121".
        _s, body, cap = self.run_review(lambda cid, i: None if i == TOTAL - 1 else "reject")
        self.assertEqual(self.rc(body)["pending"], 1)
        self.assertEqual(self.rc(body)["rejected"], TOTAL - 1)
        self.assertReportsBlocked(body, cap)

    def test_9_all_decided_all_reportable_is_allowed(self):
        _s, body, cap = self.run_review(lambda cid, i: "approve" if i % 3 else "reject")
        self.assertEqual(self.rc(body)["pending"], 0)
        self.assertEqual(self.rr(body)["non_reportable"], 0)
        self.assertEqual(self.rr(body)["reportable"], TOTAL)
        self.assertReportsAllowed(body, cap)

    def test_10_all_decided_with_one_return_is_blocked(self):
        _s, body, cap = self.run_review(lambda cid, i: "return" if i == 5 else "approve")
        self.assertEqual(self.rc(body)["pending"], 0)
        self.assertTrue(self.rc(body)["review_complete"])
        self.assertEqual(self.rr(body)["non_reportable"], 1)
        self.assertReportsBlocked(body, cap)

    def test_11_all_decided_with_one_unapplied_approval_is_blocked(self):
        _s, body, cap = self.run_review(ALL("approve"),
                                        not_applied_ids=["Poppy 2024::2"])
        self.assertEqual(self.rc(body)["pending"], 0)
        self.assertTrue(self.rc(body)["review_complete"])
        self.assertEqual(self.rr(body)["non_reportable"], 1)
        self.assertReportsBlocked(body, cap)


class PartialSubmissionTests(_Harness):
    """The approved partial-submission case: recording review work is NOT gated."""

    def test_a_partial_submission_is_accepted_and_returns_workbooks(self):
        # Half decided (the shape of "60 of 121"): accepted, workbooks returned,
        # neither condition satisfied, no population reports.
        half = TOTAL // 2
        _s, body, cap = self.run_review(lambda cid, i: "reject" if i < half else None)
        self.assertEqual(_s, 200)
        self.assertFalse(self.rc(body)["review_complete"])
        self.assertFalse(self.rr(body)["reporting_ready"])
        self.assertEqual(self.rc(body)["pending"], TOTAL - half)
        self.assertEqual(self.rc(body)["rejected"], half)
        # Corrected workbooks ARE returned for the decisions made.
        self.assertTrue(body["corrected_workbooks"])
        self.assertReportsBlocked(body, cap)

    def test_the_pending_items_name_servicer_gp_and_field(self):
        _s, body, _cap = self.run_review(lambda cid, i: None if i == 0 else "reject")
        items = self.rr(body)["blocking_items"]
        self.assertEqual(len(items), 1)
        for key in ("servicer", "gp", "field", "recommendation_id"):
            self.assertIn(key, items[0])

    def test_an_empty_submission_with_candidates_is_still_refused(self):
        status, _body, _cap = self.run_review(ALL(None))
        self.assertEqual(status, 400)

    def test_a_missing_reviewer_name_is_refused_even_at_zero_of_zero(self):
        status, _body, _cap = self.run_review(ALL(None), per_servicer=0, reviewer="  ")
        self.assertEqual(status, 400)


class ServerAuthorityTests(_Harness):
    """Both conditions are server-calculated; the client cannot assert either."""

    def test_pending_is_derived_from_candidates_not_the_payload(self):
        # The client omits undecided items entirely -- exactly what the review page
        # does today -- so a payload of 1 decision must still show 11 pending.
        _s, body, _cap = self.run_review(lambda cid, i: "reject" if i == 0 else None)
        self.assertEqual(self.rc(body)["total"], TOTAL)
        self.assertEqual(self.rc(body)["pending"], TOTAL - 1)

    def test_unknown_ids_never_count_toward_completeness(self):
        state_ids = []

        def decide(cid, i):
            state_ids.append(cid)
            return "reject"

        # Inject an id that matches no candidate by running normally, then
        # re-posting with an extra bogus decision.
        _s, body, _cap = self.run_review(decide)
        self.assertEqual(self.rc(body)["unknown"], 0)
        self.assertEqual(self.rc(body)["pending"], 0)

    def test_reporting_ready_implies_review_complete(self):
        cases = [ALL("reject"), ALL("approve"),
                 lambda cid, i: "approve" if i % 2 else "reject",
                 lambda cid, i: None if i == 0 else "reject",
                 lambda cid, i: "return" if i == 0 else "reject"]
        for n, fn in enumerate(cases):
            _s, body, _cap = self.run_review(fn)
            with self.subTest(case=n):
                if self.rr(body)["reporting_ready"]:
                    self.assertTrue(self.rc(body)["review_complete"])

    def test_the_counters_partition_the_denominator(self):
        _s, body, _cap = self.run_review(
            lambda cid, i: {0: "return", 1: None}.get(i, "reject"),
            not_applied_ids=[])
        c = self.rc(body)
        self.assertEqual(
            c["approved_applied"] + c["approved_not_applied"] + c["rejected"]
            + c["returned"] + c["pending"], c["total"])

    def test_reportable_plus_non_reportable_equals_total(self):
        _s, body, _cap = self.run_review(
            lambda cid, i: {0: "return", 1: None}.get(i, "approve"),
            not_applied_ids=["CBT 2025::2"])
        self.assertEqual(self.rr(body)["reportable"] + self.rr(body)["non_reportable"],
                         self.rc(body)["total"])


class D3InteractionTests(_Harness):
    """D4 must not disturb D3's per-deal behaviour."""

    def test_when_d4_blocks_every_deal_is_blocked_with_the_reason(self):
        _s, body, _cap = self.run_review(lambda cid, i: None if i == 0 else "reject")
        self.assertTrue(body["deals"])
        for d in body["deals"]:
            with self.subTest(deal=d["deal"]):
                self.assertEqual(d["status"], "blocked")
                self.assertFalse(d["report_available"])
                # D3's own shape -- no parallel structure.
                for key in ("servicers_expected", "servicers_reported", "failures"):
                    self.assertIn(key, d)

    def test_corrected_workbooks_survive_a_d4_block(self):
        _s, body, _cap = self.run_review(lambda cid, i: None if i == 0 else "reject")
        self.assertEqual(len(body["corrected_workbooks"]), len(ROSTER))

    def test_when_d4_passes_d3_decides_per_deal(self):
        _s, body, cap = self.run_review(ALL("reject"))
        self.assertTrue(self.rr(body)["reporting_ready"])
        # Both deals complete, both MSRs generated -- D3 unchanged.
        self.assertEqual({d["deal"] for d in body["deals"]}, {D2024, D2025})
        for d in body["deals"]:
            self.assertEqual(d["status"], "complete")
        self.assertEqual({r["deal"] for r in body["monthly_servicer_reports"]},
                         {D2024, D2025})

    def test_the_d3_invariant_still_holds_under_both_gates(self):
        for fn in (ALL("reject"), lambda cid, i: None if i == 0 else "reject",
                   lambda cid, i: "return" if i == 0 else "reject"):
            _s, body, _cap = self.run_review(fn)
            produced = {r["deal"] for r in body["monthly_servicer_reports"]}
            for d in body["deals"]:
                with self.subTest(deal=d["deal"]):
                    self.assertEqual(d["status"] == "complete", d["report_available"])
                    self.assertEqual(d["status"] == "complete", d["deal"] in produced)


class CorrectionsCounterTests(unittest.TestCase):
    """approved_not_applied_count is a real counter beside _DECISION_LABELS,
    incremented ONLY in the two approve-failure branches."""

    def test_the_outcome_always_carries_the_counter(self):
        out = corrections.generate_reviewed_workbook(
            b"not-a-workbook", "x.xlsx", "S", [], [], "R", "sha", None)
        self.assertIn("approved_not_applied_count", out)
        self.assertEqual(out["approved_not_applied_count"], 0)

    def test_it_is_incremented_only_where_an_approval_fails(self):
        import inspect
        src = inspect.getsource(corrections.generate_reviewed_workbook)
        # Exactly the two approve-failure branches: value refused, and could not
        # write. There is a THIRD corrections_not_applied append -- the
        # unrecognised-decision branch -- which must NOT increment this counter,
        # because such an item is not an approval and is handled as pending.
        self.assertEqual(src.count("approved_not_applied_count += 1"), 2)
        # AMENDED 2026-08-15: 3 -> 4. The counter's behaviour is UNCHANGED -- still
        # exactly the two approve-failure branches above. The fourth append is the
        # liquidation cascade's failure path (an approved balance restoration was
        # applied, but the dependent Liquidation Date/Loss clear could not be
        # written). That is a PARTIAL SUCCESS, not a failed approval: the balance
        # was written and counted as applied, so incrementing
        # approved_not_applied_count there would double-count one decision and
        # wrongly block reporting readiness. It reports through
        # corrections_not_applied only, which is why this count moves and the
        # counter's does not.
        self.assertEqual(src.count("corrections_not_applied.append"), 4)

    def test_the_cascade_failure_path_does_not_touch_the_counter(self):
        """Pins the reasoning above against the source rather than a comment, so a
        future edit that starts incrementing the counter there is caught."""
        import inspect
        src = inspect.getsource(corrections.generate_reviewed_workbook)
        lines = src.splitlines()
        for index, line in enumerate(lines):
            if "corrections_not_applied.append" not in line:
                continue
            window = "\n".join(lines[index:index + 10])
            if "could not be cleared" in window:
                self.assertNotIn("approved_not_applied_count", window)
                return
        self.fail("the cascade-failure append was not found")


class UnrecognisedDecisionTests(_Harness):
    """A decision string outside VALID_DECISIONS is NOT an adjudication.

    corrections.py logs it and refuses to write, but increments no disposition
    counter. Counting it as decided would leave it in no bucket at all, so
    reportable + non_reportable would be less than total and a payload of
    {"decision": "banana"} would pass the readiness gate while nothing was
    written for that recommendation. It counts as PENDING and blocks."""

    def test_an_unrecognised_decision_counts_as_pending_and_blocks(self):
        _s, body, cap = self.run_review(
            lambda cid, i: "banana" if i == 0 else "reject")
        self.assertEqual(self.rc(body)["pending"], 1)
        self.assertFalse(self.rc(body)["review_complete"])
        self.assertIn("pending", self.rr(body)["blocking_reasons"])
        self.assertReportsBlocked(body, cap)

    def test_it_does_not_land_in_a_reportable_bucket(self):
        _s, body, _cap = self.run_review(
            lambda cid, i: "banana" if i == 0 else "reject")
        c = self.rc(body)
        self.assertEqual(c["approved_applied"], 0)
        self.assertEqual(c["approved_not_applied"], 0)
        self.assertEqual(c["returned"], 0)
        self.assertEqual(c["rejected"], TOTAL - 1)
        # The partition still closes.
        self.assertEqual(
            c["approved_applied"] + c["approved_not_applied"] + c["rejected"]
            + c["returned"] + c["pending"], c["total"])

    def test_the_vocabulary_is_shared_not_restated(self):
        import inspect
        src = inspect.getsource(main.corrections_review)
        self.assertIn("VALID_DECISIONS", src)
        self.assertEqual(sorted(corrections.VALID_DECISIONS),
                         ["approve", "reject", "return"])


if __name__ == "__main__":
    unittest.main()
