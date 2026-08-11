"""
FalconBridge Audit Intelligence — FastAPI backend.

Endpoints:
  POST /api/audit               reconcile SOUP vs. each uploaded servicer SRF file
  POST /api/corrections/review  generate corrected workbooks from a reviewer's decisions
  POST /api/chat                chat with FalconBridge AI about the audit
"""

# Makes 3.10+ union type-hint syntax (e.g. `list | None`) work on Python
# 3.9 by deferring annotation evaluation to strings instead of evaluating
# them eagerly at import time — no behavior change, just compatibility.
from __future__ import annotations

import io
import os
import sys
import time
import traceback
import base64
import gc
import hashlib
import uuid

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dotenv import load_dotenv
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import List

import checks as checks_mod
import report as report_mod
import chatbot as chatbot_mod
import file_readers
import corrections as corrections_mod
import preflight as preflight_mod
import memo_audit as memo_audit_mod


load_dotenv(override=False)

app = FastAPI(title="FalconBridge Audit Intelligence")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # POC — lock down after deployment
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

REPORTS_DIR = "/tmp/reports"
os.makedirs(REPORTS_DIR, exist_ok=True)

# Per-audit-session state, held in memory between requests:
#   - "soup_df": the full SOUP dataset, so the chatbot can look up ANY
#     account on demand (flagged or not), not just what was pre-loaded into
#     the chat context.
#   - "corrections": per servicer, the ORIGINAL uploaded bytes + their
#     SHA-256 + the eligible correction candidates, so POST
#     /api/corrections/review can generate a corrected workbook from the
#     exact source artifact the reviewer saw, applying only approved values.
#
# IMPORTANT — this is PROCESS-LOCAL state and assumes a SINGLE worker (the
# Render start command runs uvicorn with no --workers flag; see render.yaml).
# The chatbot's soup_df lookup depends on the same assumption. Under a
# multi-worker deployment a review POST (or a chat request) could land on a
# different worker than the audit that created the session, and BOTH would
# fail together — moving this to a shared store (Redis/object storage) is the
# prerequisite for scaling past one worker.
#
# Sessions expire after SESSION_TTL_SECONDS.
SESSION_CACHE = {}
SESSION_TTL_SECONDS = 3600


def _cleanup_old_sessions():
    now = time.time()
    expired = [sid for sid, entry in SESSION_CACHE.items() if now - entry["created"] > SESSION_TTL_SECONDS]
    for sid in expired:
        del SESSION_CACHE[sid]


# Caps chosen against measured token cost at Lendistry scale (~1,444 rows,
# 6 checks): uncapped results + flagged history ran ~118k tokens on their
# own before conversation history or tool calls. These caps bring that down
# to ~35k tokens while keeping issues_found counts exact and full detail one
# tool call away via lookup_loan_history / search_accounts_by_status.
MAX_RECORDS_PER_RESULT = 5          # example records shipped per check, per servicer
MAX_FLAGGED_LOANS_FOR_HISTORY = 100  # distinct loans included in the pre-loaded history CSV


def _cap_results_for_chat(all_results, max_records=MAX_RECORDS_PER_RESULT):
    """Returns a copy of all_results with each result's records list capped
    for the chatbot's context. issues_found stays the TRUE total; only the
    records list sent to the chatbot is trimmed. The full list remains
    available via search_accounts_by_status, and the Excel report (built
    from the uncapped all_results) is unaffected."""
    capped = []
    for r in all_results:
        r2 = dict(r)
        records = r2.get("records", [])
        if len(records) > max_records:
            r2["records"] = records[:max_records]
            r2["records_truncated"] = True
            r2["records_shown"] = max_records
        capped.append(r2)
    return capped


def _flagged_gp_numbers(all_results, max_loans=MAX_FLAGGED_LOANS_FOR_HISTORY):
    """Collects SBA GP Numbers that appear in any check's flagged records,
    across all servicers, capped to max_loans distinct loans so the full
    monthly history CSV built from these stays a bounded size. Any loan
    beyond the cap is still reachable via lookup_loan_history on demand."""
    gp_numbers = set()
    for result in all_results:
        for record in result.get("records", []):
            gp = record.get("SBA GP Number")
            if gp is not None:
                gp_numbers.add(str(gp).strip())
            if len(gp_numbers) >= max_loans:
                return gp_numbers
    return gp_numbers


def _resolve_portfolio_summary_cycle_month(cycle_months):
    """CLIENT-CONFIRMED (2026-08-04): cycle_months is a set of every
    DISTINCT nonblank cycle-month Timestamp collected across all uploaded
    servicer files in one run (each already normalized to the 1st of the
    month by checks._infer_cycle_month, so direct set membership is a
    reliable way to detect a genuine month-level mismatch, not just a
    same-day formatting difference).

    Returns (cycle_month_or_None, error_message_or_None):
      - More than one distinct month present: (None, <clear error message>)
        — the caller must stop and return this as a 400 error rather than
        silently picking one, since combining different reporting periods
        into one monthly portfolio summary could be misleading.
      - Zero or exactly one distinct month: (that month or None, None) —
        a blank/undeterminable cycle month from one servicer does NOT
        count as a second distinct month; it simply isn't added to the set
        in the first place, so "one real month + some blanks" resolves
        cleanly to that one real month, never a false conflict."""
    if len(cycle_months) > 1:
        months_found = ", ".join(m.strftime("%B %Y") for m in sorted(cycle_months))
        return None, (
            f"The uploaded servicer files cover different reporting months ({months_found}). "
            f"All servicer files in one portfolio summary run must cover the same reporting "
            f"month — please upload files for a single reporting period at a time."
        )
    return next(iter(cycle_months), None), None


REPORT_ALWAYS_EXCLUDED_CHECK_IDS = {"PAY-01", "PAY-02", "DUP-01", "PORT-01"}
# included only when issues_found > 0. MEMO-01 belongs here rather than in the
# always-excluded tier: a clean memo file needs no tab, but an unrecognized
# Lender spelling or an orphaned memo row means real deferral evidence never
# reached the audit at all, which a reviewer has to see in the workbook itself
# -- the whole point of the check is that this failure is otherwise silent.
REPORT_CONDITIONAL_CHECK_IDS = {"POP-01", "CONT-01", "NEG-01", "MEMO-01"}


def _filter_report_results(all_results):
    """CLIENT-CONFIRMED (2026-08-06): decides which checks get their own
    tab in the detailed client-facing reconciliation workbook. Two tiers:

      - REPORT_ALWAYS_EXCLUDED_CHECK_IDS: excluded unconditionally,
        regardless of findings. PAY-01/DUP-01 are informational and can
        create noise even when they DO find something; PAY-02 isn't yet
        validated against the client's corrected SRF expectations;
        PORT-01 has its own separate Original Portfolio Summary
        workbook entirely outside this one, so it needs no tab here;
        its full result is in all_results regardless.
      - REPORT_CONDITIONAL_CHECK_IDS: excluded only when the check has
        NOTHING to report (issues_found == 0) — included with its own tab
        the moment it actually finds something. POP-01 doesn't need a tab
        for the common "no population differences" case, but a real
        mismatch is worth a client seeing directly. CONT-01 is the same
        shape: the clean case is the norm (validated on real June 2026
        data — zero findings across all eight servicer files), but a loan
        that carried a live balance last month and then vanished from the
        submission is exactly what a reviewer needs in the workbook.
        NEG-01's negative payments are important enough to see when they
        occur, but an empty PASS worksheet (the common case) adds little
        value.

    A missing worksheet under either tier is NEVER evidence a check
    failed to run — every check's real, complete result is still in
    all_results / the JSON response regardless of which tier it's in;
    this function only decides what the WORKBOOK shows."""
    return [
        r for r in all_results
        if r["id"].split("::")[0] not in REPORT_ALWAYS_EXCLUDED_CHECK_IDS
        and (
            r["id"].split("::")[0] not in REPORT_CONDITIONAL_CHECK_IDS
            or r.get("issues_found", 0) > 0
            # A result may carry findings it deliberately does NOT count as
            # issues (MEMO-01's unattributable file-level memo rows). Those
            # still have to reach the workbook, or a servicer with no memo
            # issues of its own would hide them entirely.
            or r.get("unattributable_count", 0) > 0
        )
    ]


def _flagged_loan_history_csv(soup_df, gp_numbers, columns=None, max_rows=1800):
    """Builds a CSV of SOUP's full monthly history (every Payment Month row)
    for every flagged loan — not a flat sample of the whole file, which would
    almost never happen to include a specific loan's full history. This is
    what lets the chatbot answer 'show me the payment history for loan X'
    questions about any loan the audit actually flagged."""
    if soup_df is None or soup_df.empty or "SBA GP Number" not in soup_df.columns or not gp_numbers:
        return ""
    mask = soup_df["SBA GP Number"].astype("string").str.strip().isin(gp_numbers)
    matched = soup_df[mask].copy()
    if matched.empty:
        return ""
    if "Payment Month" in matched.columns:
        matched = matched.sort_values(["SBA GP Number", "Payment Month"])
    if columns:
        keep = [c for c in columns if c in matched.columns]
        if keep:
            matched = matched[keep]
    matched = matched.head(max_rows)
    for col in matched.columns:
        if str(matched[col].dtype).startswith("datetime"):
            matched[col] = matched[col].astype("string")
    return matched.to_csv(index=False)


SOUP_CHAT_COLUMNS = [
    "SBA GP Number", "Servicer", "Payment Month", "Loan Status",
    "Delinquency Status", "Beginning Actual Balance", "Ending Actual Balance",
]
SERVICER_CHAT_COLUMNS = [
    "SBA GP Number", "Lender Loan Number", "Loan Status", "Delinquency Status",
    "Beginning Balance", "Ending Balance", "Payment Received",
    "Scheduled Payment", "Servicing Fee", "Payment Receipt Date",
]


def _truncate_csv(df, rows: int = 50, columns: list | None = None) -> str:
    if df is None or df.empty:
        return ""
    sample = df.head(rows).copy()
    if columns:
        keep = [c for c in columns if c in sample.columns]
        if keep:
            sample = sample[keep]
    for col in sample.columns:
        if str(sample[col].dtype).startswith("datetime"):
            sample[col] = sample[col].astype("string")
    return sample.to_csv(index=False)


def _cleanup_old_reports(max_age_seconds=3600):
    now = time.time()
    try:
        for name in os.listdir(REPORTS_DIR):
            path = os.path.join(REPORTS_DIR, name)
            if os.path.isfile(path) and now - os.path.getmtime(path) > max_age_seconds:
                os.remove(path)
    except OSError:
        pass


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/audit")
async def audit(
    soup_file: UploadFile = File(...),
    servicer_files: List[UploadFile] = File(...),
    memo_file: UploadFile = File(None),
):
    try:
        _cleanup_old_reports()

        soup_bytes = await soup_file.read()
        try:
            soup_df = file_readers.read_soup_file(soup_bytes)
        except ValueError as exc:
            return JSONResponse(status_code=400, content={
                "error": f"Could not read the SOUP Performance Dataset: {exc}",
            })
        del soup_bytes
        gc.collect()

        # Servicer Memo Summary is optional — STAT-01 degrades gracefully to
        # its pre-deferral-aware behavior without it (see
        # checks.run_servicer_reconciliation docstring).
        memo_df = None
        if memo_file is not None:
            memo_bytes = await memo_file.read()
            try:
                memo_df = file_readers.read_service_memo(memo_bytes)
            except ValueError as exc:
                return JSONResponse(status_code=400, content={
                    "error": f"Could not read the Servicer Memo Summary: {exc}",
                })
            del memo_bytes
            gc.collect()

            # CLIENT-CONFIRMED (2026-08-10): attribute each memo row to a loan
            # BEFORE any check or coverage classification runs, so STAT-01 and
            # MEMO-01 agree on which loan a row belongs to. Primary is the
            # existing SBA Loan ID -> SBA GP Number match; the only fallback is
            # SBA Loan ID -> Lender Loan Number, exact and unique, with the
            # matched loan supplying the servicer. Ambiguous or conflicting
            # rows are left unattributed rather than guessed at.
            #
            # Attribution only -- no data-quality rule is touched here, so a
            # row that becomes attributable keeps every defect it had.
            memo_df = file_readers.resolve_memo_attribution(
                memo_df, file_readers.build_memo_loan_index(soup_df),
            )

        all_results = []
        servicer_loan_counts = {}
        recommendations = []
        corrections_skipped = []
        # Per-servicer correction state stashed for POST /api/corrections/review:
        # original bytes + their SHA-256 + the eligible candidate set. Written
        # into SESSION_CACHE once the session_id is minted (end of this handler).
        correction_session_state = {}
        portfolio_summary_original_records = []
        # CLIENT-CONFIRMED (2026-08-11): the Original summary now reports loan
        # COUNTS beside dollars and totals PER SOUP DEAL, matching the Final
        # Audited presentation. Both extra lists are filled from the SAME
        # lender-only source as the dollars (see the loop below) -- one source
        # per table, never two.
        portfolio_summary_original_counts = []
        portfolio_summary_original_deals = []
        portfolio_summary_audited_records = []
        portfolio_summary_cycle_months = set()
        # Canonical soup_keys of the servicer files in THIS audit, used to
        # scope MEMO-01 below. The memo workbook covers every servicer; an
        # audit does not.
        uploaded_soup_keys = set()

        for uploaded in servicer_files:
            file_bytes = await uploaded.read()

            # FORMULA / CACHE PREFLIGHT — runs on the raw bytes BEFORE
            # Calamine, which exposes cached VALUES only and therefore cannot
            # tell a computed cell from a pasted literal, nor a populated
            # cache from an absent one.
            #
            # FORM-03 is a BLOCKING RE-INGESTION CONTROL. A workbook written
            # by openpyxl carries its formulas but no cached values, because
            # openpyxl does not evaluate. Calamine then reads those columns
            # as blank. Measured on Peoples_Bank_July_2026_CORRECTED.xlsx —
            # this application's own output — 60 of 60 rows returned blank
            # Loan Payment Instance, Last Payment Instance? and Servicing
            # Fee. Ingesting that file would sort postings by date instead of
            # instance, see no final-payment flag on any posting, and
            # evaluate payment sufficiency with a zero servicing fee.
            #
            # Corrected workbooks are OUTBOUND ONLY. This is the control that
            # enforces it: one must be opened and recalculated in Excel
            # before it can re-enter the pipeline.
            preflight = preflight_mod.preflight_workbook(
                file_bytes, filename=uploaded.filename,
            )
            if not preflight["ok"]:
                blocking = [f for f in preflight["findings"]
                            if f["severity"] == "FAIL"]
                return JSONResponse(status_code=400, content={
                    "error": (
                        f"'{uploaded.filename}' cannot be ingested: "
                        f"{blocking[0]['detail'] if blocking else 'preflight failed'}"
                    ),
                    "preflight": {
                        "filename": uploaded.filename,
                        "sheet": preflight["sheet"],
                        "data_rows": preflight["data_rows"],
                        "rows_affected": preflight["rows_affected"],
                        "findings": blocking,
                    },
                })

            try:
                soup_key, display_label, loan_terms_df, servicer_inputs_df, deal = file_readers.read_servicer_file(
                    file_bytes, filename=uploaded.filename
                )
            except ValueError as exc:
                return JSONResponse(status_code=400, content={
                    "error": f"Could not read servicer file '{uploaded.filename}': {exc}",
                })
            # file_bytes is deliberately NOT deleted here — the ORIGINAL bytes
            # are stashed in the session below so POST /api/corrections/review
            # can generate a corrected copy later, from the exact source
            # artifact the reviewer saw. They are held only for the session
            # TTL and are never mutated.

            # soup_key is what actually matches SOUP's Servicer column (e.g.
            # "Lendistry"); display_label is what shows up in results/report
            # tabs/summary (e.g. "Lendistry 2024") — these differ for files
            # like Lendistry's two yearly uploads, which both reconcile
            # against the same SOUP servicer but need distinct labels.
            uploaded_soup_keys.add(soup_key)
            soup_servicer_df = file_readers.filter_soup_by_servicer(soup_df, soup_key)

            results = checks_mod.run_servicer_reconciliation(
                display_label, loan_terms_df, servicer_inputs_df, soup_servicer_df,
                memo_df=memo_df, memo_lender_key=soup_key, deal=deal,
            )

            # RECOMMENDATIONS ONLY — no corrected workbook is written during
            # the audit. corrections.build_recommendations reads this file's
            # STAT-01 correction_data, applies the v1 eligibility allow-list,
            # and returns (recommendations, candidates, skipped). Generation
            # happens later, in POST /api/corrections/review, from the reviewer's
            # decisions. The correction_data/_cycle_month/audited_status keys
            # are stripped off the result below immediately afterward, so
            # nothing about HOW a recommendation was derived leaks into the
            # audit response beyond the recommendations themselves.
            stat01_result = next((r for r in results if r["id"].split("::")[0] == "STAT-01"), None)
            recs, candidates, skipped = corrections_mod.build_recommendations(
                stat01_result, display_label,
            )
            recommendations.extend(recs)
            for s in skipped:
                # Eligible-but-unreviewable (no valid posting row to write to)
                # — surfaced through the existing corrections_skipped channel,
                # with the servicer label added. Not reviewable, so never a
                # recommendation.
                corrections_skipped.append({"servicer": display_label, **s})

            # Stash this servicer's ORIGINAL bytes + hash + candidate set for
            # the review endpoint. SHA-256 captured now, alongside the bytes,
            # so the Correction Log can later prove which source artifact was
            # reviewed.
            #
            # CLIENT-CONFIRMED (2026-08-10): stored for EVERY servicer, not
            # only those with reviewable corrections. The Final Audited
            # Portfolio Summary is generated after Manual Review from these
            # bytes, and it must carry one row per uploaded servicer -- a
            # servicer the audit found nothing to correct on still holds
            # portfolio balances. Guarding this on `candidates` made those
            # servicers disappear from the summary entirely, which is a worse
            # failure than the memory it saves (~4.5 MB across the eight-bank
            # June set, held for the existing 1h TTL).
            #
            # excluded_gps carries the closed-at-zero suppression set forward,
            # captured HERE while suppressed_closed_loans is still on
            # stat01_result -- it is popped further below. Stashing the GP list
            # rather than recomputing it at review time keeps
            # loans_closed_at_zero the single derivation of that rule.
            correction_session_state[display_label] = {
                "original_bytes": file_bytes,
                "original_filename": uploaded.filename,
                "source_sha256": hashlib.sha256(file_bytes).hexdigest(),
                "institution": display_label,
                "cycle_month": (stat01_result or {}).get("_cycle_month"),
                "candidates": candidates,
                "excluded_gps": [
                    t["SBA GP Number"]
                    for t in ((stat01_result or {}).get("suppressed_closed_loans") or [])
                ],
            }
            del file_bytes
            gc.collect()

            # CLIENT-CONFIRMED (2026-08-04, revised): captured BEFORE the
            # _cycle_month pop below. Collect EVERY distinct nonblank cycle
            # month across servicers, not just the first — combining
            # different reporting periods into one monthly portfolio
            # summary could be misleading, so this must be detected and
            # stopped (see the check right after this loop), not silently
            # resolved by picking whichever servicer happened to be
            # processed first. A blank (None) cycle month from one servicer
            # does NOT count as a conflict on its own — it simply
            # contributes nothing to this set, so a mix of "one real month"
            # + "some blanks" still resolves cleanly to that one real month.
            if stat01_result is not None:
                cm = stat01_result.get("_cycle_month")
                if cm is not None:
                    portfolio_summary_cycle_months.add(cm)

            # CLIENT-CONFIRMED (2026-08-04): build this servicer's AUDITED
            # Portfolio Summary record BEFORE popping audited_status_by_loan
            # off stat01_result — this dict is the COMPLETE per-loan
            # dataset STAT-01 produces (see status_match_check's own
            # comments), reused here rather than recomputed, per the
            # explicit "do not duplicate delinquency logic" requirement.
            # Backend-internal only — audited_status_by_loan itself never
            # reaches the API response.
            if stat01_result is not None:
                audited_status_by_loan = stat01_result.get("audited_status_by_loan", {})
                # CLIENT-CONFIRMED (2026-08-10): loans suppressed by the
                # closed-at-zero rule are EXPLICITLY excluded from the audited
                # population -- neither their count nor their stale current
                # servicer balance contributes to any audited category. Passed
                # as an exclusion set rather than by withholding their audited
                # status, since a missing status falls back to the servicer's
                # own reported value.
                suppressed_trace = stat01_result.get("suppressed_closed_loans") or []
                portfolio_summary_audited_records.append(
                    checks_mod.audited_portfolio_summary_record(
                        servicer_inputs_df, audited_status_by_loan, servicer_name=display_label,
                        excluded_gps=[t["SBA GP Number"] for t in suppressed_trace],
                    )
                )

            if stat01_result is not None:
                # Suppression must be auditable even though it is deliberately
                # invisible to the client: log every withheld loan with the
                # evidence that justified it, then strip the trace from the
                # payload alongside the other internal-only sibling keys.
                for t in stat01_result.get("suppressed_closed_loans") or []:
                    print(
                        f"[closed-at-zero suppression] {t['Servicer']} GP {t['SBA GP Number']}: "
                        f"prior SOUP {t['Prior Closed Status']!r} at {t['Prior Ending Actual Balance']} "
                        f"on {t['Prior Closed Month']} (deal {t['Prior Closed Deal'] or 'n/a'}); "
                        f"servicer now reports "
                        f"{t['Current Servicer Loan Status']!r} at {t['Current Servicer Ending Balance']}; "
                        f"withheld from {', '.join(t['Checks Suppressed'])} and from all recommendations",
                        file=sys.stderr,
                    )
                stat01_result.pop("suppressed_closed_loans", None)
                stat01_result.pop("correction_data", None)
                stat01_result.pop("_cycle_month", None)
                stat01_result.pop("audited_status_by_loan", None)

            # CLIENT-CONFIRMED (2026-08-04, restored 2026-08-11, extended
            # 2026-08-11): collect this servicer's ORIGINAL Portfolio Summary
            # views -- dollars AND loan counts -- plus its SOUP deal, for the
            # workbook built after the loop. A servicer split across multiple
            # note-series files (e.g. "Lendistry 2025"/"Lendistry 2024")
            # naturally produces two separate entries here, one per file, since
            # each file gets its own run_servicer_reconciliation call with its
            # own display_label and its own deal.
            #
            # checks.original_portfolio_summary_views takes ONLY the lender's
            # submitted servicer_inputs_df: it has no parameter for an audited
            # status map, a verified override, a reviewer decision, a
            # correction, or a suppression set, so this record cannot carry an
            # audit-derived value no matter what happens above or below it.
            # That invariance is the whole point of this summary and is locked
            # in by test_original_portfolio_summary_invariance.py and
            # test_original_summary_final_format.py.
            #
            # ONE source for both tables, deliberately. The dollars are
            # asserted equal to PORT-01's own record on every real June file
            # (test_original_summary_final_format.py), so this is the same
            # tested figure -- but taking dollars from `results` and counts
            # from here would be two sources for one table, which is exactly
            # how two views of one population drift apart.
            original_views = checks_mod.original_portfolio_summary_views(
                servicer_inputs_df, servicer_name=display_label,
            )
            portfolio_summary_original_records.append(original_views["dollars"])
            portfolio_summary_original_counts.append(original_views["counts"])
            portfolio_summary_original_deals.append(deal)

            all_results.extend(results)
            servicer_loan_counts[display_label] = len(servicer_inputs_df)

        # CLIENT-CONFIRMED (2026-08-04): if the uploaded servicer files
        # cover more than one distinct reporting month, STOP the request
        # entirely with a clear error rather than silently naming the
        # Portfolio Summary workbook after whichever month was found first,
        # or falling back to a neutral "Mixed_Periods" filename — combining
        # different reporting periods into one monthly portfolio summary
        # could be misleading, and the client's explicit preference is to
        # stop and warn rather than guess. This check runs BEFORE either
        # workbook is generated, so no partial/misleading output is ever
        # produced for a mixed-period upload. See
        # _resolve_portfolio_summary_cycle_month for the (unit-tested) logic.
        portfolio_summary_cycle_month, mixed_period_error = _resolve_portfolio_summary_cycle_month(
            portfolio_summary_cycle_months
        )
        if mixed_period_error:
            return JSONResponse(status_code=400, content={"error": mixed_period_error})

        # MEMO-01 — Servicer Memo Coverage. Runs ONCE per audit, after the
        # per-servicer loop, because the memo workbook is a single global file
        # rather than per-servicer input: inside the loop it would re-sweep the
        # same memo_df once per uploaded servicer and report every finding N
        # times.
        #
        # This is a STANDALONE DIAGNOSTIC over the memo file. It does not feed
        # STAT-01, corrections.py, or any calculation — it only reports. It
        # exists because checks._lookup_deferral runs per loan and therefore
        # cannot see a memo row whose loan ID is in no population: no loan
        # iteration ever reaches that row, so no per-loan warning can fire for
        # it. Today such a row is indistinguishable from "this loan has no
        # deferral". Sweeping the memo file is the only direction that sees it.
        #
        # Two different scopes, deliberately:
        #
        #   population = every SBA GP Number in the SOUP file (all servicers).
        #     "lookup-failed" means "this loan ID exists in no population", so
        #     narrowing it to this run's servicers would reclassify every other
        #     servicer's memo rows as defects -- measured at 195 false findings
        #     instead of 22 on the delivered file.
        #   uploaded_soup_keys = only the servicers in THIS audit. Scopes which
        #     rows are reported as this audit's issues, by Lender attribution.
        #
        # Attribution and population answer different questions and must not be
        # conflated. Rows attributable to NO servicer are still reported, but
        # counted separately (unattributable_count) so they are never presented
        # as an uploaded servicer's issues.
        if memo_df is not None:
            memo_result, _memo_counts = memo_audit_mod.build_memo_coverage_result(
                memo_df,
                soup_df["SBA GP Number"].dropna().tolist() if "SBA GP Number" in soup_df.columns else [],
                uploaded_soup_keys=uploaded_soup_keys,
            )
            if memo_result is not None:
                all_results.append(memo_result)

        summary = checks_mod.summarize(all_results, len(soup_df), servicer_loan_counts)

        # Client-facing Excel report: BAL-01 and STAT-01 are the primary
        # controls (balance continuity, deferral-aware delinquency status).
        # Everything else here still runs in full above (all_results is
        # untouched) and still feeds the JSON response / chatbot / frontend
        # dashboard — _filter_report_results only controls which checks
        # get a tab in the downloadable Excel workbook. See that
        # function's docstring for the full per-check rationale.
        report_results = _filter_report_results(all_results)

        # The Excel workbook's own Summary tab banner ("Total issues: N")
        # must reflect only what the workbook actually contains (BAL-01 +
        # STAT-01), not the full 7-check total — otherwise the banner would
        # show issue counts (e.g. from PAY-02) that no visible tab explains.
        # This is intentionally a SEPARATE summary from `summary` above:
        # `summary` still reflects all_results and is what the JSON
        # response / frontend dashboard scorecards / chatbot context use,
        # unchanged.
        report_summary = checks_mod.summarize(report_results, len(soup_df), servicer_loan_counts)

        excel_buffer = io.BytesIO()
        report_mod.generate_report_bytes(report_results, report_summary, excel_buffer)
        excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode("utf-8")

        # CLIENT-CONFIRMED (2026-08-04, revised 2026-08-11): TWO separate
        # downloadable Portfolio Summary workbooks, entirely independent of the
        # detailed reconciliation workbook above and of each other — both built
        # from the SAME per-servicer records already collected during the loop
        # (Original from checks.original_portfolio_summary_views, which takes
        # the lender's submitted frame and nothing else; Audited from
        # audited_portfolio_summary_record, which reuses STAT-01's real
        # calculated_dq via audited_status_by_loan). Neither is re-derived from
        # raw file data here. Each is only generated when at least one
        # servicer's record was actually produced (always true for a normal
        # successful run; guarded purely so an edge case upstream can't produce
        # an empty, meaningless workbook).
        #
        # The two workbooks now share a LAYOUT (the stacked-section, per-SOUP-
        # deal presentation) while keeping entirely separate DATA SOURCES. The
        # Audited workbook deliberately keeps the older two-sheet format, which
        # is why generate_portfolio_summary_bytes is still used below.
        #
        # The ORIGINAL Portfolio Summary was removed from the client-facing
        # output on 2026-08-10 and RESTORED on 2026-08-11 as the PROMINENT
        # top-level download, with the Audited workbook retained as a secondary
        # download. PORT-01 itself never stopped running throughout, and is
        # unchanged by either decision. The three layers the client works
        # through are now, in order:
        #
        #     Original Portfolio Summary       what the lender submitted
        #     Audit Findings and Recommendations   what the audit calculated
        #     Final Audited Portfolio Summary  what the reviewer approved
        #                                      (POST /api/corrections/review)
        #
        # Both files keep their existing names, so a client who has been
        # filing either one keeps a stable filename convention.
        portfolio_summary_original_b64 = None
        portfolio_summary_original_filename_value = None
        if portfolio_summary_original_records:
            original_buffer = io.BytesIO()
            # CLIENT-CONFIRMED (2026-08-11): the Original summary is presented
            # in the SAME format as the Final Audited one -- one sheet, the
            # COUNT / BALANCE % / BALANCE ($) sections, and a Total per SOUP
            # deal rather than a single combined figure. Presentation and
            # aggregation only: group_portfolio_summary_by_deal and
            # generate_final_portfolio_summary_bytes both consume records that
            # are already computed and never see a data source, so reusing them
            # here cannot introduce an audited or reviewed value. The records
            # themselves come from original_portfolio_summary_views, which is
            # lender-only by signature.
            #
            # The FILENAME is deliberately unchanged (label="Original"), so a
            # client who has been filing this workbook keeps the same
            # convention -- only its contents are laid out differently.
            original_blocks = checks_mod.group_portfolio_summary_by_deal(
                portfolio_summary_original_records, portfolio_summary_original_counts,
                portfolio_summary_original_deals,
            )
            report_mod.generate_final_portfolio_summary_bytes(
                original_blocks, original_buffer,
                cycle_month=portfolio_summary_cycle_month,
                title="Original Portfolio Summary",
            )
            portfolio_summary_original_b64 = base64.b64encode(original_buffer.getvalue()).decode("utf-8")
            portfolio_summary_original_filename_value = report_mod.portfolio_summary_filename(
                portfolio_summary_cycle_month, label="Original"
            )

        portfolio_summary_audited_b64 = None
        portfolio_summary_audited_filename_value = None
        if portfolio_summary_audited_records:
            audited_buffer = io.BytesIO()
            report_mod.generate_portfolio_summary_bytes(portfolio_summary_audited_records, audited_buffer)
            portfolio_summary_audited_b64 = base64.b64encode(audited_buffer.getvalue()).decode("utf-8")
            portfolio_summary_audited_filename_value = report_mod.portfolio_summary_filename(
                portfolio_summary_cycle_month, label="Audited"
            )

        flagged_gps = _flagged_gp_numbers(all_results)  # already capped to MAX_FLAGGED_LOANS_FOR_HISTORY
        flagged_loan_history_csv = _flagged_loan_history_csv(soup_df, flagged_gps, columns=SOUP_CHAT_COLUMNS)

        # soup_csv and servicer_csvs are dropped here: both were flat samples
        # of the full files, now redundant with lookup_loan_history and
        # search_accounts_by_status querying the live soup_df on demand.
        chat_context = {
            "summary": summary,
            "results": _cap_results_for_chat(all_results),
            "flagged_loan_history_csv": flagged_loan_history_csv,
        }

        session_id = str(uuid.uuid4())
        _cleanup_old_sessions()
        SESSION_CACHE[session_id] = {
            "soup_df": soup_df,
            "created": time.time(),
            # Consumed by POST /api/corrections/review (process-local; see the
            # SESSION_CACHE definition for the single-worker caveat).
            "corrections": correction_session_state,
        }

        # No corrected workbook is generated here anymore — the audit returns
        # RECOMMENDATIONS the reviewer acts on, and generation happens in POST
        # /api/corrections/review. corrections_skipped still carries items that
        # are not reviewable (eligible but no valid posting row to write to).
        return JSONResponse({
            "results": all_results,
            "summary": summary,
            "report_b64": excel_b64,
            "report_filename": report_mod.report_filename(),
            "portfolio_summary_original_b64": portfolio_summary_original_b64,
            "portfolio_summary_original_filename": portfolio_summary_original_filename_value,
            "portfolio_summary_audited_b64": portfolio_summary_audited_b64,
            "portfolio_summary_audited_filename": portfolio_summary_audited_filename_value,
            "chat_context": chat_context,
            "session_id": session_id,
            "recommendations": recommendations,
            "corrections_skipped": corrections_skipped,
        })

    except Exception as exc:
        traceback.print_exc()
        return JSONResponse(
            status_code=400,
            content={
                "error": "We couldn't process the audit. Please check that the SOUP file and "
                         "servicer file(s) are valid and match the expected templates.",
                "detail": str(exc),
            },
        )


@app.post("/api/corrections/review")
async def corrections_review(payload: dict):
    """Second half of the review-then-generate workflow. Takes a reviewer's
    decisions on the recommendations from a prior /api/audit run and, for
    each servicer, generates a corrected workbook from that servicer's
    ORIGINAL uploaded bytes (held in the session) — writing ONLY approved
    values and recording every decision in the Correction Log.

    Nothing was written during the audit; this is the only place a corrected
    workbook is produced. All safety controls (domain-restricted approved
    values, the terminal-Loan-Status write block, per-row GP verification)
    live in corrections.py and cannot be bypassed by anything in the payload.
    """
    try:
        session_id = payload.get("session_id")
        reviewed_by = (payload.get("reviewed_by") or "").strip()
        reviewed_on = payload.get("reviewed_on")
        decisions = payload.get("decisions") or []

        if not decisions:
            return JSONResponse(status_code=400, content={
                "error": "No decisions were submitted. Approve, reject, or return at least one "
                         "recommendation before recording.",
            })

        # An approval that cannot be attributed is not an audit record. The
        # name is self-declared (no auth) and recorded as unverified, but it
        # must be present — this mirrors the review page's own submit gate.
        if not reviewed_by:
            return JSONResponse(status_code=400, content={
                "error": "A reviewer name is required to record corrections.",
            })

        _cleanup_old_sessions()
        session = SESSION_CACHE.get(session_id) if session_id else None
        if session is None:
            # Either never existed or aged out of the 1h TTL. The bytes and
            # candidates are gone; decisions cannot be recovered, and the
            # only honest path is to re-run the audit.
            return JSONResponse(status_code=410, content={
                "error": "This review session has expired or is no longer available. Please "
                         "re-upload the files and re-run the audit; the decisions cannot be "
                         "recovered.",
            })

        correction_state = session.get("corrections") or {}
        # Checks for reviewable CANDIDATES, not merely for a session entry:
        # since 2026-08-10 every servicer is stashed (the Final Audited
        # Portfolio Summary needs one row each), so a non-empty
        # correction_state no longer implies anything was reviewable. This
        # keeps the original meaning of the guard exactly.
        if not any((entry.get("candidates") or []) for entry in correction_state.values()):
            return JSONResponse(status_code=400, content={
                "error": "This audit produced no reviewable corrections.",
            })

        # Route each decision to the servicer that owns its recommendation id.
        id_to_servicer = {
            cand["id"]: servicer
            for servicer, entry in correction_state.items()
            for cand in entry["candidates"]
        }
        decisions_by_servicer = {}
        unknown_ids = []
        for d in decisions:
            servicer = id_to_servicer.get(d.get("id"))
            if servicer is None:
                unknown_ids.append(d.get("id"))
                continue
            decisions_by_servicer.setdefault(servicer, []).append(d)

        corrected_workbooks = []
        corrections_not_applied = []
        applied_total = rejected_total = returned_total = 0
        # {servicer: {(gp, field): approved_value}} -- the VERIFIED layer, the
        # single source both post-review artifacts are built from.
        verified_overrides_by_servicer = {}

        for servicer, servicer_decisions in decisions_by_servicer.items():
            entry = correction_state[servicer]
            outcome = corrections_mod.generate_reviewed_workbook(
                entry["original_bytes"],
                entry["original_filename"],
                entry["institution"],
                entry["candidates"],
                servicer_decisions,
                reviewed_by,
                entry["source_sha256"],
                entry["cycle_month"],
                reviewed_on=reviewed_on,
            )

            verified_overrides_by_servicer[servicer] = outcome.get("verified_overrides") or {}

            applied_total += outcome["applied_count"]
            rejected_total += outcome["rejected_count"]
            returned_total += outcome["returned_count"]
            unknown_ids.extend(outcome.get("unknown_ids") or [])

            for not_applied in outcome.get("corrections_not_applied") or []:
                corrections_not_applied.append({
                    "servicer": servicer,
                    "SBA GP Number": not_applied.get("gp"),
                    "field": not_applied.get("field"),
                    "reason": not_applied.get("reason"),
                })

            if outcome["file_bytes"]:
                corrected_workbooks.append({
                    "servicer": servicer,
                    "filename": outcome["filename"],
                    "file_b64": base64.b64encode(outcome["file_bytes"]).decode("utf-8"),
                    "applied_count": outcome["applied_count"],
                    "rejected_count": outcome["rejected_count"],
                    "returned_count": outcome["returned_count"],
                })
            elif outcome["skip_reason"]:
                corrections_not_applied.append({
                    "servicer": servicer,
                    "reason": outcome["skip_reason"],
                })

        # ---- Final Audited Portfolio Summary --------------------------------
        # CLIENT-CONFIRMED (2026-08-10): the third layer, generated ONLY here,
        # only after Manual Review, from
        #
        #     original lender data + verified_overrides + closed-at-zero suppression
        #
        # Built with checks_mod.verified_portfolio_summary_record, which reuses
        # _build_portfolio_summary_record -- there is no separate calculation
        # engine, and the Original (PORT-01) and Audited summaries are
        # untouched. It shares verified_overrides with the corrected workbook
        # above, so the two artifacts cannot disagree.
        #
        # Iterates correction_state, NOT decisions_by_servicer: every uploaded
        # servicer gets exactly one row, including servicers the reviewer made
        # no decisions on (their overrides are simply empty, so their row is
        # PORT-01 minus any suppressed loans). Dict insertion order preserves
        # upload order, matching the Audited workbook's row order.
        portfolio_summary_final_records = []
        portfolio_summary_final_counts = []
        portfolio_summary_final_deals = []
        final_cycle_months = set()
        for servicer, entry in correction_state.items():
            try:
                # The 5th return value is this file's own SOUP deal, parsed
                # from its "Note Series:" cell by
                # file_readers.extract_deal_from_note_series -- the same value
                # POP-01 already scopes on. Captured here (it used to be
                # discarded) so the Final Audited Portfolio Summary can total
                # per SOUP deal. Nothing is inferred from the filename and no
                # servicer name is hardcoded.
                _, _, _, servicer_inputs_df, deal = file_readers.read_servicer_file(
                    entry["original_bytes"], filename=entry.get("original_filename"),
                )
            except Exception:
                # One unreadable servicer must not lose the whole summary; the
                # others are still reported. Same isolation principle the
                # per-servicer generation loop above uses.
                traceback.print_exc()
                continue
            # ONE verified population per servicer, feeding all three views --
            # Count, Balance % and Balance ($). verified_portfolio_summary_views
            # builds it once; constructing it separately per view is the
            # failure mode this call shape exists to prevent.
            views = checks_mod.verified_portfolio_summary_views(
                servicer_inputs_df,
                verified_overrides_by_servicer.get(servicer) or {},
                servicer_name=servicer,
                excluded_gps=entry.get("excluded_gps") or [],
            )
            portfolio_summary_final_records.append(views["dollars"])
            portfolio_summary_final_counts.append(views["counts"])
            portfolio_summary_final_deals.append(deal)
            if entry.get("cycle_month") is not None:
                final_cycle_months.add(entry["cycle_month"])

        portfolio_summary_final_b64 = None
        portfolio_summary_final_filename_value = None
        if portfolio_summary_final_records:
            final_cycle_month, _ = _resolve_portfolio_summary_cycle_month(final_cycle_months)
            final_buffer = io.BytesIO()
            # CLIENT-CONFIRMED (2026-08-10): one Total PER SOUP DEAL, each
            # followed by its own servicers -- never a single figure combining
            # deals, since a 2024-deal and a 2025-deal servicer belong to
            # different securitizations. Percentages are recomputed against
            # each deal's own Outstanding, so every deal reads 100.00%.
            # A servicer whose SRF declares no usable Note Series lands in an
            # explicit "Unassigned Total" rather than being dropped.
            final_blocks = checks_mod.group_portfolio_summary_by_deal(
                portfolio_summary_final_records, portfolio_summary_final_counts,
                portfolio_summary_final_deals,
            )
            # cycle_month feeds the heading's "Reporting Period" line only --
            # the same value the filename already uses. It never touches a figure.
            report_mod.generate_final_portfolio_summary_bytes(
                final_blocks, final_buffer, cycle_month=final_cycle_month,
            )
            portfolio_summary_final_b64 = base64.b64encode(final_buffer.getvalue()).decode("utf-8")
            # Underscored: the label is interpolated straight into the
            # filename, so "Final Audited" would yield
            # "Portfolio_Summary_Final Audited_June_2026.xlsx" -- a space, and
            # inconsistent with the Original/Audited files beside it. Fixed
            # here at the call site rather than by teaching the shared filename
            # helper to rewrite labels.
            portfolio_summary_final_filename_value = report_mod.portfolio_summary_filename(
                final_cycle_month, label="Final_Audited"
            )

        return JSONResponse({
            "corrected_workbooks": corrected_workbooks,
            "portfolio_summary_final_b64": portfolio_summary_final_b64,
            "portfolio_summary_final_filename": portfolio_summary_final_filename_value,
            "summary": {
                "applied": applied_total,
                "rejected": rejected_total,
                "returned": returned_total,
                "not_applied": len(corrections_not_applied),
            },
            "corrections_not_applied": corrections_not_applied,
            "unknown_ids": [i for i in unknown_ids if i is not None],
        })

    except Exception as exc:
        traceback.print_exc()
        return JSONResponse(status_code=500, content={
            "error": "We couldn't generate the corrected workbook(s). Please try again.",
            "detail": str(exc),
        })


def _parse_markdown_tables(text):
    """Finds every markdown table in a chat reply and returns a list of
    (headers, rows) tuples. A markdown table is a block of lines starting
    with '|', where the second line is the '---' separator row."""
    lines = text.split("\n")
    tables = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith("|") and i + 1 < len(lines):
            sep_line = lines[i + 1].strip()
            if sep_line.startswith("|") and set(sep_line.replace("|", "").replace(" ", "")) <= set("-:"):
                headers = [c.strip() for c in line.strip("|").split("|")]
                rows = []
                j = i + 2
                while j < len(lines) and lines[j].strip().startswith("|"):
                    cells = [c.strip() for c in lines[j].strip().strip("|").split("|")]
                    if len(cells) == len(headers):
                        rows.append(cells)
                    j += 1
                if rows:
                    tables.append((headers, rows))
                i = j
                continue
        i += 1
    return tables


def _markdown_tables_to_excel_b64(tables):
    """Builds a formatted Excel workbook from one or more parsed markdown
    tables (one sheet per table), returns (base64_string, filename), or
    (None, None) if there's nothing to export."""
    if not tables:
        return None, None

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    HEADER_FILL = PatternFill("solid", fgColor="0C447C")
    HEADER_FONT = Font(color="FFFFFF", bold=True)
    THIN = Side(style="thin", color="E2E8F0")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    wb = Workbook()
    wb.remove(wb.active)

    for idx, (headers, rows) in enumerate(tables):
        sheet_name = f"Table {idx + 1}" if len(tables) > 1 else "Data"
        ws = wb.create_sheet(sheet_name[:31])
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.border = BORDER
        ws.freeze_panes = "A2"
        for r_offset, row in enumerate(rows):
            for c_idx, value in enumerate(row, start=1):
                # try to preserve numbers as numbers, not text, for anything
                # that looks like a plain number (skip $ and % formatted values)
                cell_value = value
                if value and not any(ch in value for ch in "$%,"):
                    try:
                        cell_value = float(value) if "." in value else int(value)
                    except ValueError:
                        pass
                ws.cell(row=r_offset + 2, column=c_idx, value=cell_value).border = BORDER
        for col_idx, header in enumerate(headers, start=1):
            max_len = len(str(header))
            for row in rows:
                v = row[col_idx - 1] if col_idx - 1 < len(row) else ""
                max_len = max(max_len, len(str(v)))
            ws.column_dimensions[get_column_letter(col_idx)].width = max(10, min(40, max_len + 2))

    buf = io.BytesIO()
    wb.save(buf)
    b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
    filename = f"FalconBridge_Chat_Export_{int(time.time())}.xlsx"
    return b64, filename


@app.post("/api/chat")
async def chat(payload: dict):
    try:
        message = payload.get("message", "")
        history = payload.get("history", [])
        audit_context = payload.get("audit_context", {})
        session_id = payload.get("session_id")
        if not message.strip():
            return JSONResponse(status_code=400, content={"error": "Empty message."})

        session_entry = SESSION_CACHE.get(session_id) if session_id else None
        soup_df = session_entry["soup_df"] if session_entry else None

        reply = chatbot_mod.chat(message, history, audit_context, soup_df=soup_df)

        tables = _parse_markdown_tables(reply)
        excel_b64, excel_filename = _markdown_tables_to_excel_b64(tables)

        response = {"reply": reply}
        if excel_b64:
            response["excel_b64"] = excel_b64
            response["excel_filename"] = excel_filename
        return response
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse(
            status_code=500,
            content={"error": "The assistant is unavailable right now. Please try again.",
                     "detail": str(exc)},
        )
