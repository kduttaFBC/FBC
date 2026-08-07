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
REPORT_CONDITIONAL_CHECK_IDS = {"POP-01", "NEG-01", "MEMO-01"}


def _filter_report_results(all_results):
    """CLIENT-CONFIRMED (2026-08-06): decides which checks get their own
    tab in the detailed client-facing reconciliation workbook. Two tiers:

      - REPORT_ALWAYS_EXCLUDED_CHECK_IDS: excluded unconditionally,
        regardless of findings. PAY-01/DUP-01 are informational and can
        create noise even when they DO find something; PAY-02 isn't yet
        validated against the client's corrected SRF expectations;
        PORT-01 has its own separate Original/Audited Portfolio Summary
        workbooks entirely outside this one.
      - REPORT_CONDITIONAL_CHECK_IDS: excluded only when the check has
        NOTHING to report (issues_found == 0) — included with its own tab
        the moment it actually finds something. POP-01 doesn't need a tab
        for the common "no population differences" case, but a real
        mismatch is worth a client seeing directly. NEG-01's negative
        payments are important enough to see when they occur, but an
        empty PASS worksheet (the common case) adds little value.

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

        all_results = []
        servicer_loan_counts = {}
        recommendations = []
        corrections_skipped = []
        # Per-servicer correction state stashed for POST /api/corrections/review:
        # original bytes + their SHA-256 + the eligible candidate set. Written
        # into SESSION_CACHE once the session_id is minted (end of this handler).
        correction_session_state = {}
        portfolio_summary_original_records = []
        portfolio_summary_audited_records = []
        portfolio_summary_cycle_months = set()

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
            # reviewed. Only stored when there is something reviewable.
            if candidates:
                correction_session_state[display_label] = {
                    "original_bytes": file_bytes,
                    "original_filename": uploaded.filename,
                    "source_sha256": hashlib.sha256(file_bytes).hexdigest(),
                    "institution": display_label,
                    "cycle_month": (stat01_result or {}).get("_cycle_month"),
                    "candidates": candidates,
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
                portfolio_summary_audited_records.append(
                    checks_mod.audited_portfolio_summary_record(
                        servicer_inputs_df, audited_status_by_loan, servicer_name=display_label
                    )
                )

            if stat01_result is not None:
                stat01_result.pop("correction_data", None)
                stat01_result.pop("_cycle_month", None)
                stat01_result.pop("audited_status_by_loan", None)

            # CLIENT-CONFIRMED (2026-08-04): collect this servicer's PORT-01
            # record for the ORIGINAL Portfolio Summary workbook (see
            # report_mod.generate_portfolio_summary_bytes below) — built
            # entirely from the same tested PORT-01 records already
            # produced above, never re-derived from raw file data. A
            # servicer split across multiple note-series files (e.g.
            # "Lendistry 2025"/"Lendistry 2024") naturally produces two
            # separate records here, one per file, since each file gets its
            # own run_servicer_reconciliation call with its own
            # display_label.
            port01_result = next((r for r in results if r["id"].split("::")[0] == "PORT-01"), None)
            if port01_result is not None and port01_result["records"]:
                portfolio_summary_original_records.append(port01_result["records"][0])

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
        # Population = every SBA GP Number in the SOUP file (all servicers),
        # not just this run's uploads: a memo row belonging to a servicer that
        # wasn't uploaded today is a legitimately-absent reconciliation, not a
        # defective memo row, and must not be reported as one.
        if memo_df is not None:
            memo_result, _memo_counts = memo_audit_mod.build_memo_coverage_result(
                memo_df,
                soup_df["SBA GP Number"].dropna().tolist() if "SBA GP Number" in soup_df.columns else [],
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

        # CLIENT-CONFIRMED (2026-08-04, revised): TWO separate downloadable
        # workbooks, entirely independent of the detailed reconciliation
        # workbook above and of each other — built from the SAME per-
        # servicer records already collected during the loop (Original
        # from PORT-01's own records; Audited from
        # audited_portfolio_summary_record, which reuses STAT-01's real
        # calculated_dq via audited_status_by_loan). Neither is re-derived
        # from raw file data here. Only generated when at least one
        # servicer's record was actually produced (always true for a
        # normal successful run; guarded purely so an edge case upstream
        # can't produce an empty, meaningless workbook).
        portfolio_summary_original_b64 = None
        portfolio_summary_original_filename_value = None
        if portfolio_summary_original_records:
            original_buffer = io.BytesIO()
            report_mod.generate_portfolio_summary_bytes(portfolio_summary_original_records, original_buffer)
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
        if not correction_state:
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

        return JSONResponse({
            "corrected_workbooks": corrected_workbooks,
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
