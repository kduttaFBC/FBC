"""
FalconBridge Audit Intelligence — FastAPI backend.

Endpoints:
  POST /api/audit            reconcile SOUP vs. each uploaded servicer SRF file
  POST /api/chat             chat with FalconBridge AI about the audit
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

# Holds the full SOUP dataset per audit session so the chatbot can look up
# ANY account on demand (flagged or not), not just whatever was pre-loaded
# into the chat context. Sessions expire after SESSION_TTL_SECONDS.
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
        corrected_workbooks = []
        corrections_skipped = []
        portfolio_summary_records = []
        portfolio_summary_cycle_months = set()

        for uploaded in servicer_files:
            file_bytes = await uploaded.read()
            try:
                soup_key, display_label, loan_terms_df, servicer_inputs_df, deal = file_readers.read_servicer_file(
                    file_bytes, filename=uploaded.filename
                )
            except ValueError as exc:
                return JSONResponse(status_code=400, content={
                    "error": f"Could not read servicer file '{uploaded.filename}': {exc}",
                })
            # file_bytes is deliberately NOT deleted yet — the correction
            # engine below needs the file's ORIGINAL bytes to build a
            # corrected copy. Released right after that step, same as
            # before, just a few lines later.

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

            # Corrected servicer workbook generation is entirely a backend
            # implementation detail (see backend/corrections.py) — it reads
            # this file's STAT-01 result's correction_data/_cycle_month,
            # decides what's eligible under the v1 allow-list, and returns
            # either a corrected workbook or a reason it didn't. Those two
            # keys are stripped off the result below immediately afterward,
            # so nothing about HOW a correction was decided ever reaches the
            # API response — the audit API stays exactly "audit results."
            stat01_result = next((r for r in results if r["id"].split("::")[0] == "STAT-01"), None)
            correction_outcome = corrections_mod.generate_corrected_workbook(
                file_bytes, uploaded.filename, display_label, stat01_result,
            )
            del file_bytes
            gc.collect()

            if correction_outcome["file_bytes"]:
                corrected_workbooks.append({
                    "servicer": display_label,
                    "filename": correction_outcome["filename"],
                    "file_b64": base64.b64encode(correction_outcome["file_bytes"]).decode("utf-8"),
                    "correction_count": correction_outcome["correction_count"],
                })
            elif correction_outcome["skip_reason"]:
                corrections_skipped.append({
                    "servicer": display_label,
                    "reason": correction_outcome["skip_reason"],
                })

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

            if stat01_result is not None:
                stat01_result.pop("correction_data", None)
                stat01_result.pop("_cycle_month", None)

            # CLIENT-CONFIRMED (2026-08-04): collect this servicer's PORT-01
            # record for the SEPARATE Portfolio Summary workbook (see
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
                portfolio_summary_records.append(port01_result["records"][0])

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

        summary = checks_mod.summarize(all_results, len(soup_df), servicer_loan_counts)

        # Client-facing Excel report: BAL-01 and STAT-01 are the primary
        # controls (balance continuity, deferral-aware delinquency status).
        # Everything else here still runs in full above (all_results is
        # untouched) and still feeds the JSON response / chatbot / frontend
        # dashboard — this set ONLY controls which checks get a tab in the
        # downloadable Excel workbook.
        #   POP-01: important control, kept running, but doesn't need its
        #     own detail tab (174-loan Servicer + Deal population match is
        #     unchanged — see population_check).
        #   PAY-01, DUP-01, NEG-01: diagnostic/internal — informational,
        #     not yet part of the client's core monthly SRF QC scope.
        #   PAY-02: business rule (Scheduled Payment - (Payment Received +
        #     Servicing Fee)) not yet validated against the client's
        #     corrected SRF workbook — held out of the client-facing report
        #     until that validation happens, NOT rewritten or disabled.
        #   PORT-01: CLIENT-CONFIRMED (2026-08-04) — deliberately excluded
        #     from THIS workbook entirely. It gets its own SEPARATE
        #     downloadable Portfolio Summary workbook instead (see
        #     portfolio_summary_records / generate_portfolio_summary_bytes
        #     below), not a tab inside the detailed reconciliation workbook.
        REPORT_EXCLUDED_CHECK_IDS = {"POP-01", "PAY-01", "PAY-02", "DUP-01", "NEG-01", "PORT-01"}
        report_results = [
            r for r in all_results
            if r["id"].split("::")[0] not in REPORT_EXCLUDED_CHECK_IDS
        ]

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

        # CLIENT-CONFIRMED (2026-08-04): a SEPARATE downloadable workbook,
        # entirely independent of the detailed reconciliation workbook
        # above — built from the SAME PORT-01 records already collected
        # per servicer during the loop, never re-derived from raw file
        # data. Only generated when at least one servicer's PORT-01 record
        # was actually produced (always true for a normal successful run
        # with 1+ servicer files; guarded here purely so an edge case
        # upstream — e.g. every servicer file failing to parse before this
        # point — can't produce an empty, meaningless workbook).
        portfolio_summary_b64 = None
        portfolio_summary_filename_value = None
        if portfolio_summary_records:
            portfolio_summary_buffer = io.BytesIO()
            report_mod.generate_portfolio_summary_bytes(portfolio_summary_records, portfolio_summary_buffer)
            portfolio_summary_b64 = base64.b64encode(portfolio_summary_buffer.getvalue()).decode("utf-8")
            portfolio_summary_filename_value = report_mod.portfolio_summary_filename(portfolio_summary_cycle_month)

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
        SESSION_CACHE[session_id] = {"soup_df": soup_df, "created": time.time()}

        return JSONResponse({
            "results": all_results,
            "summary": summary,
            "report_b64": excel_b64,
            "report_filename": report_mod.report_filename(),
            "portfolio_summary_b64": portfolio_summary_b64,
            "portfolio_summary_filename": portfolio_summary_filename_value,
            "chat_context": chat_context,
            "session_id": session_id,
            "corrected_workbooks": corrected_workbooks,
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
