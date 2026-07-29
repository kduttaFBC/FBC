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


load_dotenv(override=False)

app = FastAPI(title="FalconBridge Audit Intelligence")

ALLOWED_ORIGINS = [
    "https://falconbridge-audit-frontend-5tdw.onrender.com",  # production frontend
    "http://localhost:5173",                                  # local Vite dev server
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

REPORTS_DIR = "/tmp/reports"
os.makedirs(REPORTS_DIR, exist_ok=True)

# Holds the full SOUP dataset per audit session so the chatbot can look up
# ANY account on demand (flagged or not), not just whatever was pre-loaded
# into the chat context. Sessions expire after SESSION_TTL_SECONDS, and the
# cache is additionally hard-capped at MAX_SESSIONS entries (oldest evicted
# first) so a burst of public traffic can't grow memory unbounded.
SESSION_CACHE = {}
SESSION_TTL_SECONDS = 3600
MAX_SESSIONS = 10


def _cleanup_old_sessions():
    now = time.time()
    expired = [sid for sid, entry in SESSION_CACHE.items() if now - entry["created"] > SESSION_TTL_SECONDS]
    for sid in expired:
        del SESSION_CACHE[sid]

    overflow = len(SESSION_CACHE) - MAX_SESSIONS
    if overflow > 0:
        oldest_first = sorted(SESSION_CACHE.items(), key=lambda kv: kv[1]["created"])
        for sid, _ in oldest_first[:overflow]:
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
            del file_bytes
            gc.collect()

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
            all_results.extend(results)
            servicer_loan_counts[display_label] = len(servicer_inputs_df)

        summary = checks_mod.summarize(all_results, len(soup_df), servicer_loan_counts)

        # The downloadable Excel workbook is now a faithful export of the
        # SAME all_results the JSON response / frontend dashboard / chatbot
        # use — every check gets its own tab, with the same `summary`
        # banner. (Previously a REPORT_EXCLUDED_CHECK_IDS set held POP-01,
        # PAY-01, PAY-02, DUP-01, and NEG-01 out of the Excel export only;
        # that exclusion was a stale holdover from an earlier point in the
        # project and has been removed per client request — nothing about
        # how those checks are calculated changed, only whether they get an
        # Excel tab.)
        excel_buffer = io.BytesIO()
        report_mod.generate_report_bytes(all_results, summary, excel_buffer)
        excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode("utf-8")

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
        _cleanup_old_sessions()

        return JSONResponse({
            "results": all_results,
            "summary": summary,
            "report_b64": excel_b64,
            "report_filename": report_mod.report_filename(),
            "chat_context": chat_context,
            "session_id": session_id,
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


CHATBOT_ENABLED = False  # temporarily disabled pending review — flip to True to re-enable


@app.post("/api/chat")
async def chat(payload: dict):
    if not CHATBOT_ENABLED:
        return JSONResponse({"error": "Chat is temporarily unavailable."}, status_code=503)
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
