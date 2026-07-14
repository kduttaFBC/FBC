"""
FalconBridge Audit Intelligence — FastAPI backend.

Endpoints:
  POST /api/audit            reconcile SOUP vs. each uploaded servicer SRF file
  POST /api/chat             chat with FalconBridge AI about the audit
"""

import io
import os
import sys
import time
import traceback
import base64
import gc

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

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # POC — lock down after deployment
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

REPORTS_DIR = "/tmp/reports"
os.makedirs(REPORTS_DIR, exist_ok=True)


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

        all_results = []
        servicer_loan_counts = {}
        servicer_csvs = {}

        for uploaded in servicer_files:
            file_bytes = await uploaded.read()
            try:
                servicer_name, loan_terms_df, servicer_inputs_df = file_readers.read_servicer_file(file_bytes)
            except ValueError as exc:
                return JSONResponse(status_code=400, content={
                    "error": f"Could not read servicer file '{uploaded.filename}': {exc}",
                })
            del file_bytes
            gc.collect()

            soup_servicer_df = file_readers.filter_soup_by_servicer(soup_df, servicer_name)

            results = checks_mod.run_servicer_reconciliation(
                servicer_name, loan_terms_df, servicer_inputs_df, soup_servicer_df,
            )
            all_results.extend(results)
            servicer_loan_counts[servicer_name] = len(servicer_inputs_df)
            servicer_csvs[servicer_name] = _truncate_csv(servicer_inputs_df, columns=SERVICER_CHAT_COLUMNS)

        summary = checks_mod.summarize(all_results, len(soup_df), servicer_loan_counts)

        excel_buffer = io.BytesIO()
        report_mod.generate_report_bytes(all_results, summary, excel_buffer)
        excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode("utf-8")

        chat_context = {
            "summary": summary,
            "results": all_results,
            "soup_csv": _truncate_csv(soup_df, columns=SOUP_CHAT_COLUMNS),
            "servicer_csvs": servicer_csvs,
        }

        return JSONResponse({
            "results": all_results,
            "summary": summary,
            "report_b64": excel_b64,
            "report_filename": report_mod.report_filename(),
            "chat_context": chat_context,
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


@app.post("/api/chat")
async def chat(payload: dict):
    try:
        message = payload.get("message", "")
        history = payload.get("history", [])
        audit_context = payload.get("audit_context", {})
        if not message.strip():
            return JSONResponse(status_code=400, content={"error": "Empty message."})
        reply = chatbot_mod.chat(message, history, audit_context)
        return {"reply": reply}
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse(
            status_code=500,
            content={"error": "The assistant is unavailable right now. Please try again.",
                     "detail": str(exc)},
        )
