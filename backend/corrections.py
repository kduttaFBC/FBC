"""
Audit Intelligence — corrected servicer workbook generation.

This module is the ONLY place that decides which audit findings are
confident enough to auto-correct, and the only place that ever writes to a
copy of an uploaded servicer workbook. It is deliberately isolated from the
audit/reconciliation logic in checks.py:

  - checks.py (status_match_check / STAT-01) computes raw facts about each
    loan — what the servicer reported, what the engine calculated, and which
    warning conditions apply — and exposes them on result["correction_data"].
    It makes NO decision about what's "correctable."
  - This module reads those raw facts and applies a strict, mechanical
    allow-list (build_eligible_corrections) to decide what may be written.
  - Nothing in this file ever changes an audit result, a report, or a
    reconciliation formula. A bug here cannot silently change what the audit
    engine concludes; a bug in the audit engine cannot silently widen what
    this module is willing to write.

Version 1 scope (deliberately narrow):
  - Delinquency Status: corrected only when STAT-01 produced a real,
    unambiguous calculated value and flagged no warning of any kind for that
    loan.
  - Loan Status: corrected only to DEFERRED, only for a CONFIRMED Full P&I
    deferral in its "start" or "mid" phase. Partial (principal-only/
    interest-only) deferrals and the deferral "resume" phase are explicitly
    OUT of scope for v1 — see build_eligible_corrections for why.
  - Every other field (balances, payments, population, liquidation,
    negative payments, missing periods, duplicate postings, everything
    else) is untouched in this version.
  - The original uploaded workbook is never modified — a new workbook is
    always produced from a fresh copy of the original bytes.
"""

import io
from datetime import datetime, timezone

from openpyxl import load_workbook

from report import _write_header, _write_rows, _autofit


AUDIT_CHECK_ID = "STAT-01"

# Any of these present for a loan excludes BOTH correction types for that
# loan, even if only one field's mismatch is otherwise eligible — a row with
# any flagged uncertainty about itself isn't trusted for a different field
# either.
WARNING_REASONS = {
    "MEMO DATA QUALITY WARNING",
    "DEFERRAL DATA QUALITY WARNING",
    "SCHEDULED PAYMENT AMBIGUOUS",
}

_NOT_A_REAL_STATUS = {"", "UNVERIFIED", "REVIEW", "NAN", "NONE"}

# --- Loan Status safety rules (client-specified, independent of checks.py) --
#
# These are checked here, in corrections.py, regardless of what checks.py's
# own terminal-status classification concluded. This is deliberate defense in
# depth: rules 2 and 3 below would otherwise depend entirely on checks.py's
# internal terminal-status keyword list staying correct and complete forever
# — a dependency this module should not have to trust. If either list ever
# drifts out of sync, THIS guard is what actually protects LIQUIDATION and
# PAID-IN-FULL loans, not an assumption about the other module's behavior.
#
#   1. Never create LIQUIDATION automatically.
#      Satisfied by construction: corrected_value for Loan Status is always
#      the literal "DEFERRED" below — there is no code path that can ever
#      write "LIQUIDATION".
#   2. Never remove LIQUIDATION automatically.
#   3. Never change any PAID-IN-FULL loan to another status.
#      Checked against BOTH the servicer's current-cycle reported status AND
#      SOUP's latest recorded status — either one alone being terminal is
#      enough to block the correction.
#   4. Never change any loan into PAID-IN-FULL automatically.
#      Also satisfied by construction (corrected_value is always "DEFERRED").
#   5. Loan Status corrections are limited to "anything -> DEFERRED", only
#      for a CONFIRMED Full P&I deferral, phase "start" or "mid".
#   6. Every other Loan Status mismatch is left for manual review — no other
#      Loan Status correction branch exists in this module.
_LOAN_STATUS_TERMINAL_KEYWORDS = ("LIQUIDATION", "PAID-IN-FULL", "PIF")


def _is_terminal_loan_status(status_value):
    s = str(status_value or "").strip().upper()
    return any(k in s for k in _LOAN_STATUS_TERMINAL_KEYWORDS)

CORRECTION_LOG_SHEET_NAME = "Correction Log"
CORRECTION_LOG_HEADERS = [
    "Institution", "Source Worksheet", "Excel Row Number", "SBA GP Number",
    "Field Updated", "Original Value", "Corrected Value", "Audit Check ID",
    "Eligibility Rule", "Timestamp",
]


def _norm(s):
    return str(s).strip().lower().replace(" ", "").replace("_", "").replace("-", "")


def _is_real_dq_value(value):
    if value is None:
        return False
    return _norm(value) not in {_norm(v) for v in _NOT_A_REAL_STATUS}


def _normalize_gp_scalar(value):
    """Mirrors file_readers.normalize_id_series's per-value logic, for
    comparing a single cell read back out of the reopened workbook against
    the already-normalized gp key used throughout checks.py. Kept as a
    small local copy rather than importing a private file_readers helper,
    to keep this module's dependencies minimal and explicit."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    try:
        f = float(s)
        return str(int(f))
    except (TypeError, ValueError):
        pass
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits or None


def build_eligible_corrections(correction_data):
    """Applies the v1 allow-list to STAT-01's raw per-loan facts. Returns a
    list of correction dicts, each describing ONE cell to change:
    {gp, excel_row, sheet_name, header_excel_row, field, corrected_value,
    eligibility_rule}.

    This is the single place v1's correction policy lives. Widening scope
    later (e.g. partial deferrals, the resume phase) means changing the
    conditions below — it never requires touching checks.py."""
    eligible = []

    for c in correction_data or []:
        reasons = set(c.get("reasons") or [])
        if reasons & WARNING_REASONS:
            continue  # any data-quality warning excludes this loan entirely

        if c.get("excel_row") is None or c.get("sheet_name") is None or c.get("header_excel_row") is None:
            continue  # can't safely locate this row in the original file — skip, never guess

        # CLIENT-CONFIRMED (2026-08-03): a Delinquency Status driven by a
        # negative Payment Received amount (a reversal) requires manual
        # review before correction, even though the calculation itself is
        # deterministic and STAT-01/NEG-01 agree on it. This is controlled
        # HERE, at the eligibility layer — not by giving STAT-01 a different
        # (deliberately vaguer) calculated_dq than NEG-01 shows, which would
        # violate the "the two checks must never diverge" principle this
        # whole shared-calculation design exists to protect.
        if c.get("reversal_equivalents") is not None:
            continue

        gp = c.get("gp")

        # --- Delinquency Status -------------------------------------------------
        if (
            "DELINQUENCY STATUS MISMATCH" in reasons
            and c.get("dq_match") is False  # never True/"REVIEW" — REVIEW means UNVERIFIED, already excluded upstream
            and _is_real_dq_value(c.get("calculated_dq"))
        ):
            eligible.append({
                "gp": gp,
                "excel_row": c["excel_row"],
                "sheet_name": c["sheet_name"],
                "header_excel_row": c["header_excel_row"],
                "field": "Delinquency Status",
                "corrected_value": str(c["calculated_dq"]).strip().upper(),
                "eligibility_rule": "STAT-01 deterministic calculation",
            })

        # --- Loan Status -> DEFERRED (confirmed Full P&I only, v1) --------------
        # expected_loan_status_raw == "DEFERRED" only ever gets set by a
        # confirmed deferral override in checks.py — never a coincidental
        # string match against SOUP (that fallback uses the DIFFERENT,
        # "or soup_status" display value, which this deliberately does not
        # read). Restricting to phase in (start, mid) means a correction is
        # only ever generated INTO Deferred, never out of it — v1 never
        # auto-exits a loan from Deferred during a resume cycle, by
        # construction, not by an extra rule that could be forgotten later.
        #
        # The two _is_terminal_loan_status checks below are the independent
        # safety net for rules 2 and 3 (never remove LIQUIDATION, never
        # change PAID-IN-FULL) — checked against BOTH the servicer's current
        # report and SOUP's latest recorded status, regardless of whether
        # checks.py's own terminal-status classification already should have
        # prevented deferral from being confirmed here in the first place.
        if (
            "LOAN STATUS MISMATCH" in reasons
            and c.get("expected_loan_status_raw") == "DEFERRED"
            and c.get("deferral_type") == "full"
            and c.get("deferral_phase") in ("start", "mid")
            and not _is_terminal_loan_status(c.get("srf_status"))
            and not _is_terminal_loan_status(c.get("soup_status"))
        ):
            eligible.append({
                "gp": gp,
                "excel_row": c["excel_row"],
                "sheet_name": c["sheet_name"],
                "header_excel_row": c["header_excel_row"],
                "field": "Loan Status",
                "corrected_value": "DEFERRED",
                "eligibility_rule": f"STAT-01 confirmed Full P&I deferral ({c.get('deferral_phase')})",
            })

    return eligible


def _find_header_columns(ws, header_excel_row, field, gp_field="SBA GP Number"):
    """Scans the header row for the exact expected column names. Returns
    (field_col_idx, gp_col_idx), either of which is None if not found at
    that exact row — callers must treat None as "skip this correction,"
    never guess at a nearby column. This is the safety check that protects
    against a future template change silently writing to the wrong column."""
    field_col = None
    gp_col = None
    max_col = ws.max_column or 0
    for col in range(1, max_col + 1):
        value = ws.cell(row=header_excel_row, column=col).value
        if value is None:
            continue
        n = _norm(value)
        if n == _norm(field):
            field_col = col
        if n == _norm(gp_field):
            gp_col = col
    return field_col, gp_col


def _apply_one_correction(wb, corr):
    """Attempts to write one correction. Returns (True, log_row_dict) on
    success, or (False, skip_reason_str) if any safety check fails — a
    failed check always skips just this one correction, never raises and
    never guesses at an alternate row/column."""
    sheet_name = corr["sheet_name"]
    if sheet_name not in wb.sheetnames:
        return False, f"Sheet '{sheet_name}' not found in the original workbook."
    ws = wb[sheet_name]

    header_excel_row = corr["header_excel_row"]
    field_col, gp_col = _find_header_columns(ws, header_excel_row, corr["field"])
    if field_col is None:
        return False, (
            f"Could not locate the '{corr['field']}' column at the expected header row "
            f"in '{sheet_name}' — skipped for safety (template may have changed)."
        )
    if gp_col is None:
        return False, (
            f"Could not locate the 'SBA GP Number' column at the expected header row "
            f"in '{sheet_name}' — skipped for safety."
        )

    excel_row = corr["excel_row"]
    actual_gp = _normalize_gp_scalar(ws.cell(row=excel_row, column=gp_col).value)
    if actual_gp != corr["gp"]:
        return False, (
            f"Row {excel_row} in '{sheet_name}' no longer matches loan {corr['gp']} "
            f"(found {actual_gp!r}) — skipped for safety."
        )

    # Final write-time safety net for Loan Status specifically: this module
    # must never write LIQUIDATION or PAID-IN-FULL/PIF automatically (client
    # safety rules 1 and 4), no matter what a future change to the allow-list
    # above might otherwise permit. This is intentionally redundant with
    # build_eligible_corrections only ever producing "DEFERRED" today.
    if corr["field"] == "Loan Status" and _is_terminal_loan_status(corr["corrected_value"]):
        return False, (
            f"Refused to write Loan Status '{corr['corrected_value']}' for loan {corr['gp']} — "
            f"automatic corrections may never create or remove a LIQUIDATION or PAID-IN-FULL status."
        )

    cell = ws.cell(row=excel_row, column=field_col)
    original_value = cell.value
    cell.value = corr["corrected_value"]

    return True, {
        "sheet_name": sheet_name,
        "excel_row": excel_row,
        "gp": corr["gp"],
        "field": corr["field"],
        "original_value": original_value,
        "corrected_value": corr["corrected_value"],
        "eligibility_rule": corr["eligibility_rule"],
    }


def _unique_log_sheet_name(wb):
    base = CORRECTION_LOG_SHEET_NAME
    if base not in wb.sheetnames:
        return base
    number = 2
    while f"{base} {number}" in wb.sheetnames:
        number += 1
    return f"{base} {number}"


def _add_correction_log_sheet(wb, applied_rows, institution, timestamp):
    ws = wb.create_sheet(_unique_log_sheet_name(wb))

    rows = [
        [
            institution,
            r["sheet_name"],
            r["excel_row"],
            r["gp"],
            r["field"],
            r["original_value"],
            r["corrected_value"],
            AUDIT_CHECK_ID,
            r["eligibility_rule"],
            timestamp,
        ]
        for r in applied_rows
    ]
    _write_header(ws, CORRECTION_LOG_HEADERS)
    _write_rows(ws, CORRECTION_LOG_HEADERS, rows)
    _autofit(ws, CORRECTION_LOG_HEADERS, rows)


def _filename_month_label(cycle_month):
    if cycle_month is None:
        return None
    try:
        return cycle_month.strftime("%B_%Y")  # e.g. "June_2026"
    except (AttributeError, ValueError):
        return None


def _build_corrected_filename(institution, cycle_month):
    safe_institution = "_".join(str(institution).split())
    month_label = _filename_month_label(cycle_month) or datetime.now(timezone.utc).strftime("%Y-%m-%d")
    return f"{safe_institution}_{month_label}_CORRECTED.xlsx"


def generate_corrected_workbook(file_bytes, original_filename, institution, stat01_result):
    """Given one servicer file's ORIGINAL raw bytes (never mutated) and its
    STAT-01 result (which carries result["correction_data"] and
    result["_cycle_month"]), returns a dict:

      {"file_bytes": <corrected .xlsx bytes>, "filename": <str>,
       "correction_count": <int>, "skip_reason": None}

    or, when no corrected workbook should be produced:

      {"file_bytes": None, "filename": None, "correction_count": 0,
       "skip_reason": <str explaining why>}

    Never raises for expected conditions (missing columns, mismatched rows,
    unsupported format) — those are all reported via skip_reason so a
    problem with ONE servicer's file never breaks the rest of the audit
    response."""
    empty = {"file_bytes": None, "filename": None, "correction_count": 0, "skip_reason": None}

    correction_data = (stat01_result or {}).get("correction_data", [])
    eligible = build_eligible_corrections(correction_data)
    if not eligible:
        return empty  # nothing eligible — no misleading file, and no error; this is the normal/common case

    if original_filename and original_filename.lower().strip().endswith(".xls"):
        return {
            **empty,
            "skip_reason": (
                "Eligible corrections were identified, but a corrected copy "
                "cannot be generated safely from the legacy .xls format."
            ),
        }

    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
    except Exception as exc:
        return {**empty, "skip_reason": f"Could not reopen the original workbook: {exc}"}

    applied_rows = []
    for corr in eligible:
        ok, detail = _apply_one_correction(wb, corr)
        if ok:
            applied_rows.append(detail)
        # skipped corrections are silently omitted from the output (not from
        # this function's return value's skip_reason, which describes the
        # workbook as a whole) — a template-mismatch skip on one loan
        # shouldn't block the rest of that same file's valid corrections.

    if not applied_rows:
        return {**empty, "skip_reason": "No corrections could be safely verified against this file's actual layout."}

    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    _add_correction_log_sheet(wb, applied_rows, institution, timestamp)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = _build_corrected_filename(institution, (stat01_result or {}).get("_cycle_month"))
    return {
        "file_bytes": buf.getvalue(),
        "filename": filename,
        "correction_count": len(applied_rows),
        "skip_reason": None,
    }
