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

# Delinquency Status values a servicer may legitimately report. A value
# outside this set is a SOURCE DATA DOMAIN DEFECT, not merely a mismatch:
# it is dropped from every portfolio bucket while remaining in Outstanding,
# so the summary stops tying to itself.
#
# Observed live: Peoples Bank July 2026 loan 4017399101 reported
# 'FORCLOSURE' — misspelled, AND a Loan Status value sitting in a
# Delinquency Status field. STAT-01 calculated 30 DAYS for it, which is
# arithmetically right, and auto-applying that would have quietly normalised
# the defect away: the servicer never learns their template is being filled
# in wrong, and it recurs next cycle. 584,023.61 of balance was unbucketed
# by this one value.
_PERMITTED_DQ_STATUS = {
    "CURRENT", "DELINQUENT",
    "30 DAYS", "60 DAYS", "90 DAYS", "120 DAYS",
    "150 DAYS", "150+ DAYS", "180 DAYS", "210 DAYS",
}

REVIEW_REASON_CROSS_DOMAIN = (
    "Reported Delinquency Status contains an invalid cross-domain value."
)


def _is_cross_domain_dq(value):
    """True when a reported Delinquency Status is outside the permitted
    domain — a misspelling, or a value belonging to another field."""
    if value is None:
        return False
    v = str(value).strip().upper()
    if not v or v in _NOT_A_REAL_STATUS:
        return False
    return v not in _PERMITTED_DQ_STATUS

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
# CLIENT-CONFIRMED (2026-08-05): one log row per LOAN-LEVEL correction, not
# one per Excel row updated -- a 3-posting loan whose Delinquency Status was
# corrected produces ONE Correction Log row, with "Rows Updated" = 3 and
# "Excel Row Numbers" listing all three, rather than three separate rows
# that would misleadingly read as three distinct corrections.
# CLIENT-CONFIRMED (2026-08-05): "Skipped Row References" records any row
# reference that didn't get updated as part of this same correction —
# invalid/duplicate references filtered out before writing, or a row whose
# GP no longer matched at write time — so a partially-applied correction
# never silently loses track of what it left out.
CORRECTION_LOG_HEADERS = [
    "Institution", "Source Worksheet", "Excel Row Numbers", "Rows Updated", "SBA GP Number",
    "Field Updated", "Original Value", "Corrected Value", "Audit Check ID",
    "Eligibility Rule", "Timestamp", "Skipped Row References",
    "Review Status", "Review Reason",
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


def _coerce_excel_row(value):
    """Returns a positive int Excel row number, or None if the value is
    blank, null, nonnumeric, or otherwise not a valid row reference. Never
    guesses — a value like 12.5 (not a whole number) or a negative/zero row
    is rejected outright rather than rounded or substituted with a nearby
    row."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value if value > 0 else None
    if isinstance(value, float):
        if value != value:  # NaN
            return None
        return int(value) if value.is_integer() and value > 0 else None
    s = str(value).strip()
    if not s:
        return None
    try:
        f = float(s)
    except (TypeError, ValueError):
        return None
    if f != f or not f.is_integer() or f <= 0:
        return None
    return int(f)


def _clean_excel_rows(raw_rows):
    """CLIENT-CONFIRMED (2026-08-05): validates and deduplicates a loan's
    posting-row references BEFORE any correction is applied. Returns
    (clean_rows, skipped) where clean_rows preserves the original
    chronological order with invalid entries removed and duplicates
    collapsed to their first occurrence, and skipped is a list of
    {"row_ref": <original value>, "reason": <str>} for every entry that
    didn't make it into clean_rows — surfaced later in the Correction Log
    / returned metadata rather than silently dropped."""
    clean = []
    seen = set()
    skipped = []
    for raw in raw_rows or []:
        row = _coerce_excel_row(raw)
        if row is None:
            skipped.append({
                "row_ref": raw,
                "reason": "blank, null, non-numeric, or invalid row reference — skipped without guessing",
            })
            continue
        if row in seen:
            skipped.append({
                "row_ref": row,
                "reason": "duplicate row reference — already scheduled for this correction",
            })
            continue
        seen.add(row)
        clean.append(row)
    return clean, skipped


def build_eligible_corrections(correction_data):
    """Applies the v1 allow-list to STAT-01's raw per-loan facts. Returns a
    list of correction dicts, each describing ONE loan-level field change to
    apply across every one of that loan's posting rows this cycle:
    {gp, excel_rows, skipped_row_refs, sheet_name, header_excel_row, field,
    corrected_value, eligibility_rule}. excel_rows (CLIENT-CONFIRMED
    2026-08-05) is every posting row for the loan, not just the last — see
    checks.consolidate_cycle_postings' _all_excel_rows — so a corrected
    Loan Status or Delinquency Status is written consistently to every
    posting row for that loan/cycle, never leaving earlier postings with
    the original, now-superseded value. excel_rows here is already
    validated and deduplicated (CLIENT-CONFIRMED 2026-08-05) via
    _clean_excel_rows — skipped_row_refs carries what was removed and why.

    A loan/field that matches every business rule below is ALWAYS included
    here — even if excel_rows ends up empty after validation (e.g. every
    reference was invalid, or no location metadata existed at all).
    _apply_one_correction is the single place that decides whether zero
    valid rows means "not applied" (CLIENT-CONFIRMED 2026-08-05) — kept
    there, not here, specifically so this function's return type stays the
    single list it has always been. A loan that never matched any
    eligibility rule in the first place (no warning-free mismatch, wrong
    deferral phase, etc.) never appears here at all — there was never a
    correction to report on.

    This is the single place v1's correction policy lives. Widening scope
    later (e.g. partial deferrals, the resume phase) means changing the
    conditions below — it never requires touching checks.py."""
    eligible = []

    for c in correction_data or []:
        reasons = set(c.get("reasons") or [])
        if reasons & WARNING_REASONS:
            continue  # any data-quality warning excludes this loan entirely

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
        excel_rows, skipped_row_refs = _clean_excel_rows(c.get("excel_rows"))

        def _emit(field, corrected_value, eligibility_rule,
                  review_required=False, review_reason=None):
            eligible.append({
                "gp": gp,
                "excel_rows": excel_rows,
                "skipped_row_refs": skipped_row_refs,
                "sheet_name": c.get("sheet_name"),
                "header_excel_row": c.get("header_excel_row"),
                "field": field,
                "corrected_value": corrected_value,
                "eligibility_rule": eligibility_rule,
                # A correction flagged for review is PROPOSED, never written.
                # It is still recorded in the Correction Log so the defect is
                # visible and can be raised with the servicer, rather than
                # disappearing the way an eligibility-layer `continue` would
                # make it disappear.
                "review_required": bool(review_required),
                "review_reason": review_reason,
                # The value actually reported by the servicer, carried so a
                # proposed-only row can show what needs reviewing without
                # having opened the workbook to read the cells.
                "reported_value": c.get("srf_dq") if field == "Delinquency Status"
                                  else c.get("srf_status"),
            })

        # --- Delinquency Status -------------------------------------------------
        if (
            "DELINQUENCY STATUS MISMATCH" in reasons
            and c.get("dq_match") is False  # never True/"REVIEW" — REVIEW means UNVERIFIED, already excluded upstream
            and _is_real_dq_value(c.get("calculated_dq"))
        ):
            _cross_domain = _is_cross_domain_dq(c.get("srf_dq"))
            _emit(
                "Delinquency Status",
                str(c["calculated_dq"]).strip().upper(),
                "STAT-01 deterministic calculation",
                review_required=_cross_domain,
                review_reason=REVIEW_REASON_CROSS_DOMAIN if _cross_domain else None,
            )

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
            _emit("Loan Status", "DEFERRED", f"STAT-01 confirmed Full P&I deferral ({c.get('deferral_phase')})")

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
    """Attempts to write one LOAN-LEVEL correction across every posting row
    for that loan/cycle (corr["excel_rows"], already validated/deduplicated
    by _clean_excel_rows). Returns (True, log_row_dict) on success —
    log_row_dict describes the correction as a whole (all rows updated, all
    their original values, in row order, plus any rows that were skipped
    and why) — or (False, skip_reason_str) if the correction as a whole
    can't be safely applied at all: zero valid row references to begin
    with (CLIENT-CONFIRMED 2026-08-05 — e.g. no location metadata existed,
    or every reference was invalid/duplicate), a missing sheet/column, or
    every row failing its per-row GP check. This function is the single
    place that decides "zero valid rows" means not applied — build_eligible_
    corrections always includes a matched loan/field here regardless of
    whether it has any valid excel_rows, specifically so that decision
    lives in exactly one place and that function's return stays a plain
    list.

    Each row is still validated independently before being written — same
    "never guess" posture as before, just applied per-row instead of once:
    if one posting row's GP no longer matches (e.g. the template shifted),
    that ROW alone is skipped and the rest of the loan's rows still get the
    correction, since the safety concern is specific to that one row's
    location, not to the loan or the correction itself. Only if EVERY row
    fails does this correction get reported as fully skipped (never counted
    as successful). Every skipped reference — whether filtered out earlier
    for being invalid/duplicate, missing entirely, or skipped here for a GP
    mismatch — is recorded so nothing disappears silently. Payment amounts,
    dates, balances, formulas, and formatting are never touched — only the
    single target field column, per row."""
    skipped_row_refs = corr.get("skipped_row_refs") or []
    if not corr.get("excel_rows"):
        detail = (
            "; ".join(f"{s['row_ref']} ({s['reason']})" for s in skipped_row_refs)
            if skipped_row_refs else "no row location metadata available"
        )
        return False, (
            f"No valid posting row reference for loan {corr['gp']} ({detail}) — skipped for safety."
        )

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

    # Final write-time safety net for Loan Status specifically: this module
    # must never write LIQUIDATION or PAID-IN-FULL/PIF automatically (client
    # safety rules 1 and 4), no matter what a future change to the allow-list
    # above might otherwise permit. This is intentionally redundant with
    # build_eligible_corrections only ever producing "DEFERRED" today. Checked
    # once for the whole correction, since corrected_value is the same value
    # for every row.
    if corr["field"] == "Loan Status" and _is_terminal_loan_status(corr["corrected_value"]):
        return False, (
            f"Refused to write Loan Status '{corr['corrected_value']}' for loan {corr['gp']} — "
            f"automatic corrections may never create or remove a LIQUIDATION or PAID-IN-FULL status."
        )

    updated_rows = []
    original_values = []
    skipped_rows = list(corr.get("skipped_row_refs") or [])  # invalid/duplicate refs already filtered out upstream
    for excel_row in corr["excel_rows"]:
        actual_gp = _normalize_gp_scalar(ws.cell(row=excel_row, column=gp_col).value)
        if actual_gp != corr["gp"]:
            skipped_rows.append({
                "row_ref": excel_row,
                "reason": f"GP mismatch — expected {corr['gp']!r}, found {actual_gp!r}",
            })
            continue
        cell = ws.cell(row=excel_row, column=field_col)
        original_values.append(cell.value)
        cell.value = corr["corrected_value"]
        updated_rows.append(excel_row)

    if not updated_rows:
        detail = (
            "; ".join(f"{s['row_ref']} ({s['reason']})" for s in skipped_rows)
            if skipped_rows else "no posting rows were available"
        )
        return False, (
            f"No posting row for loan {corr['gp']} in '{sheet_name}' could be safely verified "
            f"({detail}) — skipped for safety."
        )

    return True, {
        "review_status": "Applied",
        "review_reason": None,
        "sheet_name": sheet_name,
        "excel_rows": updated_rows,
        "rows_updated": len(updated_rows),
        "gp": corr["gp"],
        "field": corr["field"],
        "original_values": original_values,
        "corrected_value": corr["corrected_value"],
        "eligibility_rule": corr["eligibility_rule"],
        "skipped_rows": skipped_rows,
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
            ", ".join(str(x) for x in r["excel_rows"]),
            r["rows_updated"],
            r["gp"],
            r["field"],
            "; ".join("" if v is None else str(v) for v in r["original_values"]),
            r["corrected_value"],
            AUDIT_CHECK_ID,
            r["eligibility_rule"],
            timestamp,
            "; ".join(f"{s['row_ref']} ({s['reason']})" for s in r.get("skipped_rows") or []),
            r.get("review_status") or "Applied",
            r.get("review_reason") or "",
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
    response.

    CLIENT-CONFIRMED (2026-08-05): the returned dict also carries
    "corrections_not_applied" — a list of {gp, field, reason} for every
    eligible correction that could NOT be safely written to even one row
    (zero valid excel_rows to begin with, or every row failing its GP
    check). Those corrections are never counted in correction_count and
    never appear in the Correction Log sheet — they're reported as not
    applied, never as successful — but the reason is still surfaced here
    rather than disappearing silently."""
    empty = {
        "file_bytes": None, "filename": None, "correction_count": 0,
        "skip_reason": None, "corrections_not_applied": [],
    }

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
    corrections_not_applied = []
    for corr in eligible:
        # CROSS-DOMAIN SOURCE DEFECT -> proposed, never written.
        # The calculation is deterministic and correct, but applying it
        # would normalise away an invalid value in the servicer's own file:
        # they would never learn the field is being populated from the wrong
        # domain, and it would recur every cycle. The proposal is recorded in
        # the Correction Log marked "Human review required", and reported to
        # the caller as not applied, so it is visible in both places without
        # being silently actioned.
        if corr.get("review_required"):
            applied_rows.append({
                "sheet_name": corr.get("sheet_name"),
                "excel_rows": corr.get("excel_rows") or [],
                "rows_updated": 0,
                "gp": corr["gp"],
                "field": corr["field"],
                "original_values": ([corr.get("reported_value")]
                                    if corr.get("reported_value") is not None else []),
                "corrected_value": corr["corrected_value"],
                "eligibility_rule": corr["eligibility_rule"],
                "skipped_rows": [],
                "review_status": "Human review required",
                "review_reason": corr.get("review_reason"),
            })
            corrections_not_applied.append({
                "gp": corr["gp"], "field": corr["field"],
                "reason": (
                    f"Human review required — {corr.get('review_reason')} "
                    f"Reported value: {corr.get('reported_value')!r}. "
                    f"The calculated value ({corr['corrected_value']}) is "
                    f"recorded in the Correction Log as proposed but was not "
                    f"written; the servicer should be informed so the error "
                    f"does not recur."
                ),
            })
            continue

        ok, detail = _apply_one_correction(wb, corr)
        if ok:
            applied_rows.append(detail)
        else:
            # A correction with zero valid rows (or any other full-loan
            # safety failure) is reported here as not applied — it never
            # contributes to correction_count and never gets a Correction
            # Log row, but the reason isn't lost: it's returned so the
            # caller can surface it if useful, rather than one servicer's
            # template mismatch silently erasing all trace of the attempt.
            corrections_not_applied.append({"gp": corr["gp"], "field": corr["field"], "reason": detail})

    really_applied = [r for r in applied_rows if r.get("rows_updated")]
    if not really_applied:
        return {
            **empty,
            "skip_reason": "No corrections could be safely verified against this file's actual layout.",
            "corrections_not_applied": corrections_not_applied,
        }

    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    _add_correction_log_sheet(wb, applied_rows, institution, timestamp)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = _build_corrected_filename(institution, (stat01_result or {}).get("_cycle_month"))
    return {
        "file_bytes": buf.getvalue(),
        "filename": filename,
        "correction_count": len(really_applied),
        "skip_reason": None,
        "corrections_not_applied": corrections_not_applied,
    }
