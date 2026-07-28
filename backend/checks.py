"""
FalconBridge Audit Intelligence — SBA loan reconciliation checks.

Replaces the old Customer Master / Bank Transactions checks (CM-01..05,
TX-01..09) with checks that reconcile a servicer's SRF submission (Servicer
Inputs + Loan Terms & Prior Period) against the SOUP Performance Dataset for
that same servicer.

Each check returns a dict:
{
    "id": "POP-01",
    "name": "Population Check",
    "category": "SBA Reconciliation",
    "severity": "FAIL" | "WARN",
    "status": "PASS" | "WARN" | "FAIL",
    "issues_found": <int>,
    "details": "<human readable summary>",
    "records": [ { ...flagged record... }, ... ],
}
"""

import pandas as pd
from file_readers import normalize_id_series as _norm_id
from file_readers import MEMO_LENDER_ALIASES


BALANCE_TOLERANCE = 1.00  # dollars


def _result(check_id, name, category, severity, records, details_fn):
    issues = len(records)
    status = "PASS" if issues == 0 else severity
    details = "No issues found." if issues == 0 else details_fn(issues)
    return {
        "id": check_id,
        "name": name,
        "category": category,
        "severity": severity,
        "status": status,
        "issues_found": issues,
        "details": details,
        "records": records,
    }


# ---------------------------------------------------------------------------
# Shared cycle-level posting consolidation, per client-confirmed rule:
# when a loan has multiple postings in one reporting cycle, they represent
# multiple transactions against ONE monthly obligation, not multiple
# months. BAL-01, STAT-01, PAY-02, and NEG-01 all previously consolidated
# postings independently and inconsistently (see rule-gap review) — this is
# the single shared implementation all four now use, so they can't diverge
# on the same GP/cycle the way BAL-01 and STAT-01 (and STAT-01 and NEG-01)
# previously could.
#
# Chronological order (needed before "first"/"last" mean anything) is
# established by, in order of preference:
#   1. Loan Payment Instance, if populated (an explicit 1, 2, 3... ordinal —
#      more reliable than date, since same-day postings can't be
#      distinguished by date alone)
#   2. Payment Receipt Date
#   3. original row order (stable sort preserves this as the final fallback)
# ---------------------------------------------------------------------------

def consolidate_cycle_postings(servicer_inputs_df):
    """Returns one consolidated cycle record per SBA GP Number:
      posting_count, loan_status, dq_status (both from the LAST chronological
      posting), received/principal/net_interest (SUMMED), scheduled (a
      single value — see below), scheduled_ambiguous (bool),
      scheduled_distinct_values (list, diagnostic), beginning_balance
      (FIRST posting), ending_balance (LAST posting), interest_from (FIRST),
      interest_to (LAST), lender_loan_number (FIRST).

    Does NOT include the four liquidation-specific fields (Liquidation Date,
    Liquidation Proceeds, Subsequent Recoveries/(Losses), Liquidation Loss)
    — LIQ-01 (liquidation_review_check) computes those itself directly from
    the raw postings via _consolidate_liquidation_fields, since a naive
    first/last choice here would silently hide genuine multi-posting
    disagreement rather than surface it for review.

    Scheduled Payment is never summed — there is one scheduled obligation
    per cycle. If every nonzero Scheduled Payment value for a loan this
    cycle agrees, that value is used. If two or more DISTINCT nonzero
    values exist for the same loan in the same cycle, `scheduled` is None
    and `scheduled_ambiguous` is True — callers must not silently sum, max,
    or pick one; the cycle obligation is genuinely ambiguous and should be
    surfaced, not guessed at.
    """
    si = servicer_inputs_df.copy()
    si["SBA GP Number"] = _norm_id(si["SBA GP Number"])
    si = si.reset_index(drop=True)  # preserve original row order as the final sort fallback

    numeric_cols = [c for c in ("Payment Received", "Principal", "Net Interest", "Scheduled Payment",
                                 "Beginning Balance", "Ending Balance") if c in si.columns]
    for c in numeric_cols:
        si[c] = pd.to_numeric(si[c], errors="coerce")

    rows = []
    for gp, group in si.groupby("SBA GP Number", sort=False):
        if "Loan Payment Instance" in group.columns and pd.to_numeric(group["Loan Payment Instance"], errors="coerce").notna().any():
            group = group.assign(_sort_key=pd.to_numeric(group["Loan Payment Instance"], errors="coerce")).sort_values("_sort_key", kind="stable")
        elif "Payment Receipt Date" in group.columns:
            group = group.assign(_sort_key=pd.to_datetime(group["Payment Receipt Date"], errors="coerce")).sort_values("_sort_key", kind="stable")
        # else: already in original row order (stable groupby preserves it)

        first = group.iloc[0]
        last = group.iloc[-1]

        sched_series = group["Scheduled Payment"] if "Scheduled Payment" in group.columns else pd.Series(dtype=float)
        nonzero = sched_series[(sched_series.notna()) & (sched_series != 0)]
        distinct_nonzero = sorted(nonzero.unique().tolist())
        scheduled_ambiguous = len(distinct_nonzero) > 1
        if scheduled_ambiguous:
            scheduled = None
        elif distinct_nonzero:
            scheduled = distinct_nonzero[0]
        elif sched_series.notna().any():
            scheduled = 0.0
        else:
            scheduled = None

        rows.append({
            "SBA GP Number": gp,
            "posting_count": len(group),
            "loan_status": last.get("Loan Status", ""),
            "dq_status": last.get("Delinquency Status", ""),
            "received": group["Payment Received"].sum(min_count=1) if "Payment Received" in group.columns else None,
            "principal": group["Principal"].sum(min_count=1) if "Principal" in group.columns else None,
            "net_interest": group["Net Interest"].sum(min_count=1) if "Net Interest" in group.columns else None,
            "scheduled": scheduled,
            "scheduled_ambiguous": scheduled_ambiguous,
            "scheduled_distinct_values": distinct_nonzero,
            "beginning_balance": first.get("Beginning Balance"),
            "ending_balance": last.get("Ending Balance"),
            "interest_from": first.get("Interest-From", ""),
            "interest_to": last.get("Interest-To", ""),
            "lender_loan_number": first.get("Lender Loan Number", ""),
        })

    return pd.DataFrame(rows).set_index("SBA GP Number") if rows else pd.DataFrame(
        columns=["posting_count", "loan_status", "dq_status", "received", "principal", "net_interest",
                 "scheduled", "scheduled_ambiguous", "scheduled_distinct_values", "beginning_balance",
                 "ending_balance", "interest_from", "interest_to", "lender_loan_number"]
    )


# ---------------------------------------------------------------------------
# POP-01 — Population Check
# Every loan in the servicer's Loan Terms & Prior Period population should
# exist in the SOUP dataset's history for that servicer, and vice versa.
# ---------------------------------------------------------------------------

def population_check(loan_terms_df, soup_servicer_df, servicer_name, deal=None):
    records = []

    has_lt_id = "SBA GP Number" in loan_terms_df.columns
    has_soup_id = "SBA GP Number" in soup_servicer_df.columns

    # Composite population key: Servicer + Deal + normalized GP Number, not
    # Servicer + GP Number alone. SOUP pools every note-series deal a
    # servicer has under one Servicer name (it has no per-deal separation of
    # its own besides the Deal column), so for any servicer split across
    # multiple SRF files -- Millennium ("2025 FBC-1" / "2024 FBC-1"),
    # Lendistry likewise -- comparing by Servicer alone silently compares
    # each file against the OTHER note series' loans too, producing large
    # false "only in SOUP" exceptions (validated: 86 false exceptions for
    # Millennium, 1,503 for Lendistry, both fully explained by this and
    # resolved by scoping to Deal). Substitution loans that legitimately
    # appear under two deals (a loan bought out of one securitized pool and
    # resold into another) are NOT special-cased — each Servicer + Deal + GP
    # combination is compared independently, so a substitution loan simply
    # matches correctly in both deals' comparisons on its own.
    #
    # `deal` is this SRF file's own note-series identity (from
    # file_readers.extract_deal_from_note_series); when it can't be
    # determined, this falls back to the prior servicer-only comparison
    # rather than silently comparing against an empty or wrong population.
    soup_scoped = soup_servicer_df
    deal_scoped = False
    if deal and has_soup_id and "Deal" in soup_servicer_df.columns:
        soup_scoped = soup_servicer_df[soup_servicer_df["Deal"].astype("string").str.strip() == deal]
        deal_scoped = True

    lt_ids = set(_norm_id(loan_terms_df["SBA GP Number"]).dropna()) if has_lt_id else set()
    soup_ids = set(_norm_id(soup_scoped["SBA GP Number"]).dropna()) if has_soup_id else set()

    only_srf = lt_ids - soup_ids
    only_soup = soup_ids - lt_ids

    deal_label = deal or "unknown"
    for gp in sorted(only_srf):
        records.append({
            "SBA GP Number": gp,
            "Deal": deal_label,
            "Issue": "In servicer's Loan Terms & Prior Period but not found in SOUP dataset for this Deal/note series",
        })
    for gp in sorted(only_soup):
        records.append({
            "SBA GP Number": gp,
            "Deal": deal_label,
            "Issue": "In SOUP dataset (this Deal/note series) but not found in servicer's Loan Terms & Prior Period",
        })

    scope_note = f", scoped to Deal '{deal}'" if deal_scoped else " (Deal could not be determined for this file -- compared by servicer only, not Servicer + Deal)"
    return _result(
        "POP-01", "Population Check", "SBA Reconciliation", "FAIL", records,
        lambda n: f"{n} loan(s) present in one source but missing from the other for {servicer_name}{scope_note}.",
    )


# ---------------------------------------------------------------------------
# BAL-01 — Balance Continuity
# This cycle's Beginning Balance (Servicer Inputs) should tie to SOUP's most
# recent Ending Actual Balance for that loan. Loans with multiple payment
# postings this cycle are flagged as WARN (informational) rather than FAIL,
# since their balance naturally reflects more than one payment.
# ---------------------------------------------------------------------------

def balance_continuity_check(servicer_inputs_df, soup_servicer_df):
    records = []

    required = {"SBA GP Number", "Beginning Balance"}
    if not required.issubset(servicer_inputs_df.columns) or "SBA GP Number" not in soup_servicer_df.columns:
        return _result("BAL-01", "Balance Continuity", "SBA Reconciliation", "FAIL", records,
                        lambda n: "")

    si = servicer_inputs_df.copy()
    si["SBA GP Number"] = _norm_id(si["SBA GP Number"])

    dup_mask = si.duplicated("SBA GP Number", keep=False)
    dup_ids = set(si.loc[dup_mask, "SBA GP Number"])

    # Beginning Balance comes from the FIRST chronological posting this
    # cycle, per client-confirmed multiple-posting rule — NOT the last
    # posting. Multiple postings alone shouldn't change which balance is
    # compared; only which specific posting's Beginning Balance value is
    # used. The continuity formula/tolerance below are unchanged.
    si_consolidated = consolidate_cycle_postings(si)

    soup = soup_servicer_df.copy()
    soup["SBA GP Number"] = _norm_id(soup["SBA GP Number"])
    sort_month = "Payment Month" if "Payment Month" in soup.columns else None
    if sort_month:
        soup[sort_month] = pd.to_datetime(soup[sort_month], errors="coerce")
        soup_latest = soup.sort_values(["SBA GP Number", sort_month]).groupby("SBA GP Number").tail(1)
    else:
        soup_latest = soup.groupby("SBA GP Number").tail(1)
    soup_latest = soup_latest.set_index("SBA GP Number")

    for gp, row in si_consolidated.iterrows():
        if gp not in soup_latest.index:
            continue  # already caught by the population check
        prior_ending = soup_latest.loc[gp].get("Ending Actual Balance")
        beginning = row.get("beginning_balance")
        if pd.isna(prior_ending) or pd.isna(beginning):
            continue
        diff = round(float(beginning) - float(prior_ending), 2)
        if abs(diff) <= BALANCE_TOLERANCE:
            continue
        # Multiple postings alone no longer downgrade this to a WARN — now
        # that Beginning Balance is correctly sourced from the FIRST
        # chronological posting (not an arbitrary last-posting value),
        # posting_count > 1 doesn't explain away a genuine continuity
        # break. "Multiple Postings This Cycle" stays as informational
        # visibility (DUP-01 already covers this in more detail); it no
        # longer affects severity.
        records.append({
            "SBA GP Number": gp,
            "Beginning Balance (this cycle)": round(float(beginning), 2),
            "SOUP Prior Ending Balance": round(float(prior_ending), 2),
            "Difference": diff,
            "Multiple Postings This Cycle": "YES" if gp in dup_ids else "NO",
        })

    details = "No issues found." if not records else (
        f"{len(records)} loan(s) with an unexplained balance break between this cycle's Beginning "
        f"Balance (from the first chronological posting) and SOUP's prior Ending Balance."
    )

    return {
        "id": "BAL-01",
        "name": "Balance Continuity",
        "category": "SBA Reconciliation",
        "severity": "FAIL",
        "status": "PASS" if not records else "FAIL",
        "issues_found": len(records),
        "details": details,
        "records": records,
    }


# ---------------------------------------------------------------------------
# LIQ-01 — Liquidation Review (formerly BAL-02, converted per client
# clarification)
#
# CLIENT-CONFIRMED RULE: there is no historical data available to
# independently determine that a loan "should" transition into LIQUIDATION.
# The engine must NEVER infer, predict, approve, or reject a LIQUIDATION
# status from a loan's prior SOUP status — if sufficient supporting data
# doesn't exist, the engine does not guess. Instead, the servicer's own
# report of Loan Status = LIQUIDATION is itself what triggers review: EVERY
# such loan is surfaced here, unconditionally — this is NOT contingent on
# any exception existing. A $0 Ending Balance is still a review item; a
# nonzero Ending Balance is a review item PLUS a clearly identified balance
# problem (the Remittance Form Instructions require LIQUIDATION loans to
# report $0).
#
# This replaces the old BAL-02, which only ever surfaced the nonzero-balance
# subset — silently passing every $0-balance LIQUIDATION loan with no
# visibility at all. Converted rather than kept alongside a separate check
# because every old BAL-02 exception is a strict subset of what LIQ-01 now
# covers; keeping both would show the same loan twice for overlapping
# reasons rather than complementary ones.
#
# Delinquency Status here is NOT invented separately — it reuses the exact
# same shared _calculate_expected_status engine as STAT-01/NEG-01 (80%
# threshold, zero-payment progression, prior baseline, deferral logic,
# UNVERIFIED/REVIEW handling, all unchanged). See _calculate_expected_status's
# is_liquidation_this_cycle carve-out for why LIQUIDATION loans now get a
# real calculation instead of the old terminal-status CURRENT shortcut.
#
# A servicer-reported LIQUIDATION loan is presentation-excluded from STAT-01
# (see status_match_check) — this check is where it belongs instead.
# ---------------------------------------------------------------------------

_LIQUIDATION_KEYWORD = "LIQUIDAT"  # substring match — same convention as
                                    # _TERMINAL_STATUS_KEYWORDS and the
                                    # STAT-01 exclusion, so the two checks
                                    # stay exact complements of each other.

_LIQUIDATION_FIELD_COLUMNS = (
    "Liquidation Date", "Liquidation Proceeds",
    "Subsequent Recoveries/(Losses)", "Liquidation Loss",
)


def _json_safe_value(value, is_date_field=False):
    """Converts one raw cell value into something json.dumps can handle.
    Calamine/pandas can hand back datetime.date/Timestamp/numpy scalar
    types depending on the source cell — none of those are JSON-
    serializable as-is, which crashed the /api/audit response the first
    time a real Liquidation Date reached this code path. Reuses the
    existing _fmt_date formatting convention for date fields so the display
    format matches every other date already shown elsewhere in the app."""
    if is_date_field:
        return _fmt_date(value)
    if hasattr(value, "item"):  # numpy scalar (int64/float64/etc.)
        return value.item()
    return value


def _consolidate_liquidation_fields(raw_group):
    """For ONE GP's raw postings this cycle, returns {column: value} for
    each of the four liquidation-specific fields, per the client's explicit
    no-guess rule for this scenario:
      - 0 or 1 posting, or all postings agree (including all-blank): that
        single/agreed value is used directly — this is NOT an aggregation
        choice, there's nothing to pick between.
      - Multiple postings with genuinely DIFFERENT non-blank values: do NOT
        pick first/last/max/min/sum. Instead the value becomes an explicit
        "AMBIGUOUS — see raw values" marker listing every distinct value
        found, and the field name is returned separately so the caller can
        surface it as a data-quality warning rather than a confident number.
    Returns (values: dict, ambiguous_fields: list[str])."""
    values = {}
    ambiguous_fields = []
    for col in _LIQUIDATION_FIELD_COLUMNS:
        if col not in raw_group.columns:
            values[col] = ""
            continue
        col_values = raw_group[col].tolist()
        non_blank = [v for v in col_values if v is not None and not (isinstance(v, float) and pd.isna(v)) and str(v).strip() != ""]
        distinct = []
        for v in non_blank:
            if v not in distinct:
                distinct.append(v)
        if len(distinct) <= 1:
            values[col] = _json_safe_value(distinct[0], is_date_field=(col == "Liquidation Date")) if distinct else ""
        else:
            # Already JSON-safe: the f-string converts every element (dates
            # included) to its string form via repr, same as the ambiguous
            # branch anywhere else in this file.
            values[col] = f"AMBIGUOUS — postings disagree: {distinct}"
            ambiguous_fields.append(col)
    return values, ambiguous_fields


def liquidation_review_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name=None):
    records = []

    required = {"SBA GP Number", "Loan Status", "Ending Balance"}
    if not required.issubset(servicer_inputs_df.columns):
        return _result("LIQ-01", "Liquidation Review", "SBA Reconciliation", "WARN", records,
                        lambda n: "")

    cycle_month = _infer_cycle_month(soup_servicer_df)
    has_payment_cols = {"Scheduled Payment", "Payment Received"}.issubset(servicer_inputs_df.columns)

    # Evaluate the CONSOLIDATED cycle record per GP — final chronological
    # Loan Status and final chronological Ending Balance — not each raw
    # posting independently. A loan with multiple postings this cycle has
    # one true end-of-cycle state; checking intermediate postings on their
    # own could flag (or miss) based on a mid-cycle row that isn't what the
    # loan actually ended the cycle as.
    consolidated = consolidate_cycle_postings(servicer_inputs_df)

    # Raw (pre-consolidation) rows, grouped by GP, used ONLY to check the
    # four liquidation-specific fields for multi-posting disagreement — see
    # _consolidate_liquidation_fields. Deliberately separate from the
    # shared consolidate_cycle_postings() used by every other check, since
    # no other check needs these fields or this ambiguity handling.
    raw = servicer_inputs_df.copy()
    raw["SBA GP Number"] = _norm_id(raw["SBA GP Number"])
    raw_by_gp = {gp: group for gp, group in raw.groupby("SBA GP Number")}

    soup = soup_servicer_df.copy() if soup_servicer_df is not None else pd.DataFrame()
    soup_latest = pd.DataFrame()
    if not soup.empty and "SBA GP Number" in soup.columns:
        soup["SBA GP Number"] = _norm_id(soup["SBA GP Number"])
        sort_month = "Payment Month" if "Payment Month" in soup.columns else None
        if sort_month:
            soup[sort_month] = pd.to_datetime(soup[sort_month], errors="coerce")
        soup_latest = (soup.sort_values(["SBA GP Number", sort_month]) if sort_month else soup).groupby("SBA GP Number").tail(1)
        soup_latest = soup_latest.set_index("SBA GP Number")

    for gp, row in consolidated.iterrows():
        loan_status = str(row.get("loan_status", "")).strip().upper()
        if _LIQUIDATION_KEYWORD not in loan_status:
            continue

        soup_status = str(soup_latest.loc[gp].get("Loan Status", "")).strip() if gp in soup_latest.index else ""
        soup_dq_raw = soup_latest.loc[gp].get("Delinquency Status", "") if gp in soup_latest.index else ""
        srf_dq = str(row.get("dq_status", "")).strip()
        scheduled_ambiguous = bool(row.get("scheduled_ambiguous"))

        # Same shared engine as STAT-01/NEG-01 — no separate liquidation
        # delinquency logic. NOT used to infer or contest the LIQUIDATION
        # status itself (expected_loan_status is deliberately ignored below
        # — the servicer's own report is what puts this loan here, full
        # stop); only calculated_dq/memo_warning are used.
        calculated_dq, _, _, memo_warning, warning_type, _ = _calculate_expected_status(
            gp, row.get("loan_status", ""), soup_status, soup_dq_raw, row.get("received"), row.get("scheduled"),
            has_payment_cols, memo_df, servicer_name, cycle_month,
            scheduled_ambiguous=scheduled_ambiguous,
        )

        if calculated_dq == "UNVERIFIED":
            dq_match = "REVIEW"
        else:
            dq_match = srf_dq.upper() == calculated_dq.upper()

        ending_balance = row.get("ending_balance")
        has_balance_figure = ending_balance is not None and not pd.isna(ending_balance)
        ending_balance_is_zero = has_balance_figure and abs(ending_balance) <= BALANCE_TOLERANCE

        balance_issue = ""
        if has_balance_figure and not ending_balance_is_zero:
            balance_issue = (
                f"LIQUIDATION loans are required to report a $0 balance. "
                f"Reported Ending Balance: {round(float(ending_balance), 2)}."
            )

        liq_field_values, liq_ambiguous_fields = _consolidate_liquidation_fields(raw_by_gp.get(gp, raw.iloc[0:0]))

        data_quality_warning_parts = []
        if memo_warning:
            data_quality_warning_parts.append(memo_warning)
        if liq_ambiguous_fields:
            data_quality_warning_parts.append(
                f"Multiple postings this cycle report DIFFERENT values for: "
                f"{', '.join(liq_ambiguous_fields)}. No aggregation rule is confirmed for this "
                f"case, so each field shows every distinct reported value rather than a single "
                f"picked one — review the underlying postings directly."
            )
        if scheduled_ambiguous:
            data_quality_warning_parts.append(
                f"Multiple distinct nonzero Scheduled Payment values this cycle "
                f"({row.get('scheduled_distinct_values')}) — cycle obligation is ambiguous; "
                f"the Calculated Delinquency Status above is NOT definitive."
            )

        records.append({
            "SBA GP Number": gp,
            "Lender Loan Number": row.get("lender_loan_number", ""),
            "Prior SOUP Loan Status": soup_status,
            "Servicer Loan Status": str(row.get("loan_status", "")).strip(),
            "Prior SOUP Delinquency Status": str(soup_dq_raw).strip(),
            "Servicer Delinquency Status": srf_dq,
            "Calculated Delinquency Status": calculated_dq,
            "DQ Status Match": dq_match if dq_match == "REVIEW" else ("YES" if dq_match else "NO"),
            "Beginning Balance": round(float(row["beginning_balance"]), 2) if pd.notna(row.get("beginning_balance")) else "",
            "Ending Balance": round(float(ending_balance), 2) if has_balance_figure else "",
            "Ending Balance Is Zero": "YES" if ending_balance_is_zero else ("NO" if has_balance_figure else ""),
            "Payment Received": round(float(row["received"]), 2) if has_payment_cols and pd.notna(row.get("received")) else "",
            "Scheduled Payment": round(float(row["scheduled"]), 2) if has_payment_cols and pd.notna(row.get("scheduled")) else "",
            "Liquidation Date": liq_field_values["Liquidation Date"],
            "Liquidation Proceeds": liq_field_values["Liquidation Proceeds"],
            "Subsequent Recoveries/(Losses)": liq_field_values["Subsequent Recoveries/(Losses)"],
            "Liquidation Loss": liq_field_values["Liquidation Loss"],
            "Posting Count": int(row.get("posting_count", 1)),
            "Balance Issue": balance_issue,
            "Data Quality Warning": " ".join(data_quality_warning_parts),
        })

    details = "No issues found." if not records else (
        f"{len(records)} loan(s) reported by the servicer as LIQUIDATION this cycle — every "
        f"such loan is surfaced here for review, regardless of Ending Balance."
    )

    return {
        "id": "LIQ-01",
        "name": "Liquidation Review",
        "category": "SBA Reconciliation",
        "severity": "WARN",
        "status": "PASS" if not records else "WARN",
        "issues_found": len(records),
        "details": details,
        "records": records,
    }


# ---------------------------------------------------------------------------
# STAT-01 — Status / Delinquency Match
#
# Loan Status: NOT a plain string comparison against SOUP's last recorded
# status — see _VALID_EXIT_TRANSITION_KEYWORDS (arriving at PAID-IN-FULL is
# a confirmed normal transition, never flagged on its own) and
# status_match_check's "Reason" tagging (REO/FORECLOSURE/BANKRUPTCY/
# LIQUIDATION and other transitions remain review items — surfaced, not
# treated as automatically wrong). Deferral phase can also override the
# expected Loan Status to DEFERRED — see _calculate_expected_status.
#
# Delinquency Status is independently CALCULATED (not string-compared)
# via _apply_payment_transition, using SOUP's last recorded status as the
# starting missed-payment count and this cycle's Payment Received vs.
# Scheduled Payment to determine the transition:
#
#   received <= 0                          -> missed_count + 1 (always a
#                                              miss, client-confirmed,
#                                              regardless of Scheduled
#                                              Payment — see zero-payment
#                                              validation)
#   0 < ratio < PAYMENT_SUFFICIENCY_THRESHOLD (80%, client-confirmed)
#                                           -> missed_count + 1
#   PAYMENT_SUFFICIENCY_THRESHOLD <= ratio < 1.0
#                                           -> missed_count unchanged
#   ratio >= 1.0 (catch-up payment)        -> missed_count reduced by
#                                              round(ratio) - 1, floored at 0
#
# Status ladder (uncapped upward in missed_count):
#   0 missed = CURRENT, 1 = DELINQUENT, 2 = 30 DAYS, 3 = 60 DAYS,
#   4 = 90 DAYS, 5 = 120 DAYS, 6 = 150 DAYS, 7 = 180 DAYS, ... (+30/miss)
#
# The 80% threshold is CONFIRMED, not provisional — see
# PAYMENT_SUFFICIENCY_THRESHOLD's own comment below for the full evidence
# trail. This resolves what once looked like a conflict with the formal
# instruction "partial payments are considered delinquent": the instruction
# establishes the principle, 80% is the client-confirmed number that
# operationalizes it. A partial payment below 80% IS treated as delinquent;
# one at or above 80% counts as that installment paid.
# ---------------------------------------------------------------------------

def _status_to_missed_count(status_str):
    """Parses a Delinquency Status label into a missed-payment count."""
    s = str(status_str).strip().upper()
    if not s or s in ("NAN", "NONE"):
        return None
    if s == "CURRENT":
        return 0
    if s == "DELINQUENT":
        return 1
    # e.g. "30 DAYS", "120 DAYS ", "150+ DAYS"
    digits = "".join(ch for ch in s.split("DAY")[0] if ch.isdigit())
    if not digits:
        return None
    days = int(digits)
    if days <= 0:
        return 0
    return (days // 30) + 1


def _missed_count_to_status(missed_count):
    """Maps a missed-payment count back to a Delinquency Status label."""
    if missed_count <= 0:
        return "CURRENT"
    if missed_count == 1:
        return "DELINQUENT"
    days = (missed_count - 1) * 30
    return f"{days} DAYS"


# *** CLIENT-CONFIRMED VALUE ***
# Client confirmed: 80% is the correct payment-sufficiency threshold.
# This resolves the apparent conflict with the Remittance Form
# Instructions' general statement that "partial payments are considered
# delinquent" — the formal instruction establishes the principle, and the
# client confirmation supplies the specific number that operationalizes it
# (a payment below 80% of one scheduled installment counts as a miss; 80%
# or above counts as that installment paid). No further confirmation
# needed on this value. Every comparison in this file MUST reference this
# constant, never a hardcoded literal, so any future change to the
# confirmed value only needs to happen in this one place.
PAYMENT_SUFFICIENCY_THRESHOLD = 0.80


def _apply_payment_transition(prior_missed_count, payment_received, scheduled_payment):
    """Returns the new missed-payment count after this cycle's payment.

    A period is only treated as "paid" if the amount received covers at
    least PAYMENT_SUFFICIENCY_THRESHOLD (80%) of the scheduled payment for
    that installment — receiving *some* payment isn't enough on its own.
    A payment below that threshold moves the loan down one delinquency
    rank, exactly the same as receiving nothing at all. Payments covering
    multiple installments' worth (e.g. a catch-up payment) still reduce
    the missed count by the number of full installments covered, each
    judged against the same 80% threshold."""
    if payment_received is None or pd.isna(payment_received):
        payment_received = 0
    received = float(payment_received)

    # a $0 payment is always a miss, regardless of what the Scheduled
    # Payment field says — some servicer submissions zero out BOTH fields
    # together for a no-payment month, which shouldn't be mistaken for
    # "no scheduled amount exists" and silently skipped
    if received <= 0:
        return prior_missed_count + 1

    if scheduled_payment is None or pd.isna(scheduled_payment) or scheduled_payment == 0:
        # can't compute a ratio without a valid scheduled amount to divide
        # by — hold steady rather than risk a divide-by-zero
        return prior_missed_count

    ratio = received / float(scheduled_payment)

    if ratio < PAYMENT_SUFFICIENCY_THRESHOLD:
        # received something, but less than 80% of one scheduled payment --
        # treated as insufficient, same as a full miss
        return prior_missed_count + 1
    if ratio < 1.0:
        # received at least 80% but less than 100% of one installment --
        # counts as that installment paid (held steady, not reduced)
        return prior_missed_count
    # ratio >= 1.0: this is a catch-up payment covering one or more
    # installments. The 80% sufficiency threshold only governs whether a
    # SINGLE period counts as paid (per the business rule above) -- it does
    # not redefine how multi-installment catch-up payments are measured, so
    # this uses standard rounding to the nearest whole installment, exactly
    # as before (e.g. a payment of 1.77x one scheduled installment rounds to
    # covering 2 installments, correctly catching up a loan that missed one
    # prior payment).
    installments_covered = round(ratio)
    if installments_covered <= 1:
        return prior_missed_count
    return max(0, prior_missed_count - (installments_covered - 1))


# ---------------------------------------------------------------------------
# Deferral awareness for STAT-01, reverse-engineered from CBT/Millennium/
# Mission Valley examples cross-checked against SOUP + Servicer_Memo_Summary:
#
# FULL P&I deferral (Servicer Action doesn't say "principal only"/"interest
# only" — see file_readers.classify_deferral_type):
#   - The cycle containing Deferral Start Date: FREEZE. Delinquency rank
#     holds at whatever it was the prior cycle, regardless of payment.
#   - Every cycle after that, through Deferral End Date: rank advances by
#     exactly +1 per cycle, regardless of payment received.
#   - Loan Status reads DEFERRED for the whole window above.
#   - CLIENT-CONFIRMED: the cycle CONTAINING Regular Payment Resume Date
#     (not Deferral End Date) is the resume cycle. Its delinquency baseline
#     resets to CURRENT and normal payment evaluation applies THAT SAME
#     cycle — paid in full -> CURRENT, not paid -> DELINQUENT — rather than
#     carrying forward whatever rank the freeze/advance sequence left it at.
#     Loan Status is no longer forced to DEFERRED from this cycle on.
#     (This was previously dead code — the window comparison excluded the
#     resume month itself, so the resume cycle silently fell through to
#     plain normal calc using the deferral-inflated prior rank as its
#     baseline. Fixed per client confirmation.)
#
# PARTIAL (principal-only / interest-only) deferral:
#   - No freeze, no auto-advance. Normal missed-payment math applies
#     unchanged — confirmed against Petro-Chem's two deferral windows,
#     where a $0 due/$0 received cycle still advanced a rank exactly like a
#     normal miss, and a fully-paid interest-only cycle held CURRENT.
#   - Loan Status still reads DEFERRED for the window, even though DQ
#     status is calculated normally.
#
# Matching a loan to the RIGHT memo row matters — a loan can have multiple
# memo entries across different report dates (repeat deferrals, or the same
# deferral re-reported later). We pick the row whose window actually covers
# the current cycle; if more than one does, we take the most recently
# reported one (highest Report Date).
# ---------------------------------------------------------------------------

def _infer_cycle_month(soup_servicer_df):
    """Best-effort guess at the calendar month this audit cycle covers: one
    month after SOUP's latest recorded Payment Month for this servicer (SOUP
    only has data through the PRIOR cycle — this cycle's servicer file is
    what's being audited against it). Returns a Timestamp normalized to the
    1st of the month, or None if it can't be determined (e.g. empty SOUP
    history) — callers should treat None as "skip deferral-window logic,"
    not raise, so this degrades safely to the pre-deferral-aware behavior."""
    if soup_servicer_df is None or soup_servicer_df.empty or "Payment Month" not in soup_servicer_df.columns:
        return None
    months = pd.to_datetime(soup_servicer_df["Payment Month"], errors="coerce").dropna()
    if months.empty:
        return None
    latest = months.max().replace(day=1)
    return (latest + pd.DateOffset(months=1)).replace(day=1)


def _lookup_deferral(memo_df, gp, servicer_key, cycle_month):
    """Finds the memo row (if any) whose deferral window covers cycle_month
    for this loan, and returns (deferral_info, warning, warning_type):

      deferral_info: None, or {"deferral_type": "full" | "partial",
                                "phase": "start" | "mid" | "resume" |
                                         "resume_unverified" | None}
      warning: None, or a short string explaining why deferral status
               couldn't be confidently verified.
      warning_type: None, "lender_identity", or "deferral_dates" —
               distinguishes "we can't even trust which memo record applies
               to this loan" (lender_identity) from "we found the right
               memo record, but its dates don't let us confidently place
               this cycle in the deferral timeline" (deferral_dates). These
               get different Reason tags in status_match_check: "MEMO DATA
               QUALITY WARNING" vs "DEFERRAL DATA QUALITY WARNING".

    servicer_key must be the canonical soup_key (e.g. "Millennium Bank",
    "Lendistry"), NOT a note-series-suffixed display label like "Millennium
    2025" — the memo file has no note-series field, so a suffixed label can
    never match it.
    """
    if memo_df is None or memo_df.empty or gp is None or cycle_month is None:
        return None, None, None
    if "gp_sba" not in memo_df.columns and "gp_lender" not in memo_df.columns:
        return None, None, None

    id_match = pd.Series(False, index=memo_df.index)
    if "gp_sba" in memo_df.columns:
        id_match = id_match | (memo_df["gp_sba"] == gp)
    if "gp_lender" in memo_df.columns:
        id_match = id_match | (memo_df["gp_lender"] == gp)

    if not id_match.any():
        return None, None, None  # no memo entry exists for this loan at all

    lender_match = pd.Series(True, index=memo_df.index)
    using_alias_table = False
    if "Lender" in memo_df.columns and servicer_key:
        known_aliases = MEMO_LENDER_ALIASES.get(servicer_key)
        lender_values = memo_df["Lender"].astype("string").str.strip()
        if known_aliases:
            # Centralized, explicit match: only the known exact Lender
            # spellings for this servicer (case-insensitive).
            using_alias_table = True
            alias_set = {a.strip().lower() for a in known_aliases}
            lender_match = lender_values.str.lower().isin(alias_set)
        else:
            # servicer_key isn't in the centralized alias table yet — fall
            # back to the prior fuzzy substring heuristic rather than
            # hard-failing, so a not-yet-catalogued servicer degrades
            # gracefully instead of losing deferral matching outright.
            def _norm_servicer(s):
                return str(s).strip().lower().replace("&", "and")
            servicer_norm = _norm_servicer(servicer_key)
            lender_match = lender_values.apply(
                lambda s: pd.notna(s) and (_norm_servicer(s) in servicer_norm or servicer_norm in _norm_servicer(s))
            )

    if not (id_match & lender_match).any():
        # This loan DOES have memo entries under this GP number, but none
        # matched a recognized Lender identity for servicer_key. Do not
        # silently treat this as "no deferral" — surface it.
        seen_lenders = sorted(memo_df.loc[id_match, "Lender"].dropna().unique().tolist()) if "Lender" in memo_df.columns else []
        alias_note = "" if using_alias_table else " (this servicer has no entry in the centralized memo lender alias table yet)"
        warning = (
            f"{int(id_match.sum())} Servicer Memo row(s) exist for this loan but none matched a known "
            f"Lender identity for '{servicer_key}'{alias_note}. Lender value(s) seen in the memo: "
            f"{seen_lenders}. Deferral status could not be verified from the memo — falling back to "
            f"normal calculation, which may be WRONG if a deferral is actually recorded under one of "
            f"these unrecognized spellings."
        )
        return None, warning, "lender_identity"

    candidates = memo_df[id_match & lender_match].copy()

    # Securitization Action is the memo's own authoritative categorical
    # field for what kind of record this is (Deferred / Charge-off /
    # Modification / an ad-hoc "Change Memo #N" note / etc.) — only rows
    # explicitly marked "Deferred" are real deferral records. This matters:
    # some administrative "Change Memo" rows carry the same Deferral Start/
    # End Date values as a real deferral (apparently copied from it) but
    # aren't deferrals themselves, and including them causes false matches.
    if "Securitization Action" in candidates.columns:
        candidates = candidates[candidates["Securitization Action"].astype("string").str.strip().str.lower() == "deferred"]

    if candidates.empty or "Deferral Start Date" not in candidates.columns:
        return None, None, None
    candidates = candidates.dropna(subset=["Deferral Start Date"])
    if candidates.empty:
        return None, None, None

    start_month = candidates["Deferral Start Date"].dt.to_period("M").dt.to_timestamp()
    resume_month = candidates["Regular Payment Resume Date"].dt.to_period("M").dt.to_timestamp() \
        if "Regular Payment Resume Date" in candidates.columns else pd.Series(pd.NaT, index=candidates.index)
    cycle_period = cycle_month.to_period("M").to_timestamp()

    # CLIENT-CONFIRMED: Regular Payment Resume Date is the controlling
    # field for the resume cycle — Deferral End Date is NEVER used as a
    # substitute resume date, even when Resume Date is blank. (This
    # replaces an earlier End Date + 1 month fallback, which had already
    # caused one real false-positive match before the Securitization
    # Action filter was added — inferring a controlling date is exactly
    # the kind of confident-but-wrong assertion that fallback produced.)
    #
    # A row can still be "still relevant" to this cycle without a known
    # resume date — that's precisely the resume_unverified case (state 5
    # below), not a reason to drop the row from matching entirely.
    started = start_month <= cycle_period
    still_relevant = started & (resume_month.isna() | (cycle_period <= resume_month))
    active = candidates[still_relevant]
    if active.empty:
        return None, None, None

    # multiple matching rows -> take the most recently reported memo entry
    if "Report Date" in active.columns and active["Report Date"].notna().any():
        active = active.sort_values("Report Date")
    chosen = active.iloc[-1]

    deferral_type = chosen.get("deferral_type", "full")
    chosen_start_month = pd.Timestamp(chosen["Deferral Start Date"]).to_period("M").to_timestamp()
    chosen_end_month = pd.Timestamp(chosen["Deferral End Date"]).to_period("M").to_timestamp() \
        if pd.notna(chosen.get("Deferral End Date")) else None
    chosen_resume_month = pd.Timestamp(chosen["Regular Payment Resume Date"]).to_period("M").to_timestamp() \
        if pd.notna(chosen.get("Regular Payment Resume Date")) else None

    warning = None
    warning_type = None
    if deferral_type == "full":
        # Checked in this order:
        #   1. start — the cycle containing Deferral Start Date
        #   2. resume — the cycle containing an ACTUAL Regular Payment
        #      Resume Date (checked before "mid" so a resume date landing
        #      in the same month as Deferral End Date correctly wins, per
        #      client-confirmed phase precedence)
        #   3. mid — any cycle after start, on/before Deferral End Date
        #   4. resume_unverified — past Deferral End Date with no known
        #      Resume Date; do NOT infer one
        if cycle_period == chosen_start_month:
            phase = "start"
        elif chosen_resume_month is not None and cycle_period == chosen_resume_month:
            phase = "resume"
        elif chosen_end_month is not None and cycle_period <= chosen_end_month:
            phase = "mid"
        elif chosen_resume_month is None:
            phase = "resume_unverified"
            warning_type = "deferral_dates"
            warning = (
                "Deferral End Date has passed, but Regular Payment Resume Date is blank. "
                "The audit cannot verify when normal payment evaluation should resume, because "
                "the client confirmed that Regular Payment Resume Date is the controlling field."
            )
        else:
            # cycle_period > chosen_end_month but also > chosen_resume_month
            # -- shouldn't be reachable, since `still_relevant` above already
            # excludes any cycle past a KNOWN resume month. Kept as a safe,
            # conservative fallback rather than assuming "mid" indefinitely.
            phase = "resume_unverified"
            warning_type = "deferral_dates"
            warning = (
                "This loan's deferral phase could not be determined confidently for this cycle "
                "— surfaced for manual review rather than guessing."
            )
    else:
        phase = None

    return {"deferral_type": deferral_type, "phase": phase}, warning, warning_type


# A Loan Status comparison of "prior != current" is not the same question as
# "is this transition legitimate" — the SRF Instructions tab explicitly
# defines PAID-IN-FULL as a valid, expected exit: "The loan paid off in the
# reporting period. After a loan has been PIF it is no longer necessary to
# include that loan on future reports." So arriving AT this status this
# cycle is never itself a mismatch, regardless of what the loan's prior
# SOUP-recorded status was (ACTIVE, DEFERRED, etc. -> PAID-IN-FULL is a
# normal loan lifecycle event, not an audit exception).
#
# Deliberately narrow and directional: this only ever suppresses a flag when
# the CURRENT reported status is one of these — it says nothing about
# leaving one of these statuses for something else (e.g. a loan reactivating
# out of LIQUIDATION), which is a different, unvalidated scenario and stays
# flagged under the normal comparison. Only PAID-IN-FULL is included here
# because only it has explicit textual confirmation from the Instructions
# tab; other terminal-looking statuses (LIQUIDATION, FORECLOSURE, REO,
# BANKRUPTCY, CANCEL, WITHDRAWN) have NOT been confirmed the same way and
# are intentionally left out rather than assumed to belong here too.
_VALID_EXIT_TRANSITION_KEYWORDS = ("PAID-IN-FULL", "PIF")


def _calculate_expected_status(gp, srf_status, soup_status, soup_dq_raw, received, scheduled,
                                has_payment_cols, memo_df, servicer_name, cycle_month,
                                scheduled_ambiguous=False):
    """The single shared decision tree for 'what should this loan's Loan
    Status and Delinquency Status be this cycle' — used by BOTH STAT-01 and
    NEG-01 so the two checks can never diverge on the same loan/cycle (they
    previously ran two separate, independently-maintained copies of this
    logic — NEG-01's copy didn't even know about deferrals).

    scheduled_ambiguous (from consolidate_cycle_postings): True when
    multiple DISTINCT nonzero Scheduled Payment values exist for this
    loan's postings this cycle — the cycle obligation is unresolved. Any
    branch below that would divide Payment Received by Scheduled Payment
    (normal calc, Full P&I resume, partial/Principal-Only deferral)
    returns "UNVERIFIED" instead of a real DQ label when this is True,
    rather than silently holding the prior rank steady and presenting that
    as a confident calculation. Full P&I start/mid (payment is deliberately
    ignored) and terminal-status handling are UNAFFECTED — neither branch
    ever consults Scheduled Payment, so an ambiguous value can't corrupt
    them.

    Returns (calculated_dq, expected_loan_status, deferral, memo_warning,
    warning_type, is_terminal_this_cycle). expected_loan_status is None
    when normal comparison against SOUP's last status applies (no deferral
    override)."""
    srf_status_u = str(srf_status).upper()
    soup_status_u = str(soup_status).upper()

    # A loan reported as paid in full, liquidated, or otherwise in a
    # terminal state has no ongoing payment schedule to be delinquent
    # against — its correct calculated status is CURRENT, regardless of
    # what the missed-payment math would otherwise say. Check BOTH the
    # servicer's self-reported status this cycle AND SOUP's last recorded
    # status — either source flagging a terminal status is enough.
    is_terminal_this_cycle = any(
        k in srf_status_u or k in soup_status_u for k in _TERMINAL_STATUS_KEYWORDS
    )
    # Client clarification: LIQUIDATION loans still need a REAL calculated
    # Delinquency Status (same validated engine as any other loan — 80%
    # threshold, zero-payment progression, prior baseline, etc.), routed to
    # the dedicated LIQ-01 review rather than STAT-01. Every other terminal
    # status (PAID-IN-FULL, FORECLOSURE, CHARGE-OFF, REO, CANCEL, WITHDRAWN)
    # keeps the existing CURRENT shortcut below, unchanged.
    is_liquidation_this_cycle = "LIQUIDAT" in srf_status_u or "LIQUIDAT" in soup_status_u

    deferral, memo_warning, warning_type = (None, None, None) if is_terminal_this_cycle else _lookup_deferral(memo_df, gp, servicer_name, cycle_month)

    prior_missed = _status_to_missed_count(soup_dq_raw) if has_payment_cols else None
    expected_loan_status = None

    if is_terminal_this_cycle and not is_liquidation_this_cycle:
        calculated_dq = "CURRENT"
    elif deferral and deferral["deferral_type"] == "full" and deferral["phase"] == "resume":
        # Client-confirmed: the cycle containing Regular Payment Resume
        # Date resets the delinquency baseline to CURRENT, then applies
        # normal payment evaluation THIS cycle — the rank accumulated
        # during the deferral window is explicitly NOT carried forward.
        if scheduled_ambiguous:
            calculated_dq = "UNVERIFIED"
        elif has_payment_cols:
            new_missed = _apply_payment_transition(0, received, scheduled)
            calculated_dq = _missed_count_to_status(new_missed)
        else:
            calculated_dq = "CURRENT"

        # BUG FIX (GP 5658839107): the resume cycle is the loan LEAVING
        # DEFERRED — that's the expected, correct outcome of a confirmed
        # resume, not a mismatch. Leaving expected_loan_status as None here
        # made it fall back to comparing against soup_status, which is
        # SOUP's PRIOR (still-DEFERRED) month — incorrectly flagging every
        # legitimate deferral exit as a Loan Status mismatch. Accept
        # whatever non-deferred status the servicer reports as the expected
        # outcome of this transition (same "arriving is not a mismatch"
        # principle already applied to PAID-IN-FULL).
        #
        # The one case NOT auto-accepted: if the servicer is STILL
        # reporting DEFERRED on a cycle the memo confirms should have
        # resumed, that's a real contradiction worth a human look — surface
        # it as a deferral data-quality review item rather than silently
        # accepting or silently flagging it as a hard mismatch.
        if "DEFERRED" in srf_status_u:
            expected_loan_status = None  # matches SOUP's still-DEFERRED prior; won't hard-fail, but...
            warning_type = "deferral_dates"
            memo_warning = (
                "The Servicer Memo Summary confirms this cycle is the Full P&I deferral resume "
                "cycle (Regular Payment Resume Date falls in this reporting month), but the "
                "current servicer file still reports Loan Status = DEFERRED. This is a "
                "contradiction between the memo and the servicer's own current-cycle report — "
                "surfaced for manual review rather than assumed to be an error on either side."
            )
        else:
            expected_loan_status = srf_status
    elif deferral and deferral["deferral_type"] == "full" and deferral["phase"] == "resume_unverified":
        # Client-confirmed: past Deferral End Date with no known Regular
        # Payment Resume Date. Do NOT advance the deferral rank (that would
        # assume still-deferred), do NOT reset to CURRENT (that would
        # assume a confirmed resume), and do NOT force Loan Status to
        # DEFERRED. Show the last confirmed value rather than asserting a
        # new one — memo_warning (from _lookup_deferral) is what actually
        # surfaces this for manual review; the record is forced to appear
        # regardless of whether this coincidentally "matches."
        expected_loan_status = None
        calculated_dq = str(soup_dq_raw).strip().upper()
    elif deferral and deferral["deferral_type"] == "full":
        expected_loan_status = "DEFERRED"
        if deferral["phase"] == "start":
            calculated_dq = _missed_count_to_status(prior_missed) if prior_missed is not None else str(soup_dq_raw).strip().upper()
        else:  # "mid"
            calculated_dq = _missed_count_to_status(prior_missed + 1) if prior_missed is not None else str(soup_dq_raw).strip().upper()
    elif deferral and deferral["deferral_type"] == "partial":
        # Loan Status should read DEFERRED for the window, but DQ still
        # follows normal payment-vs-scheduled math below.
        expected_loan_status = "DEFERRED"
        if scheduled_ambiguous:
            calculated_dq = "UNVERIFIED"
        elif has_payment_cols and prior_missed is not None:
            new_missed = _apply_payment_transition(prior_missed, received, scheduled)
            calculated_dq = _missed_count_to_status(new_missed)
        else:
            calculated_dq = str(soup_dq_raw).strip().upper()
    elif has_payment_cols and prior_missed is not None:
        if scheduled_ambiguous:
            calculated_dq = "UNVERIFIED"
        else:
            new_missed = _apply_payment_transition(prior_missed, received, scheduled)
            calculated_dq = _missed_count_to_status(new_missed)
    else:
        # fall back to the original straight string comparison if payment
        # columns aren't available or prior status is unparseable
        calculated_dq = str(soup_dq_raw).strip().upper()

    # DEFERRAL DATA QUALITY SAFEGUARD: no memo-confirmed deferral window
    # applies to this cycle (deferral is None), yet SOUP's prior status
    # and/or the servicer's current status say DEFERRED anyway. Normal
    # calc above already ran (using SOUP's prior DQ as its baseline, same
    # as any other loan) — this doesn't override that calculation, it just
    # makes sure the uncertainty is surfaced rather than silently trusted,
    # since "DEFERRED per SOUP/servicer, but no supporting memo evidence"
    # is exactly the kind of contradiction that shouldn't be quietly
    # resolved either way.
    if deferral is None and not is_terminal_this_cycle and not memo_warning:
        soup_says_deferred = "DEFERRED" in soup_status_u
        srf_says_deferred = "DEFERRED" in srf_status_u
        if soup_says_deferred and srf_says_deferred:
            warning_type = "deferral_dates"
            memo_warning = (
                "Prior SOUP and the current servicer file both report this loan as DEFERRED, but "
                "the available Servicer Memo information does not provide sufficient deferral "
                "dates to verify the current deferral phase."
            )
        elif soup_says_deferred:
            warning_type = "deferral_dates"
            memo_warning = (
                "Prior SOUP reports the loan as DEFERRED, but the available Servicer Memo "
                "information does not provide sufficient deferral dates to verify the current "
                "deferral phase."
            )
        elif srf_says_deferred:
            warning_type = "deferral_dates"
            memo_warning = (
                "The current servicer file reports the loan as DEFERRED, but the available "
                "Servicer Memo information does not provide sufficient deferral dates to verify "
                "that status."
            )

    return calculated_dq, expected_loan_status, deferral, memo_warning, warning_type, is_terminal_this_cycle


def status_match_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name=None):
    records = []

    required = {"SBA GP Number", "Loan Status", "Delinquency Status"}
    if not required.issubset(servicer_inputs_df.columns) or not required.issubset(soup_servicer_df.columns):
        return _result("STAT-01", "Status / Delinquency Match", "SBA Reconciliation", "WARN", records,
                        lambda n: "")

    cycle_month = _infer_cycle_month(soup_servicer_df)
    has_payment_cols = {"Scheduled Payment", "Payment Received"}.issubset(servicer_inputs_df.columns)

    # Cycle-level consolidation, per the client-confirmed multiple-posting
    # rule: chronologically-last Loan Status/Delinquency Status, summed
    # Payment Received, single (not summed) Scheduled Payment. Same shared
    # helper BAL-01, PAY-02, and NEG-01 all use now, so they can't diverge.
    si_grouped = consolidate_cycle_postings(servicer_inputs_df)

    soup = soup_servicer_df.copy()
    soup["SBA GP Number"] = _norm_id(soup["SBA GP Number"])
    sort_month = "Payment Month" if "Payment Month" in soup.columns else None
    if sort_month:
        soup[sort_month] = pd.to_datetime(soup[sort_month], errors="coerce")
    soup_latest = (soup.sort_values(["SBA GP Number", sort_month]) if sort_month else soup).groupby("SBA GP Number").tail(1)
    soup_latest = soup_latest.set_index("SBA GP Number")

    for gp, row in si_grouped.iterrows():
        if gp not in soup_latest.index:
            continue
        srf_status = str(row.get("loan_status", "")).strip()
        srf_dq = str(row.get("dq_status", "")).strip()
        soup_status = str(soup_latest.loc[gp].get("Loan Status", "")).strip()
        soup_dq_raw = soup_latest.loc[gp].get("Delinquency Status", "")
        scheduled_ambiguous = bool(row.get("scheduled_ambiguous"))

        # Client clarification: there is no historical data to independently
        # establish that LIQUIDATION is an incorrect transition from any
        # prior status, so a servicer-reported LIQUIDATION loan must never
        # be flagged here as a Loan Status mismatch. It belongs exclusively
        # in the dedicated LIQ-01 review instead — see liquidation_review_check.
        if "LIQUIDAT" in srf_status.upper():
            continue

        calculated_dq, expected_loan_status, deferral, memo_warning, warning_type, _ = _calculate_expected_status(
            gp, srf_status, soup_status, soup_dq_raw, row.get("received"), row.get("scheduled"),
            has_payment_cols, memo_df, servicer_name, cycle_month,
            scheduled_ambiguous=scheduled_ambiguous,
        )

        # UNVERIFIED means the calculation branch that ran actually needed
        # Scheduled Payment and couldn't trust it — this is NOT the same as
        # a confirmed delinquency mismatch, so it gets its own match state
        # ("REVIEW") rather than a hard YES/NO. Do not let the prior status
        # happening to differ from what was reported get relabeled as
        # DELINQUENCY STATUS MISMATCH just because calculated_dq isn't a
        # real status string.
        if calculated_dq == "UNVERIFIED":
            dq_match = "REVIEW"
        else:
            dq_match = srf_dq.upper() == calculated_dq.upper()

        # "Status changed" and "status mismatch" are not the same question.
        # A transition INTO a confirmed valid exit status (PAID-IN-FULL —
        # see _VALID_EXIT_TRANSITION_KEYWORDS) is normal loan lifecycle
        # activity, not an exception, regardless of what the prior status
        # was. This does NOT suppress the reverse direction — a loan
        # leaving a terminal status for something else stays flagged by the
        # normal comparison below, since that's a materially different and
        # unvalidated scenario.
        is_valid_exit_transition = any(k in srf_status.upper() for k in _VALID_EXIT_TRANSITION_KEYWORDS)
        status_match = is_valid_exit_transition or (srf_status.upper() == (expected_loan_status or soup_status).upper())

        # A memo-related warning is surfaced even when the calculated status
        # happens to match what was reported — the point is that this
        # loan's deferral/memo context couldn't be verified at all, so a
        # coincidental match is not the same as a confirmed one. Same
        # principle for an ambiguous Scheduled Payment: the row must stay
        # visible even if Loan Status happens to match, since the DQ side
        # is explicitly unresolved (dq_match == "REVIEW", never treated as
        # a silent pass).
        if status_match and dq_match is True and not memo_warning and not scheduled_ambiguous:
            continue

        # Two distinct memo-related problems get two distinct Reason tags
        # and two distinct message columns, per warning_type from
        # _lookup_deferral / _calculate_expected_status:
        #   "lender_identity" — can't even trust which memo record applies
        #       to this loan (a Lender-spelling match failure)
        #   "deferral_dates"  — found the right memo record, but its dates
        #       don't let us confidently place this cycle in the deferral
        #       timeline (missing Resume Date, or DEFERRED status with no
        #       supporting memo evidence at all)
        # A short, unambiguous categorical tag so this row is immediately
        # distinguishable in the UI/Excel from a genuine calculation
        # mismatch — a loan flagged ONLY because of a data-quality gap
        # (Loan Status Match = YES, DQ Status Match = YES/REVIEW) should
        # never be visually confused with a real delinquency discrepancy.
        # Lists ALL applicable reasons rather than picking just one, since
        # a record can genuinely have more than one at once.
        reasons = []
        if not status_match:
            reasons.append("LOAN STATUS MISMATCH")
        # DELINQUENCY STATUS MISMATCH is only ever asserted when the DQ
        # result is a real, independently-established value — never when
        # it's UNVERIFIED (dq_match == "REVIEW"), per the ambiguous-
        # Scheduled-Payment rule: don't classify a row as a delinquency
        # mismatch merely because the prior status happens to differ from
        # what was reported, when the calculation itself couldn't run.
        if dq_match is False:
            reasons.append("DELINQUENCY STATUS MISMATCH")
        if memo_warning and warning_type == "lender_identity":
            reasons.append("MEMO DATA QUALITY WARNING")
        if memo_warning and warning_type == "deferral_dates":
            reasons.append("DEFERRAL DATA QUALITY WARNING")
        if scheduled_ambiguous:
            reasons.append("SCHEDULED PAYMENT AMBIGUOUS")

        records.append({
            "SBA GP Number": gp,
            "Reason": " + ".join(reasons),
            "Servicer Loan Status": srf_status,
            "Expected Loan Status": expected_loan_status or soup_status,
            "Loan Status Match": "YES" if status_match else "NO",
            "Servicer DQ Status": srf_dq,
            "Prior (SOUP) DQ Status": str(soup_dq_raw).strip(),
            "Calculated DQ Status": calculated_dq,
            "DQ Status Match": dq_match if dq_match == "REVIEW" else ("YES" if dq_match else "NO"),
            "Deferral Type": deferral["deferral_type"] if deferral else "",
            "Memo Lookup Warning": memo_warning if warning_type == "lender_identity" else "",
            "Deferral Data Quality Warning": memo_warning if warning_type == "deferral_dates" else "",
            "Posting Count": int(row.get("posting_count", 1)),
            "Scheduled Payment Ambiguous": "YES" if scheduled_ambiguous else "NO",
            "Data Quality Issue": (
                f"Multiple distinct nonzero Scheduled Payment values this cycle "
                f"({row.get('scheduled_distinct_values')}) — cycle obligation is ambiguous; "
                f"the Calculated DQ Status above is NOT definitive."
            ) if scheduled_ambiguous else "",
            "Payment Received (this cycle)": round(float(row.get("received")), 2) if has_payment_cols and pd.notna(row.get("received")) else "",
            "Scheduled Payment (this cycle)": round(float(row.get("scheduled")), 2) if has_payment_cols and pd.notna(row.get("scheduled")) else "",
        })

    return _result(
        "STAT-01", "Status / Delinquency Match", "SBA Reconciliation", "WARN", records,
        lambda n: f"{n} loan(s) where Loan Status doesn't match SOUP's last recorded status, or "
                  f"where Delinquency Status doesn't match the status calculated from "
                  f"missed-payment history (Payment Received vs. Scheduled Payment against the "
                  f"loan's prior recorded status).",
    )


# ---------------------------------------------------------------------------
# NEG-01 — Negative Payment Flags
# Flags any individual payment posting this cycle with a negative Payment
# Received amount (typically a reversal of a prior payment). This is purely
# a visibility flag for manual review — it does NOT alter the delinquency
# status calculation in STAT-01 in any way. The "current calculated status"
# shown here is computed independently, using the exact same missed-payment
# math as STAT-01, purely so a reviewer can see what the system currently
# thinks the status is without having to cross-reference another tab.
# ---------------------------------------------------------------------------

def negative_payment_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name=None):
    records = []

    required = {"SBA GP Number", "Payment Received"}
    if not required.issubset(servicer_inputs_df.columns):
        return _result("NEG-01", "Negative Payment Flags", "SBA Reconciliation", "WARN", records,
                        lambda n: "")

    si = servicer_inputs_df.copy()
    si["SBA GP Number"] = _norm_id(si["SBA GP Number"])
    si["Payment Received"] = pd.to_numeric(si["Payment Received"], errors="coerce")
    has_payment_cols = {"Scheduled Payment", "Payment Received"}.issubset(si.columns)

    cycle_month = _infer_cycle_month(soup_servicer_df)

    # Same shared consolidation and shared delinquency calculation STAT-01
    # uses — this is the actual fix here: NEG-01 previously ran its own
    # independent aggregation (with the same missing-sort-order bug as
    # STAT-01 used to have) AND its own independent calculation that didn't
    # even know about deferrals, so it could silently disagree with STAT-01
    # on the exact same loan/cycle. Reusing both means that's no longer
    # possible.
    si_grouped = consolidate_cycle_postings(si)

    soup = soup_servicer_df.copy() if soup_servicer_df is not None else pd.DataFrame()
    soup_latest = pd.DataFrame()
    if not soup.empty and "SBA GP Number" in soup.columns:
        soup["SBA GP Number"] = _norm_id(soup["SBA GP Number"])
        sort_month = "Payment Month" if "Payment Month" in soup.columns else None
        if sort_month:
            soup[sort_month] = pd.to_datetime(soup[sort_month], errors="coerce")
        soup_latest = (soup.sort_values(["SBA GP Number", sort_month]) if sort_month else soup).groupby("SBA GP Number").tail(1)
        soup_latest = soup_latest.set_index("SBA GP Number")

    calculated_dq_by_loan = {}
    prior_status_by_loan = {}
    for gp, row in si_grouped.iterrows():
        soup_status = str(soup_latest.loc[gp].get("Loan Status", "")).strip() if gp in soup_latest.index else ""
        soup_dq_raw = soup_latest.loc[gp].get("Delinquency Status", "") if gp in soup_latest.index else ""
        prior_status_by_loan[gp] = str(soup_dq_raw).strip()

        srf_status = str(row.get("loan_status", "")).strip()
        calculated_dq, _, _, _, _, _ = _calculate_expected_status(
            gp, srf_status, soup_status, soup_dq_raw, row.get("received"), row.get("scheduled"),
            has_payment_cols, memo_df, servicer_name, cycle_month,
            scheduled_ambiguous=bool(row.get("scheduled_ambiguous")),
        )
        # NEG-01 only ever DISPLAYS calculated_dq for manual-review context
        # (Current Period Status (Calculated)) — it never re-derives a
        # match/mismatch verdict from it, so "UNVERIFIED" safely passes
        # through as plain text here rather than risking being treated as
        # a real status label.
        calculated_dq_by_loan[gp] = calculated_dq

    # scan the RAW, un-aggregated rows -- a negative posting can be masked by
    # a positive one in the same cycle's summed total, so each posting needs
    # to be checked individually rather than relying on the aggregate sum
    negative_rows = si[si["Payment Received"] < 0]
    for _, row in negative_rows.iterrows():
        gp = row["SBA GP Number"]
        period = _fmt_date(row.get("Payment Receipt Date")) if "Payment Receipt Date" in si.columns else ""
        records.append({
            "SBA GP Number": gp,
            "Period / Month": period,
            "Negative Payment Amount": round(float(row["Payment Received"]), 2),
            "Prior Period Status": prior_status_by_loan.get(gp, ""),
            "Current Period Status (Calculated)": calculated_dq_by_loan.get(gp, ""),
            "Note": "Reversed payment — review delinquency status manually",
        })

    records.sort(key=lambda r: (r["SBA GP Number"], r["Period / Month"]))

    return _result(
        "NEG-01", "Negative Payment Flags", "SBA Reconciliation", "WARN", records,
        lambda n: f"{n} payment posting(s) with a negative Payment Received amount this cycle "
                  f"(typically a reversal) — flagged for manual review. Delinquency status "
                  f"calculations are NOT auto-adjusted for these; a reviewer should confirm the "
                  f"correct status by hand.",
    )


# ---------------------------------------------------------------------------
# DUP-01 — Duplicate Postings
# Informational: a loan with more than one row in Servicer Inputs this cycle.
# ---------------------------------------------------------------------------

def _fmt_date(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    try:
        ts = pd.Timestamp(value)
        if pd.isna(ts):
            return ""
        return ts.strftime("%m/%d/%Y")
    except (TypeError, ValueError):
        return str(value)


def duplicate_postings_check(servicer_inputs_df):
    records = []

    if "SBA GP Number" not in servicer_inputs_df.columns:
        return {
            "id": "DUP-01", "name": "Duplicate Postings", "category": "SBA Reconciliation",
            "severity": "INFO", "status": "PASS", "issues_found": 0,
            "details": "No issues found.", "records": [],
        }

    si = servicer_inputs_df.copy()
    si["SBA GP Number"] = _norm_id(si["SBA GP Number"])
    if "Payment Received" in si.columns:
        si["Payment Received"] = pd.to_numeric(si["Payment Received"], errors="coerce")
    dup_mask = si.duplicated("SBA GP Number", keep=False)
    dup_rows = si[dup_mask].sort_values("SBA GP Number")

    # consolidate one row per loan instead of one row per posting -- a loan
    # with 3 postings this cycle should appear ONCE here, with all 3
    # payments/dates combined, rather than as 3 separate rows repeating the
    # same Beginning/Ending Balance
    for gp, group in dup_rows.groupby("SBA GP Number", sort=False):
        # sort each loan's own postings into chronological order so the
        # combined list reads correctly and "last posting" (for Ending
        # Balance) is determined reliably. Prefer "Loan Payment Instance"
        # when the template provides it (an explicit 1, 2, 3... ordinal) --
        # more reliable than date, since same-day postings can't be
        # distinguished by date alone. Fall back to Payment Receipt Date.
        if "Loan Payment Instance" in group.columns:
            group = group.assign(_sort_key=pd.to_numeric(group["Loan Payment Instance"], errors="coerce")).sort_values("_sort_key")
        elif "Payment Receipt Date" in group.columns:
            group = group.assign(_sort_key=pd.to_datetime(group["Payment Receipt Date"], errors="coerce")).sort_values("_sort_key")

        first = group.iloc[0]
        last = group.iloc[-1]
        payment_parts = []
        total_received = 0.0
        for _, r in group.iterrows():
            amt = r.get("Payment Received", "")
            date_str = _fmt_date(r.get("Payment Receipt Date", "")) if "Payment Receipt Date" in group.columns else ""
            if pd.notna(amt) and amt != "":
                total_received += float(amt)
                amt_str = f"${float(amt):,.2f}"
            else:
                amt_str = ""
            payment_parts.append(f"{amt_str} ({date_str})" if date_str else amt_str)

        records.append({
            "SBA GP Number": gp,
            "Lender Loan Number": first.get("Lender Loan Number", ""),
            "Beginning Balance": first.get("Beginning Balance", ""),
            "Ending Balance": last.get("Ending Balance", ""),
            "Number of Postings": len(group),
            "Total Payment Received": round(total_received, 2),
            "Payment Details (amount / date)": "; ".join(p for p in payment_parts if p),
        })

    issues = len(records)
    details = (
        "No issues found." if not records else
        f"{issues} loan(s) with more than one payment posted this cycle "
        f"— informational only, not an error."
    )

    return {
        "id": "DUP-01",
        "name": "Duplicate Postings",
        "category": "SBA Reconciliation",
        "severity": "INFO",
        "status": "PASS" if not records else "INFO",
        "issues_found": issues,
        "details": details,
        "records": records,
    }


# ---------------------------------------------------------------------------
# PAY-01 — Missing Payment Periods
# For each loan still on the servicer's active population, looks at its full
# monthly reporting history in SOUP and flags:
#   - interior gaps: a month skipped in the middle of the loan's history
#   - stopped reporting: the loan hasn't shown up in SOUP for the most
#     recent month(s) even though it's still listed as active
# ---------------------------------------------------------------------------

# Deliberately a SEPARATE constant from _TERMINAL_STATUS_KEYWORDS, not an
# addition to it. _TERMINAL_STATUS_KEYWORDS is shared by STAT-01's
# delinquency calculation and PAY-02's shortfall check — folding
# "SUBSTITUTION-OUT" in there would silently change both of those checks'
# behavior for any substituted-out loan. This list answers a narrower, PAY-
# 01-specific question — "has this loan legitimately left the servicer's
# reporting population" — checked against SOUP's own last-recorded Loan
# Status for that loan, not the servicer's current-cycle Loan Terms sheet.
#
# Validated against real Community Bank & Trust June data: 9 loans PAY-01
# was flagging as "stopped reporting despite still being active" all have a
# final SOUP Loan Status of SUBSTITUTION-OUT — a real population exit, not a
# reporting gap. Only that status is included here; other statuses that
# happen to appear in _TERMINAL_STATUS_KEYWORDS (PAID, LIQUIDAT, FORECLOS,
# etc.) have NOT been validated against real PAY-01 output the same way, so
# they're intentionally left out rather than assumed to belong here too.
_POPULATION_EXIT_STATUS_KEYWORDS = ("SUBSTITUTION-OUT",)


def missing_payment_periods_check(soup_servicer_df, loan_terms_df):
    records = []

    if "SBA GP Number" not in soup_servicer_df.columns or "Payment Month" not in soup_servicer_df.columns:
        return _result("PAY-01", "Missing Payment Periods", "SBA Reconciliation", "FAIL", records,
                        lambda n: "")

    soup = soup_servicer_df.copy()
    soup["SBA GP Number"] = _norm_id(soup["SBA GP Number"])
    soup["Payment Month"] = pd.to_datetime(soup["Payment Month"], errors="coerce")
    soup = soup.dropna(subset=["Payment Month"])

    if soup.empty:
        return _result("PAY-01", "Missing Payment Periods", "SBA Reconciliation", "FAIL", records,
                        lambda n: "")

    servicer_latest_month = soup["Payment Month"].max()

    active_ids = None
    if loan_terms_df is not None and "SBA GP Number" in loan_terms_df.columns and not loan_terms_df.empty:
        lt = loan_terms_df.copy()
        if "Loan Status" in lt.columns:
            status_upper = lt["Loan Status"].astype("string").str.upper().fillna("")
            is_terminal = status_upper.apply(lambda s: any(k in s for k in _TERMINAL_STATUS_KEYWORDS))
            lt = lt[~is_terminal]
        active_ids = set(_norm_id(lt["SBA GP Number"]).dropna())

    for gp, group in soup.groupby("SBA GP Number"):
        if active_ids is not None and gp not in active_ids:
            continue  # loan isn't on the current active population — gaps after payoff are expected

        months = sorted(group["Payment Month"].unique())
        if len(months) < 2:
            continue

        first, last = months[0], months[-1]
        expected = pd.date_range(first, last, freq="MS")
        actual = set(months)
        missing = [m for m in expected if m not in actual]

        # SOUP's own last-recorded Loan Status for this loan (not the
        # servicer's current-cycle Loan Terms sheet, which active_ids above
        # is based on) — a loan can still be sitting in active_ids while
        # SOUP's history already shows it exited the population.
        last_month_status = ""
        if "Loan Status" in group.columns:
            last_month_statuses = group.loc[group["Payment Month"] == last, "Loan Status"] \
                .astype("string").str.upper().fillna("")
            if len(last_month_statuses):
                last_month_status = last_month_statuses.iloc[-1]
        population_exited = any(k in last_month_status for k in _POPULATION_EXIT_STATUS_KEYWORDS)

        # A loan that legitimately exited the population (e.g. substituted
        # out of the deal) is EXPECTED to stop appearing — that's not a
        # reporting gap. Interior-gap detection above is unaffected; this
        # only suppresses the "stopped reporting" flag itself.
        stopped_reporting = (last < servicer_latest_month) and not population_exited

        if not missing and not stopped_reporting:
            continue

        records.append({
            "SBA GP Number": gp,
            "First Reported Month": first.strftime("%m/%Y"),
            "Last Reported Month": last.strftime("%m/%Y"),
            "Interior Gap Month(s)": ", ".join(m.strftime("%m/%Y") for m in missing) if missing else "None",
            "Stopped Reporting": "YES" if stopped_reporting else "NO",
            "Most Recent Servicer Month": servicer_latest_month.strftime("%m/%Y"),
            "Last SOUP Loan Status": last_month_status,
        })

    records.sort(key=lambda r: (r["Stopped Reporting"] != "YES", r["Interior Gap Month(s)"] == "None"))

    stopped_count = sum(1 for r in records if r["Stopped Reporting"] == "YES")
    gap_count = sum(1 for r in records if r["Interior Gap Month(s)"] != "None")

    details = "No issues found." if not records else (
        f"{stopped_count} loan(s) stopped reporting before the most recent month "
        f"({servicer_latest_month.strftime('%m/%Y')}) despite still being active; "
        f"{gap_count} loan(s) have a skipped month somewhere in their reporting history."
    )

    return {
        "id": "PAY-01",
        "name": "Missing Payment Periods",
        "category": "SBA Reconciliation",
        "severity": "FAIL",
        "status": "PASS" if not records else "FAIL",
        "issues_found": len(records),
        "details": details,
        "records": records,
    }


# ---------------------------------------------------------------------------
# PAY-02 — Payment Shortfall Check
# Compares what the servicer actually remitted (Payment Received + Servicing
# Fee, summed across any multiple postings this cycle) against the
# Scheduled Payment for that loan. Most loans will show Payment Received a
# bit below Scheduled Payment simply because the servicing fee is netted out
# before remittance — that's expected and NOT flagged. Only genuine
# shortfalls (or overages) beyond the fee-adjusted amount are surfaced.
#
# Note: this check depends on Scheduled Payment and Servicing Fee columns
# being present in the servicer's file — not every servicer template may
# include Servicing Fee, so this check can be safely disabled if it doesn't
# apply to a particular servicer's reporting format.
# ---------------------------------------------------------------------------

SHORTFALL_TOLERANCE = 1.00  # dollars

# Loans in a terminal status shouldn't be compared against their normal
# monthly Scheduled Payment — a payoff, liquidation, or charge-off amount is
# expected to look nothing like a scheduled payment and isn't a real shortfall.
_TERMINAL_STATUS_KEYWORDS = ("PAID", "PIF", "LIQUIDAT", "FORECLOS", "CHARGE", "REO", "CANCEL", "WITHDRAWN")


def payment_shortfall_check(servicer_inputs_df):
    records = []

    required = {"SBA GP Number", "Scheduled Payment", "Payment Received"}
    if not required.issubset(servicer_inputs_df.columns):
        return {
            "id": "PAY-02", "name": "Payment Shortfall Check", "category": "SBA Reconciliation",
            "severity": "FAIL", "status": "PASS", "issues_found": 0,
            "details": "No issues found.", "records": [],
        }

    si = servicer_inputs_df.copy()

    if "Loan Status" in si.columns:
        status_upper = si["Loan Status"].astype("string").str.upper().fillna("")
        is_terminal = status_upper.apply(lambda s: any(k in s for k in _TERMINAL_STATUS_KEYWORDS))
        si = si[~is_terminal]

    si["SBA GP Number"] = _norm_id(si["SBA GP Number"])
    has_fee = "Servicing Fee" in si.columns
    if has_fee:
        si["Servicing Fee"] = pd.to_numeric(si["Servicing Fee"], errors="coerce").fillna(0)

    # Payment Received: summed (shared helper). Scheduled Payment: ONE
    # cycle obligation, NOT summed — per client-confirmed rule, multiple
    # postings in a cycle are multiple transactions against one monthly
    # obligation, not separate obligations to add together. Servicing Fee
    # is transaction-level and still summed here directly (that was
    # already the intended behavior, untouched).
    consolidated = consolidate_cycle_postings(si)
    fee_by_gp = si.groupby("SBA GP Number")["Servicing Fee"].sum() if has_fee else None

    for gp, row in consolidated.iterrows():
        received = row.get("received")
        fee = float(fee_by_gp.loc[gp]) if has_fee and gp in fee_by_gp.index else 0.0

        if row.get("scheduled_ambiguous"):
            # Multiple DISTINCT nonzero Scheduled Payment values for the
            # same loan/cycle — the cycle obligation is genuinely
            # ambiguous. Do NOT silently sum, max, or pick one; surface a
            # data quality warning instead of a computed shortfall.
            records.append({
                "SBA GP Number": gp,
                "Lender Loan Number": row.get("lender_loan_number", ""),
                "Scheduled Payment": "AMBIGUOUS",
                "Payment Received": round(float(received), 2) if pd.notna(received) else "",
                "Servicing Fee": round(fee, 2) if has_fee else "N/A",
                "Shortfall (Scheduled - Received - Fee)": "",
                "Direction": "",
                "Issue": f"Multiple distinct nonzero Scheduled Payment values this cycle "
                         f"({row.get('scheduled_distinct_values')}) — cycle obligation is "
                         f"ambiguous, shortfall not calculated.",
            })
            continue

        scheduled = row.get("scheduled")
        if scheduled is None or pd.isna(scheduled) or received is None or pd.isna(received):
            continue

        net_remitted = float(received) + fee
        shortfall = round(float(scheduled) - net_remitted, 2)
        if abs(shortfall) <= SHORTFALL_TOLERANCE:
            continue
        records.append({
            "SBA GP Number": gp,
            "Lender Loan Number": row.get("lender_loan_number", ""),
            "Scheduled Payment": round(float(scheduled), 2),
            "Payment Received": round(float(received), 2),
            "Servicing Fee": round(fee, 2) if has_fee else "N/A",
            "Shortfall (Scheduled - Received - Fee)": shortfall,
            "Direction": "Underpaid" if shortfall > 0 else "Overpaid",
            "Issue": "",
        })

    # ambiguity-warning rows surface first (unresolved data quality issue),
    # then real shortfalls by magnitude descending
    records.sort(key=lambda r: (
        0 if r["Issue"] else 1,
        -abs(r["Shortfall (Scheduled - Received - Fee)"]) if isinstance(r["Shortfall (Scheduled - Received - Fee)"], (int, float)) else 0,
    ))

    ambiguous_count = sum(1 for r in records if r["Issue"])
    shortfall_count = len(records) - ambiguous_count
    details = "No issues found." if not records else (
        f"{shortfall_count} loan(s) where the amount remitted (payment received plus servicing fee) "
        f"doesn't match the single cycle Scheduled Payment by more than ${SHORTFALL_TOLERANCE:.2f}"
        + (f"; {ambiguous_count} loan(s) with an ambiguous Scheduled Payment (multiple distinct "
           f"nonzero values this cycle) — shortfall not calculated for those." if ambiguous_count else ".")
        + " Paid-off, liquidated, and other terminal-status loans are excluded since their payoff "
          "amount isn't comparable to a scheduled monthly payment."
    )

    return {
        "id": "PAY-02",
        "name": "Payment Shortfall Check",
        "category": "SBA Reconciliation",
        "severity": "FAIL",
        "status": "PASS" if not records else "FAIL",
        "issues_found": len(records),
        "details": details,
        "records": records,
    }


# ---------------------------------------------------------------------------
# Runner — one servicer file against SOUP at a time
# ---------------------------------------------------------------------------

ALL_CHECKS_META = [
    ("POP-01", "Population Check"),
    ("BAL-01", "Balance Continuity"),
    ("LIQ-01", "Liquidation Review"),
    ("STAT-01", "Status / Delinquency Match"),
    ("PAY-01", "Missing Payment Periods"),
    ("PAY-02", "Payment Shortfall Check"),
    ("DUP-01", "Duplicate Postings"),
    ("NEG-01", "Negative Payment Flags"),
]


def run_servicer_reconciliation(servicer_name, loan_terms_df, servicer_inputs_df, soup_servicer_df, memo_df=None, memo_lender_key=None, deal=None):
    """Runs all checks for a single servicer file against its SOUP history.

    memo_df (from file_readers.read_service_memo) is optional — when not
    supplied, STAT-01 falls back to its pre-deferral-aware behavior
    (straight comparison against SOUP's last recorded status) rather than
    erroring, so this stays backward-compatible for callers that don't have
    a Servicer Memo Summary file to upload.

    memo_lender_key is the canonical servicer identity to match against the
    memo's "Lender" column — pass the same soup_key already used for
    filter_soup_by_servicer, NOT the display-facing servicer_name. This
    matters for any servicer split across multiple note-series SRF files
    (e.g. "Millennium 2025"/"Millennium 2024", "Lendistry 2025"/"Lendistry
    2024"): the memo file has no note-series field, so a display label like
    "Millennium 2025" never matches memo Lender values like "Millennium
    Bank" — confirmed to silently return zero matches for every Millennium
    loan, and to miss the ~28% of Lendistry memo rows recorded under
    "Lendistry SBLC, LLC, CA" rather than plain "Lendistry". Falls back to
    servicer_name if not supplied, for backward compatibility.

    deal (from file_readers.extract_deal_from_note_series) scopes POP-01 to
    Servicer + Deal + normalized GP Number instead of Servicer + GP Number
    alone — see population_check's docstring-level comment for why. Only
    affects POP-01; BAL-01, LIQ-01, STAT-01, PAY-01/02, DUP-01, and NEG-01
    are unchanged by this parameter."""
    results = [
        population_check(loan_terms_df, soup_servicer_df, servicer_name, deal=deal),
        balance_continuity_check(servicer_inputs_df, soup_servicer_df),
        liquidation_review_check(servicer_inputs_df, soup_servicer_df, memo_df=memo_df, servicer_name=memo_lender_key or servicer_name),
        status_match_check(servicer_inputs_df, soup_servicer_df, memo_df=memo_df, servicer_name=memo_lender_key or servicer_name),
        missing_payment_periods_check(soup_servicer_df, loan_terms_df),
        payment_shortfall_check(servicer_inputs_df),
        duplicate_postings_check(servicer_inputs_df),
        negative_payment_check(servicer_inputs_df, soup_servicer_df, memo_df=memo_df, servicer_name=memo_lender_key or servicer_name),
    ]
    # tag every result with which servicer it belongs to, so multi-servicer
    # runs can be told apart in the combined results list -- always the
    # display-facing servicer_name (e.g. "Millennium 2025"), never the memo
    # lookup key, so report tabs/summary still distinguish note series.
    for r in results:
        r["servicer"] = servicer_name
        r["id"] = f"{r['id']}::{servicer_name}"

    return results


def summarize(all_results, soup_len, servicer_loan_counts):
    total_issues = sum(r["issues_found"] for r in all_results)
    checks_passed = sum(1 for r in all_results if r["status"] == "PASS")
    checks_flagged = sum(1 for r in all_results if r["status"] in ("WARN", "FAIL"))
    checks_informational = sum(1 for r in all_results if r["status"] == "INFO")
    total_servicer_records = sum(servicer_loan_counts.values())

    return {
        "total_records": int(soup_len + total_servicer_records),
        "soup_records": int(soup_len),
        "servicer_records": int(total_servicer_records),
        # exact per-servicer counts (this cycle's Servicer Inputs row count),
        # so the chatbot can answer "how many records for X" precisely
        # instead of estimating from a truncated CSV sample.
        "servicer_record_counts": {k: int(v) for k, v in servicer_loan_counts.items()},
        "servicers_processed": list(servicer_loan_counts.keys()),
        "total_issues": int(total_issues),
        "checks_passed": int(checks_passed),
        "checks_flagged": int(checks_flagged),
        "checks_informational": int(checks_informational),
        "total_checks": len(all_results),
    }
