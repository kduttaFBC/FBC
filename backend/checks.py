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
# POP-01 — Population Check
# Every loan in the servicer's Loan Terms & Prior Period population should
# exist in the SOUP dataset's history for that servicer, and vice versa.
# ---------------------------------------------------------------------------

def population_check(loan_terms_df, soup_servicer_df, servicer_name):
    records = []

    lt_ids = set(_norm_id(loan_terms_df["SBA GP Number"]).dropna()) if "SBA GP Number" in loan_terms_df.columns else set()
    soup_ids = set(_norm_id(soup_servicer_df["SBA GP Number"]).dropna()) if "SBA GP Number" in soup_servicer_df.columns else set()

    only_srf = lt_ids - soup_ids
    only_soup = soup_ids - lt_ids

    for gp in sorted(only_srf):
        records.append({
            "SBA GP Number": gp,
            "Issue": "In servicer's Loan Terms & Prior Period but not found in SOUP dataset",
        })
    for gp in sorted(only_soup):
        records.append({
            "SBA GP Number": gp,
            "Issue": "In SOUP dataset but not found in servicer's Loan Terms & Prior Period",
        })

    return _result(
        "POP-01", "Population Check", "SBA Reconciliation", "FAIL", records,
        lambda n: f"{n} loan(s) present in one source but missing from the other for {servicer_name}.",
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

    # use the last posting per loan this cycle if there's a payment date to sort by.
    # Real-world files mix populated dates with blank strings ("" for loans with
    # no payment posted this cycle) in the same column — pandas can't sort a
    # column containing both str and datetime.date objects, so coerce to a
    # proper datetime dtype first (blanks/unparseable values become NaT, which
    # sorts safely to one end rather than raising).
    sort_col = "Payment Receipt Date" if "Payment Receipt Date" in si.columns else None
    if sort_col:
        si[sort_col] = pd.to_datetime(si[sort_col], errors="coerce")
        si_last = si.sort_values(sort_col).groupby("SBA GP Number").tail(1)
    else:
        si_last = si.groupby("SBA GP Number").tail(1)

    soup = soup_servicer_df.copy()
    soup["SBA GP Number"] = _norm_id(soup["SBA GP Number"])
    sort_month = "Payment Month" if "Payment Month" in soup.columns else None
    if sort_month:
        soup[sort_month] = pd.to_datetime(soup[sort_month], errors="coerce")
        soup_latest = soup.sort_values(["SBA GP Number", sort_month]).groupby("SBA GP Number").tail(1)
    else:
        soup_latest = soup.groupby("SBA GP Number").tail(1)
    soup_latest = soup_latest.set_index("SBA GP Number")

    for _, row in si_last.iterrows():
        gp = row["SBA GP Number"]
        if gp not in soup_latest.index:
            continue  # already caught by the population check
        prior_ending = soup_latest.loc[gp].get("Ending Actual Balance")
        beginning = row.get("Beginning Balance")
        if pd.isna(prior_ending) or pd.isna(beginning):
            continue
        diff = round(float(beginning) - float(prior_ending), 2)
        if abs(diff) <= BALANCE_TOLERANCE:
            continue
        records.append({
            "SBA GP Number": gp,
            "Beginning Balance (this cycle)": round(float(beginning), 2),
            "SOUP Prior Ending Balance": round(float(prior_ending), 2),
            "Difference": diff,
            "Multiple Postings This Cycle": "YES" if gp in dup_ids else "NO",
            "Severity": "WARN" if gp in dup_ids else "FAIL",
        })

    # Sort so FAILs surface first
    records.sort(key=lambda r: 0 if r["Severity"] == "FAIL" else 1)

    hard_fails = [r for r in records if r["Severity"] == "FAIL"]
    severity = "FAIL" if hard_fails else "WARN"
    status = "PASS" if not records else severity

    details = "No issues found." if not records else (
        f"{len(hard_fails)} loan(s) with an unexplained balance break; "
        f"{len(records) - len(hard_fails)} additional loan(s) flagged only because of multiple "
        f"payment postings this cycle."
    )

    return {
        "id": "BAL-01",
        "name": "Balance Continuity",
        "category": "SBA Reconciliation",
        "severity": "FAIL",
        "status": status,
        "issues_found": len(records),
        "details": details,
        "records": records,
    }


# ---------------------------------------------------------------------------
# STAT-01 — Status / Delinquency Match
# Loan Status reported this cycle should match SOUP's last recorded Loan
# Status for that loan (simple string comparison, unchanged).
#
# Delinquency Status is handled differently: rather than a straight string
# comparison against SOUP's last DQ status, this independently CALCULATES
# what the DQ status should be based on missed-payment math, using SOUP's
# last recorded status as the starting point and this cycle's Payment
# Received vs. Scheduled Payment to determine the transition:
#
#   multiple = round(Payment Received / Scheduled Payment)
#     multiple == 0  -> missed_count + 1   (status moves DOWN one rank)
#     multiple == 1  -> missed_count unchanged (status stays the SAME rank)
#     multiple >= 2  -> missed_count - (multiple - 1), floored at 0
#                       (status moves UP one rank per full extra payment)
#
# Status ladder (uncapped upward in missed_count):
#   0 missed = CURRENT, 1 = DELINQUENT, 2 = 30 DAYS, 3 = 60 DAYS,
#   4 = 90 DAYS, 5 = 120 DAYS, 6 = 150 DAYS, 7 = 180 DAYS, ... (+30/miss)
#
# NOTE: partial payments (0 < received < 1x scheduled) are currently treated
# the same as a full regular payment (multiple rounds to 1, status holds
# steady) per client direction — this is a placeholder pending refinement
# on how aggressively partials should round down to a "miss" instead.
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


PAYMENT_SUFFICIENCY_THRESHOLD = 0.80  # a payment covering >= 80% of one scheduled
                                       # installment counts as that installment paid


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


def status_match_check(servicer_inputs_df, soup_servicer_df):
    records = []

    required = {"SBA GP Number", "Loan Status", "Delinquency Status"}
    if not required.issubset(servicer_inputs_df.columns) or not required.issubset(soup_servicer_df.columns):
        return _result("STAT-01", "Status / Delinquency Match", "SBA Reconciliation", "WARN", records,
                        lambda n: "")

    si = servicer_inputs_df.copy()
    si["SBA GP Number"] = _norm_id(si["SBA GP Number"])
    has_payment_cols = {"Scheduled Payment", "Payment Received"}.issubset(si.columns)
    if has_payment_cols:
        si["Scheduled Payment"] = pd.to_numeric(si["Scheduled Payment"], errors="coerce")
        si["Payment Received"] = pd.to_numeric(si["Payment Received"], errors="coerce")

    # combine multiple postings this cycle into one period per loan: last
    # Loan Status/Delinquency Status reported. Payment Received is summed
    # (it genuinely represents total cash received this cycle), but
    # Scheduled Payment is NOT summed — multiple postings in one cycle
    # typically represent one payment period split into sub-periods (e.g.
    # a rate change or interim payment date splitting April 29-June 5 into
    # two rows), not two separate months' worth of dues. Summing Scheduled
    # Payment across postings inflates the monthly-due figure and causes
    # false DELINQUENT flags on loans that actually caught up. We use the
    # max of the postings' Scheduled Payment as the single recurring
    # monthly figure.
    agg_dict = {
        "loan_status": ("Loan Status", "last"),
        "dq_status": ("Delinquency Status", "last"),
    }
    if has_payment_cols:
        agg_dict["scheduled"] = ("Scheduled Payment", "max")
        agg_dict["received"] = ("Payment Received", "sum")
    si_grouped = si.groupby("SBA GP Number").agg(**agg_dict).reset_index()

    soup = soup_servicer_df.copy()
    soup["SBA GP Number"] = _norm_id(soup["SBA GP Number"])
    sort_month = "Payment Month" if "Payment Month" in soup.columns else None
    if sort_month:
        soup[sort_month] = pd.to_datetime(soup[sort_month], errors="coerce")
    soup_latest = (soup.sort_values(["SBA GP Number", sort_month]) if sort_month else soup).groupby("SBA GP Number").tail(1)
    soup_latest = soup_latest.set_index("SBA GP Number")

    for _, row in si_grouped.iterrows():
        gp = row["SBA GP Number"]
        if gp not in soup_latest.index:
            continue
        srf_status = str(row.get("loan_status", "")).strip()
        srf_dq = str(row.get("dq_status", "")).strip()
        soup_status = str(soup_latest.loc[gp].get("Loan Status", "")).strip()
        soup_dq_raw = soup_latest.loc[gp].get("Delinquency Status", "")

        status_match = srf_status.lower() == soup_status.lower()

        # A loan reported as paid in full, liquidated, or otherwise in a
        # terminal state has no ongoing payment schedule to be delinquent
        # against — its correct calculated status is CURRENT, regardless of
        # what the missed-payment math would otherwise say (e.g. a loan that
        # was 60 DAYS delinquent and then gets paid off is not still "90
        # DAYS delinquent" — it's paid off, full stop).
        #
        # Check BOTH the servicer's self-reported status this cycle AND
        # SOUP's last recorded status, rather than trusting either source
        # alone — a servicer's file could be wrong or lag behind, and SOUP
        # could equally be stale, so either source flagging a terminal
        # status is enough to treat the loan as closed out.
        is_terminal_this_cycle = any(
            k in srf_status.upper() or k in soup_status.upper()
            for k in _TERMINAL_STATUS_KEYWORDS
        )

        # calculate what DQ status should be, based on missed-payment math,
        # rather than a straight string comparison against SOUP's last DQ
        prior_missed = _status_to_missed_count(soup_dq_raw) if has_payment_cols else None
        if is_terminal_this_cycle:
            calculated_dq = "CURRENT"
            dq_match = srf_dq.upper() == calculated_dq
        elif has_payment_cols and prior_missed is not None:
            new_missed = _apply_payment_transition(prior_missed, row.get("received"), row.get("scheduled"))
            calculated_dq = _missed_count_to_status(new_missed)
            dq_match = srf_dq.upper() == calculated_dq
        else:
            # fall back to the original straight string comparison if
            # payment columns aren't available or prior status is unparseable
            calculated_dq = str(soup_dq_raw).strip().upper()
            dq_match = srf_dq.lower() == calculated_dq.lower()

        if status_match and dq_match:
            continue
        records.append({
            "SBA GP Number": gp,
            "Servicer Loan Status": srf_status,
            "SOUP Loan Status": soup_status,
            "Loan Status Match": "YES" if status_match else "NO",
            "Servicer DQ Status": srf_dq,
            "SOUP DQ Status": str(soup_dq_raw).strip(),
            "Calculated DQ Status": calculated_dq,
            "DQ Status Match": "YES" if dq_match else "NO",
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

def negative_payment_check(servicer_inputs_df, soup_servicer_df):
    records = []

    required = {"SBA GP Number", "Payment Received"}
    if not required.issubset(servicer_inputs_df.columns):
        return _result("NEG-01", "Negative Payment Flags", "SBA Reconciliation", "WARN", records,
                        lambda n: "")

    si = servicer_inputs_df.copy()
    si["SBA GP Number"] = _norm_id(si["SBA GP Number"])
    si["Payment Received"] = pd.to_numeric(si["Payment Received"], errors="coerce")
    has_scheduled_col = "Scheduled Payment" in si.columns
    if has_scheduled_col:
        si["Scheduled Payment"] = pd.to_numeric(si["Scheduled Payment"], errors="coerce")
    has_dq_col = "Delinquency Status" in si.columns

    # per-loan aggregate calculated status, using the identical missed-payment
    # math STAT-01 uses -- computed here independently so this tab can show
    # "what the system currently calculates" without touching status_match_check
    calculated_dq_by_loan = {}
    if has_dq_col:
        agg_dict = {"dq_status": ("Delinquency Status", "last")}
        if has_scheduled_col:
            agg_dict["scheduled"] = ("Scheduled Payment", "max")
        agg_dict["received"] = ("Payment Received", "sum")
        si_grouped = si.groupby("SBA GP Number").agg(**agg_dict).reset_index()
    else:
        si_grouped = pd.DataFrame(columns=["SBA GP Number"])

    soup = soup_servicer_df.copy() if soup_servicer_df is not None else pd.DataFrame()
    soup_latest = pd.DataFrame()
    if not soup.empty and "SBA GP Number" in soup.columns:
        soup["SBA GP Number"] = _norm_id(soup["SBA GP Number"])
        sort_month = "Payment Month" if "Payment Month" in soup.columns else None
        if sort_month:
            soup[sort_month] = pd.to_datetime(soup[sort_month], errors="coerce")
        soup_latest = (soup.sort_values(["SBA GP Number", sort_month]) if sort_month else soup).groupby("SBA GP Number").tail(1)
        soup_latest = soup_latest.set_index("SBA GP Number")

    for _, row in si_grouped.iterrows():
        gp = row["SBA GP Number"]
        soup_dq_raw = soup_latest.loc[gp].get("Delinquency Status", "") if gp in soup_latest.index else ""
        prior_missed = _status_to_missed_count(soup_dq_raw)
        if prior_missed is not None and has_scheduled_col:
            new_missed = _apply_payment_transition(prior_missed, row.get("received"), row.get("scheduled"))
            calculated_dq_by_loan[gp] = _missed_count_to_status(new_missed)
        else:
            calculated_dq_by_loan[gp] = str(soup_dq_raw).strip().upper() if soup_dq_raw != "" else ""

    prior_status_by_loan = {}
    if not soup_latest.empty:
        for gp in soup_latest.index:
            prior_status_by_loan[gp] = str(soup_latest.loc[gp].get("Delinquency Status", "")).strip()

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
    dup_mask = si.duplicated("SBA GP Number", keep=False)
    dup_rows = si[dup_mask].sort_values("SBA GP Number")

    for _, row in dup_rows.iterrows():
        records.append({
            "SBA GP Number": row.get("SBA GP Number"),
            "Lender Loan Number": row.get("Lender Loan Number", ""),
            "Beginning Balance": row.get("Beginning Balance", ""),
            "Ending Balance": row.get("Ending Balance", ""),
            "Payment Received": row.get("Payment Received", ""),
            "Payment Receipt Date": _fmt_date(row.get("Payment Receipt Date", "")),
        })

    issues = len(records)
    details = (
        "No issues found." if not records else
        f"{issues} row(s) reflecting loans with more than one payment posted this cycle "
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

        stopped_reporting = last < servicer_latest_month

        if not missing and not stopped_reporting:
            continue

        records.append({
            "SBA GP Number": gp,
            "First Reported Month": first.strftime("%m/%Y"),
            "Last Reported Month": last.strftime("%m/%Y"),
            "Interior Gap Month(s)": ", ".join(m.strftime("%m/%Y") for m in missing) if missing else "None",
            "Stopped Reporting": "YES" if stopped_reporting else "NO",
            "Most Recent Servicer Month": servicer_latest_month.strftime("%m/%Y"),
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
    si["Scheduled Payment"] = pd.to_numeric(si["Scheduled Payment"], errors="coerce")
    si["Payment Received"] = pd.to_numeric(si["Payment Received"], errors="coerce")
    has_fee = "Servicing Fee" in si.columns
    si["Servicing Fee"] = pd.to_numeric(si["Servicing Fee"], errors="coerce").fillna(0) if has_fee else 0

    # sum Scheduled Payment, Payment Received, and Servicing Fee across any
    # multiple postings this cycle — a loan with two postings usually means
    # two separate payment periods being caught up (see Interest-From/To),
    # so all three figures need to be summed together, not just the actual
    # payment figures against a single scheduled amount.
    grouped = si.groupby("SBA GP Number").agg(
        scheduled=("Scheduled Payment", "sum"),
        received=("Payment Received", "sum"),
        fee=("Servicing Fee", "sum"),
        lender_loan=("Lender Loan Number", "first"),
    ).reset_index()

    grouped["net_remitted"] = grouped["received"] + grouped["fee"]
    grouped["shortfall"] = (grouped["scheduled"] - grouped["net_remitted"]).round(2)

    for _, row in grouped.iterrows():
        if pd.isna(row["shortfall"]) or abs(row["shortfall"]) <= SHORTFALL_TOLERANCE:
            continue
        records.append({
            "SBA GP Number": row["SBA GP Number"],
            "Lender Loan Number": row["lender_loan"],
            "Scheduled Payment": round(float(row["scheduled"]), 2) if pd.notna(row["scheduled"]) else "",
            "Payment Received": round(float(row["received"]), 2),
            "Servicing Fee": round(float(row["fee"]), 2) if has_fee else "N/A",
            "Shortfall (Scheduled - Received - Fee)": row["shortfall"],
            "Direction": "Underpaid" if row["shortfall"] > 0 else "Overpaid",
        })

    records.sort(key=lambda r: -abs(r["Shortfall (Scheduled - Received - Fee)"]))

    details = "No issues found." if not records else (
        f"{len(records)} loan(s) where the amount remitted (payment received plus servicing fee) "
        f"doesn't match the scheduled payment by more than ${SHORTFALL_TOLERANCE:.2f} — after "
        f"accounting for the servicing fee that's normally netted out. Paid-off, liquidated, and "
        f"other terminal-status loans are excluded since their payoff amount isn't comparable to "
        f"a scheduled monthly payment."
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
    ("STAT-01", "Status / Delinquency Match"),
    ("PAY-01", "Missing Payment Periods"),
    ("PAY-02", "Payment Shortfall Check"),
    ("DUP-01", "Duplicate Postings"),
    ("NEG-01", "Negative Payment Flags"),
]


def run_servicer_reconciliation(servicer_name, loan_terms_df, servicer_inputs_df, soup_servicer_df):
    """Runs all checks for a single servicer file against its SOUP history."""
    results = [
        population_check(loan_terms_df, soup_servicer_df, servicer_name),
        balance_continuity_check(servicer_inputs_df, soup_servicer_df),
        status_match_check(servicer_inputs_df, soup_servicer_df),
        missing_payment_periods_check(soup_servicer_df, loan_terms_df),
        payment_shortfall_check(servicer_inputs_df),
        duplicate_postings_check(servicer_inputs_df),
        negative_payment_check(servicer_inputs_df, soup_servicer_df),
    ]
    # tag every result with which servicer it belongs to, so multi-servicer
    # runs can be told apart in the combined results list
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
