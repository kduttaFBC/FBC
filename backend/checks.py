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

import math
import re

import pandas as pd
from file_readers import normalize_id_series as _norm_id
from file_readers import MEMO_LENDER_ALIASES


BALANCE_TOLERANCE = 1.00  # dollars


def _round2(value):
    """Coerce a value to a 2-dp float, or None if it isn't numeric. Used
    only to expose already-computed display figures (payment, servicing
    fee, balance) on correction_data for the review page — never in any
    calculation."""
    v = pd.to_numeric(value, errors="coerce")
    return round(float(v), 2) if pd.notna(v) else None


def _result(check_id, name, category, severity, records, details_fn, issues_found=None):
    issues = len(records) if issues_found is None else issues_found
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
# POSTING SEQUENCE (needed before "last"/"effective_last" mean anything) is
# established by, in order of preference:
#   1. Loan Payment Instance, if populated (an explicit 1, 2, 3... ordinal —
#      more reliable than date, since same-day postings can't be
#      distinguished by date alone)
#   2. Payment Receipt Date
#   3. original row order (stable sort preserves this as the final fallback)
#
# CLIENT-CONFIRMED (2026-08-14): the CYCLE OPENING is a SEPARATE question and
# is resolved separately, by _cycle_opening_posting — the lender does not
# guarantee that a loan's earliest-dated posting is written first, so the
# first row of the sequence above is not reliably the row the cycle opened at.
# The sequence itself is deliberately UNCHANGED: measured across 16 real
# workbooks, Loan Payment Instance reproduces a coherent balance chain on more
# multi-posting loans than Payment Receipt Date does (159 vs 158 of 160), so
# the ordinal stays the sequencing mechanism and the date is used ONLY to find
# the opening row. See _cycle_opening_posting for the full measurement.
# ---------------------------------------------------------------------------

def _resolve_scheduled_payment_step_pattern(group, sched_series):
    """CLIENT-CONFIRMED (2026-08-04): more than one distinct nonzero
    Scheduled Payment value in a cycle doesn't automatically mean genuine
    ambiguity — the field can legitimately change mid-cycle (a
    re-amortization or correction taking effect). This checks whether the
    CHRONOLOGICAL sequence of nonzero values forms a clean, one-directional
    step pattern: each distinct value forms exactly one contiguous run,
    with no reversion to an earlier value once the sequence has moved past
    it. This allows any number of sequential updates (A, then B, then C)
    while still rejecting real oscillation (A, then B, then back to A) as
    genuinely ambiguous — oscillation is not a "later correction," it's
    exactly the kind of instability that should stay UNVERIFIED.

    When a "Last Payment Instance?" column is present, the final run's
    value is also cross-checked against whichever row is flagged 'Y' there
    — a contradiction is treated as ambiguous rather than trusted, an
    extra safety margin beyond the pattern shape alone.

    Returns the final chronological value if the pattern qualifies,
    otherwise None (caller falls back to today's UNVERIFIED behavior)."""
    nonzero_mask = sched_series.notna() & (sched_series != 0)
    chronological_values = sched_series[nonzero_mask].tolist()
    if not chronological_values:
        return None

    # Run-length compress: collapse consecutive duplicates into one entry
    # each, so we're checking the SHAPE of the sequence, not raw repeats.
    runs = []
    for v in chronological_values:
        if not runs or runs[-1] != v:
            runs.append(v)

    # A value reappearing in a later, non-adjacent run means the sequence
    # went back to something it had already moved past -- oscillation, not
    # a one-directional update.
    if len(set(runs)) != len(runs):
        return None

    final_value = runs[-1]

    if "Last Payment Instance?" in group.columns:
        last_flagged = group[group["Last Payment Instance?"].astype("string").str.strip().str.upper() == "Y"]
        if not last_flagged.empty:
            flagged_scheduled = pd.to_numeric(last_flagged.iloc[-1].get("Scheduled Payment"), errors="coerce")
            if pd.notna(flagged_scheduled) and flagged_scheduled != 0 and flagged_scheduled != final_value:
                return None  # contradicts the servicer's own "last posting" flag -- don't trust the pattern

    return final_value


def _is_blank_value(value):
    """CLIENT-CONFIRMED (2026-08-06): a value counts as blank for the
    incomplete-final-posting fallback below if it's NaN/None, OR an
    empty/whitespace-only string. Deliberately does NOT treat a real 0 or
    a real, meaningful zero-length business value as blank -- this is
    purely about detecting a genuinely EMPTY cell, not a legitimate zero
    balance or terminal status."""
    return pd.isna(value) or str(value).strip() == ""


def _cycle_opening_posting(group):
    """CLIENT-CONFIRMED (2026-08-14): the posting a loan's cycle OPENS at —
    which is a different question from where that posting sits in the posting
    SEQUENCE, and is resolved independently of it.

    WHY THE TWO ARE SEPARATE. Postings are sequenced by Loan Payment Instance
    (see the doctrine comment above consolidate_cycle_postings). In the SRF
    template that column is the Excel formula COUNTIFS($A$10:A<row>,A<row>) --
    it counts occurrences of the GP from the top of the sheet down, so it is an
    exact re-encoding of PHYSICAL ROW POSITION and carries no transaction-time
    information at all. Measured across 16 real workbooks, that is nonetheless
    the BETTER sequence signal: the servicer's own posting ordinal reproduces a
    coherent Beginning->Ending balance chain on 159 of 160 multi-posting loans,
    against 158 for Payment Receipt Date, and on the single loan where the two
    disagree (Lendistry 2024 July GP 4704259107) the ordinal is the UNIQUE
    correct chain and the date order breaks it. So the sequence stays on the
    ordinal and is deliberately NOT changed here.

    What the ordinal cannot answer is which posting the cycle STARTED from,
    because the lender does not guarantee that the earliest-dated payment is
    written first. Measured on Lendistry 2025 July GP 5951029110: the earliest
    posting (2026-06-29) is physically the SECOND row, so `group.iloc[0]`
    returned the 2026-07-01 row and Beginning Balance came back 19,504.69
    against a prior finalized SOUP Ending Actual Balance of 19,805.64 -- a
    -300.95 BAL-01 exception that is purely an artifact of row placement. The
    2026-06-29 posting carries Beginning Balance 19,805.64 and ties to the cent.

    THE ORDER, exactly:
        1. Payment Receipt Date ascending
        2. blank / unparseable Payment Receipt Date LAST (never treated as
           earliest -- an absent date must not win the cycle opening)
        3. Loan Payment Instance ascending, the deterministic tie-breaker
           (101 of 163 real multi-posting loans carry duplicate receipt dates,
           overwhelmingly same-day reversal pairs, so a date-only comparison
           would not be deterministic)
        4. original row index ascending, then position, so the key is a TOTAL
           order and the result cannot depend on dict/groupby iteration order

    Because clause 3 reproduces today's answer for every same-date group, and
    clause 1 only ever moves a loan whose earliest posting was written out of
    place, this changes exactly one loan in the 16 available real workbooks.

    Returns the opening posting as a Series. SCOPE: read ONLY for values that
    genuinely describe the START of the cycle (Beginning Balance, Interest-From).
    `last`/`effective_last` and every ending-state value -- Ending Balance, Loan
    Status, Delinquency Status, the Payment Received/Principal/Net Interest/
    Servicing Fee totals, reversal facts -- are untouched by this function."""
    if len(group) == 1:
        return group.iloc[0]

    dates = (pd.to_datetime(group["Payment Receipt Date"], errors="coerce")
             if "Payment Receipt Date" in group.columns
             else pd.Series(pd.NaT, index=group.index))
    instances = (pd.to_numeric(group["Loan Payment Instance"], errors="coerce")
                 if "Loan Payment Instance" in group.columns
                 else pd.Series(float("nan"), index=group.index))
    index_values = list(group.index)

    def opening_sort_key(position):
        date = dates.iloc[position]
        instance = instances.iloc[position]
        return (
            1 if pd.isna(date) else 0,           # blank/unparseable date sorts LAST
            0 if pd.isna(date) else date.value,
            1 if pd.isna(instance) else 0,       # blank instance sorts last within a date
            0.0 if pd.isna(instance) else float(instance),
            index_values[position],
            position,
        )

    return group.iloc[min(range(len(group)), key=opening_sort_key)]


# CLIENT-CONFIRMED (2026-08-12): the current-period terminal-event date columns
# whose PRESENCE corrections.py routes Loan Status eligibility on. Mapped to the
# consolidated key each becomes. The SRF template spells Paid-In-Full Date with
# hyphens; the reader preserves the workbook's own spelling, so these must match
# it exactly — a typo here yields a silently absent fact, not an error.
_TERMINAL_EVENT_DATE_COLUMNS = {
    "Bankruptcy Date": "bankruptcy_date",
    "Paid-In-Full Date": "paid_in_full_date",
    "Liquidation Date": "liquidation_date",
}


def _cycle_terminal_event_dates(group):
    """Returns {"<key>": display_string, "<key>_populated": bool} for each
    terminal-event date column, judged across EVERY posting in the cycle.

    A PRESENCE test, deliberately not a date parse — see the call site in
    consolidate_cycle_postings for why. Raw facts only: nothing in this module's
    calculation logic reads the result back."""
    out = {}
    for column, key in _TERMINAL_EVENT_DATE_COLUMNS.items():
        display = ""
        if column in group.columns:
            for value in group[column]:
                if not _is_blank_value(value):
                    display = str(value).strip()
                    break
        out[key] = display
        out[f"{key}_populated"] = bool(display)
    return out


# CLIENT-CONFIRMED (2026-08-13): additional current-period REVIEW EVIDENCE
# columns, carried for display on the Human Review loan card so a reviewer can
# judge a Loan Status contradiction without opening the workbook alongside the
# audit. DISPLAY EVIDENCE ONLY.
#
# Deliberately SEPARATE from _TERMINAL_EVENT_DATE_COLUMNS above, which is not
# merely a different list but a different KIND of thing: corrections.py routes
# Loan Status ELIGIBILITY on those (bankruptcy_date_populated is a write gate).
# Nothing routes on these. Merging the two dicts would silently widen what the
# eligibility layer could key on, so they stay apart.
_REVIEW_EVIDENCE_DATE_COLUMNS = {
    "Foreclosure Date": "foreclosure_date",
    "REO Date": "reo_date",
}

# Money columns, SUMMED across every posting in the cycle with min_count=1 --
# the same consolidation `received` / `principal` / `net_interest` already use,
# so a loan whose proceeds arrive across two postings reports the cycle total
# and an all-blank column stays None rather than becoming 0.00.
_REVIEW_EVIDENCE_AMOUNT_COLUMNS = {
    "Liquidation Proceeds": "liquidation_proceeds",
    "Liquidation Loss": "liquidation_loss",
    "Subsequent Recoveries/(Losses)": "subsequent_recoveries",
}


def _cycle_review_evidence(group):
    """Returns the additional display-only review evidence for one loan's cycle:
    {foreclosure_date, foreclosure_date_populated, reo_date, reo_date_populated,
     liquidation_proceeds, liquidation_loss, subsequent_recoveries}.

    Dates follow _cycle_terminal_event_dates exactly -- a PRESENCE test plus the
    first posting that actually carries a value, never a date parse, because these
    columns are as mixed-typed in real files as Liquidation Date is. Amounts are
    summed across postings.

    NOT inputs to any calculation. Nothing in the delinquency ladder,
    payment-sufficiency/BR-002, the reversal deficiency formula,
    expected_loan_status, consolidation, the Portfolio Summary, the Monthly
    Servicer Report or corrections.py's eligibility allow-list reads any of them --
    every existing consolidated value is byte-for-byte unchanged by their
    presence, and a file lacking these columns behaves exactly as before. They
    exist so the review page can SHOW the reviewer the evidence, and for nothing
    else. Never feed one back into a calculation here."""
    out = {}
    for column, key in _REVIEW_EVIDENCE_DATE_COLUMNS.items():
        display = ""
        if column in group.columns:
            for value in group[column]:
                if not _is_blank_value(value):
                    display = str(value).strip()
                    break
        out[key] = display
        out[f"{key}_populated"] = bool(display)
    for column, key in _REVIEW_EVIDENCE_AMOUNT_COLUMNS.items():
        if column in group.columns:
            total = pd.to_numeric(group[column], errors="coerce").sum(min_count=1)
            out[key] = None if pd.isna(total) else _round2(total)
        else:
            out[key] = None
    return out


def consolidate_cycle_postings(servicer_inputs_df):
    """Returns one consolidated cycle record per SBA GP Number:
      posting_count, loan_status, dq_status (both from the LAST chronological
      posting), received/servicing_fee/principal/net_interest (SUMMED),
      scheduled (a
      single value — see below), scheduled_ambiguous (bool),
      scheduled_payment_updated_during_cycle (bool), scheduled_distinct_values
      (list, diagnostic), beginning_balance (the CYCLE OPENING posting — see
      _cycle_opening_posting; NOT necessarily the sequence's first posting),
      ending_balance (LAST posting), interest_from (CYCLE OPENING),
      interest_to (LAST), lender_loan_number (FIRST).

    Scheduled Payment is never summed — there is one scheduled obligation
    per cycle. If every nonzero Scheduled Payment value for a loan this
    cycle agrees, that value is used. If two or more DISTINCT nonzero
    values exist, this first checks whether they form a clean, one-
    directional step pattern (see _resolve_scheduled_payment_step_pattern)
    — a legitimate mid-cycle update/correction, not genuine ambiguity. If
    so, `scheduled` is the final chronological value and
    scheduled_payment_updated_during_cycle is True. Otherwise `scheduled`
    is None and `scheduled_ambiguous` is True — callers must not silently
    sum, max, or pick one; the cycle obligation is genuinely ambiguous and
    should be surfaced, not guessed at.

    CLIENT-CONFIRMED (2026-08-06): a NARROW fallback for a genuinely
    incomplete final posting. Discovered on real Peoples Bank data (GP
    9031018509, June 2026) — a loan's chronologically LAST posting (the
    one flagged "Last Payment Instance? = Y") was completely blank across
    Loan Status, Delinquency Status, and Ending Balance, while four
    earlier, complete postings existed for the same cycle. Taking that
    blank row at face value silently zeroed out a real $650,128.27 balance
    and blanked its Loan Status/DQ Status, removing the loan from every
    downstream total with no visible sign anything had gone wrong.

    The fallback is deliberately narrow: it only activates when the LAST
    posting has NO authoritative signal at all (Loan Status, Delinquency
    Status, AND Ending Balance all blank/NaN) — never for a posting that
    legitimately reports a real $0 balance or a real terminal status. When
    it activates, loan_status/dq_status/ending_balance/_excel_row are all
    taken from the last posting that actually has real data instead
    (`effective_last`), and this is recorded, never silently: see
    incomplete_final_posting/original_last_excel_row/
    effective_last_excel_row below, and the separate "Final Posting
    Incomplete" check (final_posting_incomplete_check) that surfaces this
    for every affected loan.

    Beginning Balance is deliberately UNCHANGED by this fallback — it is
    sourced from the CYCLE OPENING posting (`opening`, below), never from
    `last`, so a blank last posting was never able to affect it in the
    first place. Flagged explicitly here since it's easy to assume
    otherwise.
    """
    si = servicer_inputs_df.copy()
    si["SBA GP Number"] = _norm_id(si["SBA GP Number"])
    si = si.reset_index(drop=True)  # preserve original row order as the final sort fallback

    numeric_cols = [c for c in ("Payment Received", "Servicing Fee", "Principal", "Net Interest",
                                 "Scheduled Payment", "Beginning Balance", "Ending Balance")
                    if c in si.columns]
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
        # CLIENT-CONFIRMED (2026-08-14): the row the cycle OPENED at, resolved
        # independently of the sequence above because the lender does not
        # guarantee the earliest-dated posting is written first. Read ONLY for
        # start-of-cycle values (Beginning Balance, Interest-From) below; `last`
        # and `effective_last` are unaffected and keep every ending-state value.
        opening = _cycle_opening_posting(group)

        # CLIENT-CONFIRMED (2026-08-06): narrow incomplete-final-posting
        # fallback -- see the function docstring above for the full
        # rationale and the real loan that surfaced this.
        final_state_blank = (
            _is_blank_value(last.get("Loan Status"))
            and _is_blank_value(last.get("Delinquency Status"))
            and pd.isna(last.get("Ending Balance"))
        )
        if final_state_blank:
            # CORRECTED (2026-08-06): the originally-proposed filter used
            # raw `.notna()` on each column, which does NOT catch an empty
            # string ('') the way _is_blank_value above does -- pandas
            # treats '' as "not NaN." On the real loan that surfaced this
            # whole issue, the blank final posting's Loan Status/
            # Delinquency Status cells were empty strings, not NaN, which
            # meant the original filter still counted that row as "valid"
            # and effective_last ended up identical to the blank `last`.
            # Rewritten to use the SAME _is_blank_value definition as the
            # detection above, so "valid" and "blank" can never disagree
            # with each other.
            valid_mask = (
                group["Ending Balance"].notna()
                | ~group["Loan Status"].apply(_is_blank_value)
                | ~group["Delinquency Status"].apply(_is_blank_value)
            )
            valid_rows = group[valid_mask]
            effective_last = valid_rows.iloc[-1] if not valid_rows.empty else last
            incomplete_final_posting = True
        else:
            effective_last = last
            incomplete_final_posting = False

        sched_series = group["Scheduled Payment"] if "Scheduled Payment" in group.columns else pd.Series(dtype=float)
        nonzero = sched_series[(sched_series.notna()) & (sched_series != 0)]
        distinct_nonzero = sorted(nonzero.unique().tolist())
        scheduled_payment_updated_during_cycle = False
        if len(distinct_nonzero) > 1:
            step_value = _resolve_scheduled_payment_step_pattern(group, sched_series)
            if step_value is not None:
                scheduled = step_value
                scheduled_ambiguous = False
                scheduled_payment_updated_during_cycle = True
            else:
                scheduled = None
                scheduled_ambiguous = True
        elif distinct_nonzero:
            scheduled = distinct_nonzero[0]
            scheduled_ambiguous = False
        elif sched_series.notna().any():
            scheduled = 0.0
            scheduled_ambiguous = False
        else:
            scheduled = None
            scheduled_ambiguous = False

        rows.append({
            "SBA GP Number": gp,
            "posting_count": len(group),
            "loan_status": effective_last.get("Loan Status", ""),
            "dq_status": effective_last.get("Delinquency Status", ""),
            "received": group["Payment Received"].sum(min_count=1) if "Payment Received" in group.columns else None,
            # Purely ADDITIVE raw fact (2026-08-12): did ANY posting in this
            # cycle carry a negative Payment Received? `received` above is the
            # NET across postings and cannot answer this — a +X/-X pair nets to
            # exactly $0 and is indistinguishable from a genuine no-payment
            # cycle by the net alone. Recorded here because this is the one
            # place that already sees every posting for the loan; deriving it
            # anywhere else would be the independent re-implementation this
            # shared consolidation exists to prevent.
            #
            # NOT an input to any calculation. Nothing in the delinquency
            # ladder, payment-sufficiency/BR-002, or the reversal deficiency
            # formula reads it — STAT-01's calculated value is byte-for-byte
            # unchanged by its presence. It exists so corrections.py can route
            # eligibility, which is that module's job, not this one's.
            "has_negative_posting": (
                bool((pd.to_numeric(group["Payment Received"], errors="coerce") < 0).any())
                if "Payment Received" in group.columns else False
            ),
            # CLIENT-CONFIRMED (2026-08-12). Purely ADDITIVE raw fact, same kind
            # and same reason as has_negative_posting above: did the servicer
            # actually FILE a posting for this cycle, as evidenced by the
            # template's own posting-identity fields?
            #
            #   Last Payment Instance? == "Y"  on any posting, OR
            #   Loan Payment Instance numerically populated on any posting
            #
            # Either-or, not both: a servicer may populate the ordinal without
            # the flag or the reverse, and both independently mean "this row is
            # a filed posting". Presence is `any posting in the cycle`, matching
            # has_negative_posting -- a reversal pair can carry the ordinal on
            # one row and leave it blank on another, and the loan does not stop
            # having been filed because one row omitted it.
            #
            # WHY IT IS A FACT AND NOT A JUDGEMENT. Both columns are literal
            # template data that this function ALREADY reads to order postings
            # and identify the cycle's last one (see the sort above). This
            # records what it already saw; it invents nothing. Deliberately NOT
            # derived from the "Confirm Principal + Interest = Payment Received"
            # column, which separates the same loans equally well but is an
            # Excel FORMULA (=H-G-F) -- python-calamine reads cached values
            # only, and preflight.py exists precisely because a formula's cache
            # can be stale, overwritten or absent. A business rule must not
            # depend on a servicer's helper formula.
            #
            # NOT an input to any calculation. Nothing in the delinquency
            # ladder, payment-sufficiency/BR-002, the reversal deficiency
            # formula, expected_loan_status, or consolidation itself reads it --
            # every consolidated value above is byte-for-byte unchanged by its
            # presence. Its one consumer is the Portfolio Summary "Not Making
            # Payment" rule (_build_portfolio_summary_record), which needs to
            # tell a filed no-payment posting from a row that was never filed.
            "has_filed_posting": (
                bool(
                    (group["Last Payment Instance?"].astype("string").str.strip().str.upper() == "Y").any()
                    if "Last Payment Instance?" in group.columns else False
                )
                or bool(
                    pd.to_numeric(group["Loan Payment Instance"], errors="coerce").notna().any()
                    if "Loan Payment Instance" in group.columns else False
                )
            ),
            # CLIENT-CONFIRMED (2026-08-12). Purely ADDITIVE raw facts, the same
            # kind and for the same reason as has_negative_posting above: the
            # current cycle's own terminal-event dates, recorded here because
            # this is the one place that already sees every posting for the loan.
            #
            # Carried as (bool populated, display string) pairs rather than as
            # parsed dates. The columns are genuinely mixed-typed in real files —
            # Bankruptcy Date arrives as a datetime.date while Liquidation Date
            # arrives as the string "06/29/2026" — so the client's rule ("the
            # column CONTAINS a date") is a PRESENCE test. Parsing would invent a
            # second, stricter rule and would silently drop an unparseable-but-
            # populated cell, which is the opposite of the intended reading.
            #
            # Presence is `any posting in the cycle`, matching has_negative_
            # posting: a reversal pair or an incomplete final posting can leave
            # the date on one row and blank on another, and the loan does not
            # stop having a bankruptcy date because the last row omitted it.
            # The display string comes from the FIRST posting that actually
            # carries a value, so the reviewer is shown a real date rather than
            # whichever row happened to be last.
            #
            # NOT inputs to any calculation. Nothing in the delinquency ladder,
            # payment-sufficiency/BR-002, the reversal deficiency formula, or
            # expected_loan_status reads them — STAT-01's calculated Loan Status
            # and Delinquency Status are byte-for-byte unchanged by their
            # presence. They exist so corrections.py can route eligibility,
            # which is that module's job, not this one's.
            **_cycle_terminal_event_dates(group),
            # CLIENT-CONFIRMED (2026-08-13): additional DISPLAY-ONLY review
            # evidence -- Foreclosure/REO Date, Liquidation Proceeds/Loss and
            # Subsequent Recoveries. Same additive posture as the block above and
            # deliberately kept as its own helper: nothing routes eligibility on
            # these. See _cycle_review_evidence.
            **_cycle_review_evidence(group),
            "servicing_fee": group["Servicing Fee"].sum(min_count=1) if "Servicing Fee" in group.columns else None,
            "principal": group["Principal"].sum(min_count=1) if "Principal" in group.columns else None,
            "net_interest": group["Net Interest"].sum(min_count=1) if "Net Interest" in group.columns else None,
            "scheduled": scheduled,
            "scheduled_ambiguous": scheduled_ambiguous,
            "scheduled_payment_updated_during_cycle": scheduled_payment_updated_during_cycle,
            "scheduled_distinct_values": distinct_nonzero,
            # START-OF-CYCLE values come from the cycle OPENING posting
            # (_cycle_opening_posting), not from the first row of the posting
            # sequence -- see that function and the doctrine comment above.
            "beginning_balance": opening.get("Beginning Balance"),
            "ending_balance": effective_last.get("Ending Balance"),
            "interest_from": opening.get("Interest-From", ""),
            "interest_to": last.get("Interest-To", ""),
            # Loan IDENTITY, not a start-of-cycle value: every posting for a
            # loan carries the same Lender Loan Number, so this deliberately
            # stays on the sequence's first row rather than moving with the
            # cycle opening.
            "lender_loan_number": first.get("Lender Loan Number", ""),
            # Source-location metadata for the correction engine (backend/
            # corrections.py) only — these come from the SAME chronologically
            # LAST posting whose loan_status/dq_status are used above, since
            # that's the row a correction would actually need to overwrite.
            # Not used by any existing check or the audit report.
            "_excel_row": effective_last.get("_excel_row"),
            "_sheet_name": last.get("_sheet_name"),
            "_header_excel_row": last.get("_header_excel_row"),
            # CLIENT-CONFIRMED (2026-08-05): every posting row's Excel row
            # number for this loan/cycle, in the same chronological order
            # established above -- NOT just the last posting's row. Purely
            # additive location metadata for corrections.py so an approved
            # Loan Status/Delinquency Status correction can be written to
            # every posting row for the loan, not only the row `last`
            # happens to point at. Does not affect received/scheduled/
            # loan_status/dq_status or any other consolidated value above;
            # sheet_name and header_excel_row are unchanged (single file,
            # single loan, single sheet/header row for all its postings).
            "_all_excel_rows": group["_excel_row"].tolist() if "_excel_row" in group.columns else [],
            # CLIENT-CONFIRMED (2026-08-06): disclosure fields for the
            # incomplete-final-posting fallback above. incomplete_final_posting
            # is False for the overwhelming majority of loans (nothing
            # changed); when True, the two _excel_row fields let a reviewer
            # (and final_posting_incomplete_check, below) point at exactly
            # which row was blank and which row's values were actually used.
            "incomplete_final_posting": incomplete_final_posting,
            "original_last_excel_row": last.get("_excel_row"),
            "effective_last_excel_row": effective_last.get("_excel_row"),
        })

    return pd.DataFrame(rows).set_index("SBA GP Number") if rows else pd.DataFrame(
        columns=["posting_count", "loan_status", "dq_status", "received", "has_negative_posting",
                 "has_filed_posting",
                 "bankruptcy_date", "bankruptcy_date_populated",
                 "paid_in_full_date", "paid_in_full_date_populated",
                 "liquidation_date", "liquidation_date_populated",
                 "foreclosure_date", "foreclosure_date_populated",
                 "reo_date", "reo_date_populated",
                 "liquidation_proceeds", "liquidation_loss", "subsequent_recoveries",
                 "servicing_fee", "principal", "net_interest",
                 "scheduled", "scheduled_ambiguous", "scheduled_payment_updated_during_cycle",
                 "scheduled_distinct_values", "beginning_balance",
                 "ending_balance", "interest_from", "interest_to", "lender_loan_number",
                 "_excel_row", "_sheet_name", "_header_excel_row", "_all_excel_rows",
                 "incomplete_final_posting", "original_last_excel_row", "effective_last_excel_row"]
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
# CONT-01 — Prior Period Loan Continuity
# CLIENT-CONFIRMED (2026-08-11).
#
# Answers exactly ONE question, deliberately narrow:
#
#   A loan carried an Ending Actual Balance above BALANCE_TOLERANCE in the
#   IMMEDIATELY PRIOR finalized reporting month, but does not appear at all
#   in this cycle's Servicer Inputs.
#
# SOUP is the authoritative prior finalized period. _infer_cycle_month already
# codifies the architectural fact this depends on: SOUP only has data through
# the PRIOR cycle, so SOUP's latest Payment Month IS last month's finalized
# report. Validated on real data -- SOUP 2026-05 Ending Actual Balance matched
# the June 2026 SRF's own "Loan Terms & Prior Period" closing balance for all
# 122 loans in common, with no loan on either side only.
#
# NOT a balance-comparison check. When the loan IS present, whether its
# Beginning Balance ties to the prior Ending Balance is BAL-01's question and
# BAL-01's alone -- re-deriving that here would create a second implementation
# of a live rule that could drift from the first, exactly what
# consolidate_cycle_postings exists to prevent. CONT-01 covers the gap (loan
# vanished); BAL-01 covers the overlap (loan present, balance doesn't tie).
# Together they are the whole continuity rule with nothing duplicated.
#
# Also NOT POP-01. POP-01 compares the servicer's Loan Terms population
# against SOUP's; both sources can agree a loan exists while the loan is
# simply absent from the postings that report on it. Validated on the real
# July 2026 Lendistry 2024 fixture: GP 5561999102 (BANKRUPTCY, prior balance
# $12,675.95) sits in BOTH the Loan Terms sheet and SOUP, so POP-01 passes it
# correctly, while it is missing from Servicer Inputs entirely. POP-01 is not
# modified by this check and reuse of it was investigated and rejected.
#
# PRIOR-MONTH ANCHORING. The prior population is the rows at SOUP's LATEST
# Payment Month for this servicer/deal -- not the latest row per loan. A loan
# whose SOUP history ended months ago is PAY-01's "stopped reporting" case,
# and anchoring per-loan would re-flag it here every cycle forever. Anchoring
# also removed the only two false positives seen across real June data: two
# PAID-IN-FULL loans whose stale final rows carry floating-point balance
# residue (Lendistry 5214559105 at 2.04e-12, Poppy 6625075009 at 0.0007).
#
# ZERO IS TOLERANCE-JUDGED, never a strict `> 0`, for that same residue --
# consistent with every other balance comparison in this module.
#
# LEGITIMATE EXITS, and only these:
#   - a prior balance at or below BALANCE_TOLERANCE (the loan already closed
#     to zero -- this alone accounted for every real dropout across all eight
#     June 2026 servicer files: 15 loans, all at zero, zero false positives)
#   - SUBSTITUTION-OUT, via the existing _POPULATION_EXIT_STATUS_KEYWORDS
#   - closed-at-zero suppression, via the shared loans_closed_at_zero set
# Deliberately NOT _TERMINAL_STATUS_KEYWORDS: a PAID-IN-FULL or LIQUIDATION
# status still carrying a real balance is precisely a loan that must keep
# reporting until it closes, so widening the exit list would suppress the
# finding this check exists to make. BANKRUPTCY, FORECLOSURE, REO and
# deferral are never exits -- the fixture case is a BANKRUPTCY loan with a
# live balance.
#
# No documented-approved-exception bypass in this version (client decision,
# 2026-08-11): there is no exception register in this codebase yet, and an
# unreviewed suppression channel is the one thing that could make this check
# silently miss a real gap. Revisit only if real cases demand it.
# ---------------------------------------------------------------------------

def prior_period_continuity_check(servicer_inputs_df, soup_servicer_df,
                                  loan_terms_df=None, servicer_name=None,
                                  deal=None, closed_at_zero=None):
    records = []
    suppressed_closed_gps = []

    def _empty():
        return {
            "id": "CONT-01",
            "name": "Prior Period Loan Continuity",
            "category": "SBA Reconciliation",
            "severity": "FAIL",
            "status": "PASS",
            "issues_found": 0,
            "details": "No issues found.",
            "records": [],
            "suppressed_closed_gps": [],
        }

    # Degrade to "nothing to report" rather than raising, the same way every
    # other check in this module handles a source it can't evaluate.
    if soup_servicer_df is None or soup_servicer_df.empty:
        return _empty()
    if not {"SBA GP Number", "Payment Month", "Ending Actual Balance"}.issubset(soup_servicer_df.columns):
        return _empty()
    if servicer_inputs_df is None or servicer_inputs_df.empty:
        # A submission with no readable postings at all is a file-level
        # failure the readers and preflight already surface. Reporting every
        # loan in the portfolio as "missing" would bury that, not reveal it.
        return _empty()
    if "SBA GP Number" not in servicer_inputs_df.columns:
        return _empty()

    # DEAL SCOPING, for the same reason POP-01 needs it: SOUP files every note
    # series a servicer has under one Servicer value, so comparing Lendistry
    # 2024 against unscoped Lendistry SOUP would drag in the 2025 deal's 1,241
    # loans and report almost all of them as missing. Falls back to
    # servicer-only, with an explicit note, when the deal can't be determined.
    soup_scoped = soup_servicer_df
    deal_scoped = False
    if deal and "Deal" in soup_servicer_df.columns:
        soup_scoped = soup_servicer_df[soup_servicer_df["Deal"].astype("string").str.strip() == deal]
        deal_scoped = True

    scope_note = (
        f", scoped to Deal '{deal}'" if deal_scoped
        else " (Deal could not be determined for this file -- compared by servicer only, not Servicer + Deal)"
    )

    if soup_scoped.empty:
        return _empty()

    soup = soup_scoped.copy()
    soup["_gp"] = _norm_id(soup["SBA GP Number"])
    soup["_month"] = pd.to_datetime(soup["Payment Month"], errors="coerce")
    soup = soup[soup["_gp"].notna()].dropna(subset=["_month"])
    if soup.empty:
        return _empty()

    prior_month = soup["_month"].max()
    prior = soup[soup["_month"] == prior_month]
    # One record per loan even if the prior month somehow carries more than
    # one row for the same GP within a single deal.
    prior = prior.groupby("_gp").tail(1)

    reported_this_cycle = set(_norm_id(servicer_inputs_df["SBA GP Number"]).dropna())

    # Corroborating reviewer evidence ONLY: the current submission's own
    # "Loan Terms & Prior Period" sheet states the prior closing balance it
    # then failed to report on. Read-only display, never an input to the
    # test -- the same convention correction_data's review-evidence fields
    # follow. SOUP alone decides whether a loan is flagged, because a
    # servicer that drops a loan from Servicer Inputs could equally drop it
    # from its own prior-period sheet.
    lt_prior_balances = {}
    if (loan_terms_df is not None and not loan_terms_df.empty
            and {"SBA GP Number", "Ending Balance"}.issubset(loan_terms_df.columns)):
        lt = loan_terms_df.copy()
        lt["_gp"] = _norm_id(lt["SBA GP Number"])
        lt["_bal"] = pd.to_numeric(lt["Ending Balance"], errors="coerce")
        for lt_gp, lt_bal in zip(lt["_gp"], lt["_bal"]):
            if pd.notna(lt_gp) and pd.notna(lt_bal):
                lt_prior_balances[str(lt_gp).strip()] = round(float(lt_bal), 2)

    deal_label = deal or "unknown"
    prior_month_label = prior_month.strftime("%m/%Y")

    for _, row in prior.iterrows():
        gp = str(row["_gp"]).strip()

        if gp in reported_this_cycle:
            continue  # reported -- any balance break is BAL-01's question

        prior_balance = pd.to_numeric(row.get("Ending Actual Balance"), errors="coerce")
        if pd.isna(prior_balance):
            continue  # can't evaluate the rule without a prior balance
        if prior_balance <= BALANCE_TOLERANCE:
            continue  # already closed to zero -- a legitimate exit

        prior_status = str(row.get("Loan Status", "") or "").strip()
        if any(k in prior_status.upper() for k in _POPULATION_EXIT_STATUS_KEYWORDS):
            continue  # substituted out of the deal -- expected to stop reporting

        # CLOSED-AT-ZERO SUPPRESSION (see _CLOSED_AT_ZERO_STATUS_KEYWORDS),
        # applied LAST rather than first as BAL-01/BAL-02 do. Those iterate
        # the current cycle's population, so checking up front bounds their
        # trace; this iterates the PRIOR population, where an up-front check
        # would fill the trace with closed loans that were never going to be
        # flagged anyway. Applied here, a trace entry means what it says: this
        # loan would have been a CONT-01 finding and was withheld.
        if closed_at_zero and gp in closed_at_zero:
            suppressed_closed_gps.append(gp)
            continue

        prior_balance = round(float(prior_balance), 2)
        records.append({
            "SBA GP Number": gp,
            "Deal": deal_label,
            "Prior Reporting Month": prior_month_label,
            "Prior Loan Status": prior_status,
            "Prior Delinquency Status": str(row.get("Delinquency Status", "") or "").strip(),
            "Prior Ending Balance": prior_balance,
            "Reported This Cycle": "NO",
            "Issue": (
                f"Loan carried a prior Ending Balance of ${prior_balance:,.2f} and "
                f"did not appear in this cycle's Servicer Inputs."
            ),
            "Prior Period Sheet Balance": lt_prior_balances.get(gp, ""),
        })

    records.sort(key=lambda r: r["SBA GP Number"])

    details = "No issues found." if not records else (
        f"{len(records)} loan(s) carried a prior Ending Balance above zero but were not "
        f"reported in this cycle for {servicer_name}{scope_note}."
    )

    return {
        "id": "CONT-01",
        "name": "Prior Period Loan Continuity",
        "category": "SBA Reconciliation",
        "severity": "FAIL",
        "status": "PASS" if not records else "FAIL",
        "issues_found": len(records),
        "details": details,
        "records": records,
        "suppressed_closed_gps": suppressed_closed_gps,
    }


# ---------------------------------------------------------------------------
# BAL-01 — Balance Continuity
# This cycle's Beginning Balance (Servicer Inputs) should tie to SOUP's most
# recent Ending Actual Balance for that loan. Loans with multiple payment
# postings this cycle are flagged as WARN (informational) rather than FAIL,
# since their balance naturally reflects more than one payment.
# ---------------------------------------------------------------------------

def balance_continuity_check(servicer_inputs_df, soup_servicer_df, closed_at_zero=None):
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

    suppressed_closed_gps = []
    for gp, row in si_consolidated.iterrows():
        # CLOSED-AT-ZERO SUPPRESSION (see _CLOSED_AT_ZERO_STATUS_KEYWORDS).
        # Prior SOUP already establishes this loan as closed at a zero
        # balance, so it is not eligible to be a client-facing exception
        # whatever the servicer now reports.
        if closed_at_zero and str(gp).strip() in closed_at_zero:
            suppressed_closed_gps.append(str(gp).strip())
            continue
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
        "suppressed_closed_gps": suppressed_closed_gps,
    }


# ---------------------------------------------------------------------------
# BAL-02 — Liquidation Balance Validation
# The Remittance Form Instructions formally define LIQUIDATION as: the loan
# has been charged off, the balance is reduced to $0, and the loan is no
# longer expected to make regular payments. This is a CURRENT-CYCLE rule —
# a loan reporting Loan Status = LIQUIDATION this cycle should report an
# Ending Balance of $0 this cycle, full stop. Deliberately independent of
# BAL-01: this doesn't look at prior SOUP balance at all, so a loan can fail
# BAL-01 (continuity break vs. last cycle) and BAL-02 (current-cycle
# LIQUIDATION balance rule) simultaneously, for two genuinely separate
# reasons, exactly as validated against loan 4338609104.
# ---------------------------------------------------------------------------

_LIQUIDATION_STATUS_KEYWORD = "LIQUIDATION"


def liquidation_balance_check(servicer_inputs_df, closed_at_zero=None):
    records = []

    required = {"SBA GP Number", "Loan Status", "Ending Balance"}
    if not required.issubset(servicer_inputs_df.columns):
        return _result("BAL-02", "Liquidation Balance Validation", "SBA Reconciliation", "FAIL", records,
                        lambda n: "")

    # Evaluate the CONSOLIDATED cycle record per GP — final chronological
    # Loan Status and final chronological Ending Balance — not each raw
    # posting independently. A loan with multiple postings this cycle has
    # one true end-of-cycle state; checking intermediate postings on their
    # own could flag (or miss) based on a mid-cycle row that isn't what the
    # loan actually ended the cycle as.
    consolidated = consolidate_cycle_postings(servicer_inputs_df)

    suppressed_closed_gps = []
    for gp, row in consolidated.iterrows():
        # CLOSED-AT-ZERO SUPPRESSION (see _CLOSED_AT_ZERO_STATUS_KEYWORDS).
        # Prior SOUP already establishes this loan as closed at a zero
        # balance, so it is not eligible to be a client-facing exception
        # whatever the servicer now reports.
        if closed_at_zero and str(gp).strip() in closed_at_zero:
            suppressed_closed_gps.append(str(gp).strip())
            continue
        loan_status = str(row.get("loan_status", "")).strip().upper()
        if loan_status != _LIQUIDATION_STATUS_KEYWORD:
            continue

        ending_balance = row.get("ending_balance")
        if ending_balance is None or pd.isna(ending_balance):
            continue  # can't evaluate without a balance figure
        if abs(ending_balance) <= BALANCE_TOLERANCE:
            continue  # within tolerance of $0 -- PASS for this loan

        records.append({
            "SBA GP Number": gp,
            "Loan Status": str(row.get("loan_status", "")).strip(),
            "Ending Balance": round(float(ending_balance), 2),
            "Expected Ending Balance": 0.00,
            "Difference": round(float(ending_balance), 2),
            "Issue": "LIQUIDATION loans are required to report a $0 balance.",
        })

    details = "No issues found." if not records else (
        f"{len(records)} loan(s) reported as LIQUIDATION with a nonzero Ending Balance, "
        f"contrary to the Remittance Form Instructions."
    )

    return {
        "id": "BAL-02",
        "name": "Liquidation Balance Validation",
        "category": "SBA Reconciliation",
        "severity": "FAIL",
        "status": "PASS" if not records else "FAIL",
        "issues_found": len(records),
        "details": details,
        "records": records,
        "suppressed_closed_gps": suppressed_closed_gps,
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
#   received < 0 (client-confirmed 2026-08-04, second revision -- a
#   reversal removing payment value, which also inherently means the
#   current cycle's own scheduled payment went unpaid; DOES use the same
#   80% threshold as every other branch below, as the divisor)
#                                           -> missed_count + max(
#                                                  1,
#                                                  floor(
#                                                      (
#                                                          scheduled_payment
#                                                          + abs(received)
#                                                      )
#                                                      / (
#                                                          scheduled_payment
#                                                          * PAYMENT_SUFFICIENCY_THRESHOLD
#                                                      )
#                                                  )
#                                              )
#                                              (UNVERIFIED if Scheduled
#                                              Payment isn't a valid positive
#                                              number, or if the prior SOUP
#                                              status can't be parsed into a
#                                              missed count)
#   received == 0                          -> missed_count + 1 (always a
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

def _is_offsetting_zero_net_reversal(consolidated_row):
    """CLIENT-CONFIRMED (2026-08-12). True when a loan's cycle contains at
    least one NEGATIVE Payment Received posting and those postings offset to a
    consolidated net of exactly $0.00.

    This is the blind spot in the net-based reversal test. `reversal_equivalents`
    is only ever set when the CONSOLIDATED payment is negative, so a reversal
    that is exactly cancelled by an offsetting positive posting reads as an
    ordinary no-payment cycle: same net ($0), same ladder result (+1 level),
    no reversal marker. Confirmed on real June 2026 data — 33 recommendations,
    every one a clean +X/-X pair (e.g. Lendistry 5519029100 [201.47, -201.47],
    6010139100 [1263.51, -1263.51, 1214.03, -1214.03]).

    The existence of an offsetting positive posting does not make a reversal
    stop being a reversal, so these must be judged by a human exactly as a
    net-negative reversal is.

    A RAW FACT, not an eligibility judgment — corrections.py decides what to do
    with it. Nothing here or anywhere else in this module's calculation logic
    reads it back: STAT-01's calculated Delinquency Status is unchanged.

    "Exactly $0" is judged at cent precision (the precision every figure in
    this application is reported and compared at), so a pair that leaves
    floating-point residue behind still qualifies."""
    if not bool(consolidated_row.get("has_negative_posting")):
        return False
    received = pd.to_numeric(consolidated_row.get("received"), errors="coerce")
    if pd.isna(received):
        return False
    return round(float(received), 2) == 0.00


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
# needed on this value. Every payment-sufficiency comparison in this file
# — including the reversal-driven calculation for a negative Payment
# Received amount (client-confirmed 2026-08-04, second revision) — MUST
# reference this constant, never a hardcoded literal, so any future change
# to the confirmed value only needs to happen in this one place.
PAYMENT_SUFFICIENCY_THRESHOLD = 0.80


def _apply_payment_transition(prior_missed_count, payment_received, scheduled_payment,
                              servicing_fee=None):
    """Returns (new_missed_count, payment_equivalents_removed).

    ADJUSTED PAYMENT TOTAL — the sufficiency numerator is
    Payment Received + Servicing Fee, summed across every posting for the
    loan this cycle, matching the SRF template's own column AD
    (=SUMIF(Payment Received) + SUMIF(Servicing Fee)). Payment Received is
    reported NET of the servicing fee — see the PAY-02 note on the fee
    being "netted out before remittance" — so judging sufficiency on the
    net figure measures what the servicer remitted rather than what the
    borrower paid. Adding the fee back restores the gross amount, which is
    what the 80% bar is meant to test.

    servicing_fee is optional and treated as zero when absent: not every
    servicer template carries the column (same caveat PAY-02 records).
    A servicer reporting no fee therefore behaves exactly as before.

    NOT YET CLIENT-CONFIRMED. The 0.80 threshold is confirmed
    (see PAYMENT_SUFFICIENCY_THRESHOLD); the composition of the numerator
    is not. Measured impact on the June and July 2026 Peoples Bank cycles:
    zero loans change count. Confirmation should be obtained before this is
    relied upon for a cycle where it does bite.

    The fee is added ONLY to the positive-payment sufficiency judgement.
    The negative/reversal branch below continues to key off the raw
    Payment Received amount and its client-confirmed (2026-08-04, second
    revision) deficiency formula, which is deliberately untouched here: a
    small negative payment plus a fee could otherwise sum to a positive
    figure and silently bypass the reversal path altogether.

    new_missed_count is None when the transition genuinely cannot be
    computed — callers MUST treat None as "this cycle's result is
    UNVERIFIED", never as zero or as "hold steady".

    payment_equivalents_removed is None on every path except a NEGATIVE
    Payment Received amount, where it's the integer count of scheduled-
    payment equivalents that amount represents. Returned alongside the
    missed count so callers (NEG-01 in particular) can describe a
    reversal's impact without re-deriving this math independently, which
    is exactly the kind of duplicated calculation this file was already
    burned by once (see the module note on _calculate_expected_status).

    A period is only treated as "paid" if the amount received covers at
    least PAYMENT_SUFFICIENCY_THRESHOLD (80%) of the scheduled payment for
    that installment — receiving *some* payment isn't enough on its own.
    A payment below that threshold moves the loan down one delinquency
    rank, exactly the same as receiving nothing at all. Payments covering
    multiple installments' worth (e.g. a catch-up payment) still reduce
    the missed count by the number of full installments covered, each
    judged against the same 80% threshold. This paragraph describes the
    POSITIVE-payment branches only — see below for the separate deficiency-
    based formula that applies when payment_received is negative (which
    also uses this same 80% threshold, just structured differently).

    CLIENT-CONFIRMED (2026-08-04, second revision — supersedes the same-day
    100%-divisor version of this calculation, which itself had superseded
    the original 2026-08-03 version): a NEGATIVE Payment Received amount
    represents payment value being actively removed from the loan, not
    simply "no payment this cycle" — its magnitude matters, AND it
    inherently also means the current cycle's own scheduled payment went
    unpaid (a negative net total is, by definition, less than the
    scheduled amount). Both pieces are combined into one total deficiency,
    then measured in units of the SAME 80% sufficiency threshold used by
    every other judgment in this function (positive-payment sufficiency,
    the catch-up installment count above):

        total_payment_deficiency = scheduled_payment + abs(negative_amount)
        payment_equivalents_removed = max(1, floor(
            total_payment_deficiency / (scheduled_payment * PAYMENT_SUFFICIENCY_THRESHOLD)
        ))

    This DOES apply PAYMENT_SUFFICIENCY_THRESHOLD (80%) as the divisor —
    the reversal calculation is now consistent with the rest of this
    function rather than a special-cased 100% division. Verified against
    loan 5521259109: $212.37 scheduled, -$202.15 received -> deficiency
    $414.52 -> one threshold unit = $212.37 x 80% = $169.90 -> $414.52 /
    $169.90 = 2.44 -> 2 levels -> 30 DAYS from a CURRENT baseline.

    Minimum one level applies regardless, since ANY negative net payment
    means no valid payment was received this cycle even if the total
    deficiency is below one full threshold unit. The ladder then advances
    that many levels from whatever rank the loan carried in
    (prior_missed_count), NOT from a fixed baseline — this is deliberately
    the same "+N from wherever it already was" pattern used by every other
    branch in this function, not an attempt to identify which specific
    prior payment was reversed."""
    if payment_received is None or pd.isna(payment_received):
        payment_received = 0
    received = float(payment_received)

    # Adjusted Payment Total, used for POSITIVE-payment sufficiency only.
    fee = 0.0
    if servicing_fee is not None and not pd.isna(servicing_fee):
        fee = float(servicing_fee)

    if received < 0:
        if (
            scheduled_payment is None
            or pd.isna(scheduled_payment)
            or float(scheduled_payment) <= 0
        ):
            # Can't convert a negative amount into payment equivalents
            # without a valid, POSITIVE Scheduled Payment to measure it
            # against. A negative Scheduled Payment isn't just "missing" --
            # dividing by it would flip the sign of the ratio, and max(1, ...)
            # below would then silently mask that invalid input by always
            # reporting at least one equivalent regardless of what the
            # (wrong) math actually produced. Every caller must surface this
            # as UNVERIFIED, never guess.
            return None, None
        total_payment_deficiency = float(scheduled_payment) + abs(received)
        threshold_unit = float(scheduled_payment) * PAYMENT_SUFFICIENCY_THRESHOLD
        payment_equivalents_removed = max(
            1,
            math.floor(total_payment_deficiency / threshold_unit),
        )
        return prior_missed_count + payment_equivalents_removed, payment_equivalents_removed

    # a $0 payment is always a miss, regardless of what the Scheduled
    # Payment field says — some servicer submissions zero out BOTH fields
    # together for a no-payment month, which shouldn't be mistaken for
    # "no scheduled amount exists" and silently skipped
    if received == 0:
        return prior_missed_count + 1, None

    if scheduled_payment is None or pd.isna(scheduled_payment) or scheduled_payment == 0:
        # can't compute a ratio without a valid scheduled amount to divide
        # by — hold steady rather than risk a divide-by-zero
        return prior_missed_count, None

    # Adjusted Payment Total: Payment Received + Servicing Fee, both already
    # summed across every posting for this loan by consolidate_cycle_postings.
    adjusted_received = received + fee
    ratio = adjusted_received / float(scheduled_payment)

    if ratio < PAYMENT_SUFFICIENCY_THRESHOLD:
        # received something, but less than 80% of one scheduled payment --
        # treated as insufficient, same as a full miss
        return prior_missed_count + 1, None
    if ratio < 1.0:
        # received at least 80% but less than 100% of one installment --
        # counts as that installment paid (held steady, not reduced)
        return prior_missed_count, None
    # ratio >= 1.0: this is a catch-up payment covering one or more
    # installments.
    #
    # CLIENT-CONFIRMED (2026-08-04 — supersedes the prior round(ratio)
    # approach): installment counting here must use the SAME 80%
    # sufficiency threshold as every other judgment in this function, not
    # a raw round() of the ratio. round() could credit a second installment
    # with as little as 50% of it actually paid (anything from 1.5x to
    # 2.49x scheduled rounds to "2 installments"), which is inconsistent
    # with the 80% bar applied everywhere else (single-period sufficiency,
    # the reversal deficiency formula). Caught empirically on loan
    # 5571508408: a $635.00 payment against a $412.20 scheduled payment is
    # 1.54x — round(1.54) = 2, incorrectly crediting a second installment
    # off only 54% of it ($222.80 of the $329.76 needed), which reduced a
    # DELINQUENT loan to CURRENT with no real second installment paid.
    #
    # This instead walks the payment one installment at a time: the current
    # cycle's own installment is satisfied first (guaranteed here, since
    # ratio >= 1.0 already clears the 80% bar on its own), and each
    # ADDITIONAL installment only counts if the amount still remaining
    # after consuming the previous one clears 80% of one scheduled payment.
    # The full scheduled amount is consumed per credited installment (not
    # just the 80% portion), consistent with how an 80-100% single-period
    # payment is already treated as "that installment fully satisfied,"
    # never partially credited.
    installments_covered = 0
    remaining = adjusted_received
    threshold_amount = float(scheduled_payment) * PAYMENT_SUFFICIENCY_THRESHOLD
    while remaining >= threshold_amount:
        installments_covered += 1
        remaining -= float(scheduled_payment)
    if installments_covered <= 1:
        return prior_missed_count, None
    return max(0, prior_missed_count - (installments_covered - 1)), None


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


# CLIENT-CONFIRMED (2026-08-04): Securitization Action = "CHANGE IN TERMS
# AGREEMENT" may confirm that an EXISTING prior deferral is resuming this
# cycle — it is NEVER treated as evidence of a NEW deferral (that still
# requires Securitization Action = "Deferred", unchanged). Deliberately
# does NOT generalize to MODIFICATION, CHANGE MEMO, or any other
# Securitization Action value — real memo data shows those categories mix
# genuine temporary deferrals with permanent restructurings (e.g. "Extend
# the maturity date... Reduce the fixed interest rate...") with no reliable
# way to distinguish them from this one field alone. Widening this beyond
# CHANGE IN TERMS AGREEMENT requires a separate, explicitly confirmed rule.
#
# Text fields like "Other" are deliberately NOT used as evidence here —
# re-amortization language is consistent with a deferral having occurred,
# but the same language can also describe a permanent restructuring, so it
# doesn't independently prove anything. Instead, every condition below is
# either a structural fact (dates, their ordering) or a cross-check against
# SOUP's own independently-recorded history (condition 2) — nothing here
# depends on interpreting free-text.
_CHANGE_IN_TERMS_TERMINAL_DISQUALIFIERS = ("CHARGE", "LIQUIDAT", "BANKRUPT", "FORECLOS", "REO")


def _check_change_in_terms_resume(candidates, soup_status, cycle_month):
    """Implements all six client-specified conditions for the CHANGE IN
    TERMS AGREEMENT resume-only exception. Returns None if the exception
    doesn't apply (caller falls through to the standard Deferred-only
    logic), otherwise a deferral_info dict with phase="resume" and a
    diagnostic marker (resume_via_change_in_terms_exception=True) so the
    result stays visible for review rather than looking identical to a
    normal, standard-path resume.

    Returns (deferral_info, matched_row), both None when the exception does
    not apply. The row is returned so the caller can build the loan's deferral
    evidence from the same record this decision was made on."""
    if candidates is None or candidates.empty or cycle_month is None:
        return None, None
    if "Securitization Action" not in candidates.columns:
        return None, None

    # Condition 1: Securitization Action is exactly "CHANGE IN TERMS AGREEMENT".
    cit_rows = candidates[
        candidates["Securitization Action"].astype("string").str.strip().str.lower() == "change in terms agreement"
    ]
    if cit_rows.empty:
        return None, None

    # Condition 2: Prior SOUP Loan Status for this loan must be DEFERRED —
    # this is what makes the memo evidence of a RESUME, not a new deferral.
    # Checked against SOUP's own independent history, not the memo itself.
    if "DEFERRED" not in str(soup_status or "").upper():
        return None, None

    cycle_period = cycle_month.to_period("M").to_timestamp()

    for _, row in cit_rows.iterrows():
        # Condition 6: explicit, redundant safety check against a terminal
        # event — even though condition 1's exact match already structurally
        # excludes "Charge Off"/"Liquidation"/etc. as the ACTION value
        # itself, this guards against a future variant spelling or a
        # terminal event described only in Servicer Action.
        combined_text = " ".join(
            str(row.get(col, "")) for col in ("Securitization Action", "Servicer Action") if col in row.index
        ).upper()
        if any(k in combined_text for k in _CHANGE_IN_TERMS_TERMINAL_DISQUALIFIERS):
            continue

        # Condition 3: Regular Payment Resume Date present.
        resume_date = row.get("Regular Payment Resume Date")
        if pd.isna(resume_date):
            continue

        # Condition 5: Deferral Start/End Date present and chronologically valid.
        start_date = row.get("Deferral Start Date")
        end_date = row.get("Deferral End Date")
        if pd.isna(start_date) or pd.isna(end_date):
            continue
        start_ts, end_ts, resume_ts = pd.Timestamp(start_date), pd.Timestamp(end_date), pd.Timestamp(resume_date)
        if not (start_ts <= end_ts <= resume_ts):
            continue

        # Condition 4: the current audit cycle matches the Resume Date's month.
        resume_month = resume_ts.to_period("M").to_timestamp()
        if cycle_period != resume_month:
            continue

        # All six conditions satisfied — confirmed resume, never a new deferral.
        # The matched row rides along so the caller can shape its evidence
        # record from the SAME row this decision was made on, rather than
        # re-finding it and risking a different one.
        return {
            "deferral_type": "full",
            "phase": "resume",
            "months_since_start": None,
            "resume_via_change_in_terms_exception": True,
        }, row

    return None, None


# Deferral types that take the FULL P&I calculation path below.
#
# "unclassified" is here on purpose, and what it means must be stated
# precisely: this is a PROVISIONAL LENIENT ASSUMPTION, not neutral arithmetic.
# See docs/deferralevidencebrief.md, "Interim treatment for unclassified
# deferrals (governing decision, 2026-08-07)", which governs this comment.
#
# file_readers.classify_deferral_type returns "unclassified" for Servicer
# Action text with no payment-type signal at all ("three months deferment").
# Full P&I freezes the delinquency rank, because no payment was due; partial
# runs normal missed-payment math, because interest was still due. With the
# type unknown, keeping the freeze LEANS LENIENT — exactly as recomputing on
# the no-deferral path would lean strict. Neither is objectively correct until
# the type is known, so the rule preserves the established number, declares it
# provisional, and routes the resolution to a human. Describing this as "no
# number moves" is true of the diff and misleading about the claim: the
# retained figure is an assumption, and every unclassified delinquency value
# is provisional pending confirmation of the deferral type.
#
# Why preserve rather than recompute: (1) it does not introduce a new
# financial result merely because the parser got more honest; (2)
# "unclassified" plus DEFERRAL DATA QUALITY WARNING stop the system
# presenting the assumption as fact; (3) exclusion from auto-correction stops
# the lenient assumption quietly becoming a written status change. Measured on
# the June-2026 run, recomputing instead moved $280,162.64 from Current into
# 60 Days on one Millennium loan.
#
# corrections.py deliberately does NOT use this set — its Loan Status ->
# DEFERRED gate stays a strict `== "full"`, so an unclassified deferral can
# never trigger an automatic status write even if its warning were ever
# dropped from WARNING_REASONS.
_FULL_PNI_CALC_TYPES = ("full", "unclassified")

# ---------------------------------------------------------------------------
# THE DEFERRAL EVIDENCE CONTRACT
#
# One vocabulary describing what the audit actually knows about a loan's
# deferral evidence this cycle. Computed ONCE, here in STAT-01, and consumed
# verbatim by corrections.py, report.py and the correction review page. None of
# those three may recompute or reinterpret it — that is the whole point of
# naming it once, since three renderers inventing three readings of the same
# facts is how they drift.
#
#   documented                      a deferral applies this cycle, its dates
#                                   are internally consistent and its type is
#                                   known
#   no-deferral                     no memo deferral applies to this cycle.
#                                   A FACT, not a defect
#   deferral-claimed-dates-missing  the memo asserts a deferral but wrote no
#                                   usable Deferral Start Date, so the window
#                                   cannot be placed on any cycle
#   lookup-failed                   memo rows exist for this loan but none
#                                   matched a known Lender identity for this
#                                   servicer — a DEFECT that must never render
#                                   as "no deferral"
#   date-anomaly                    the matched row's dates are impossible
#                                   (end or resume before start)
#   type-unclassified               a deferral applies, but the Servicer
#                                   Action text never says full P&I vs partial
#
# NB "no-deferral" and "lookup-failed" are the two that look identical in the
# output today and are not the same thing at all — separating them is the
# reason this field exists.
EVIDENCE_DOCUMENTED = "documented"
EVIDENCE_NO_DEFERRAL = "no-deferral"
EVIDENCE_DATES_MISSING = "deferral-claimed-dates-missing"
EVIDENCE_LOOKUP_FAILED = "lookup-failed"
EVIDENCE_DATE_ANOMALY = "date-anomaly"
EVIDENCE_TYPE_UNCLASSIFIED = "type-unclassified"

DEFERRAL_EVIDENCE_STATUSES = (
    EVIDENCE_DOCUMENTED, EVIDENCE_NO_DEFERRAL, EVIDENCE_DATES_MISSING,
    EVIDENCE_LOOKUP_FAILED, EVIDENCE_DATE_ANOMALY, EVIDENCE_TYPE_UNCLASSIFIED,
)


def _memo_date_order_is_impossible(row):
    """End or resume before start. Deliberately the SAME two rules
    memo_audit's date-anomaly flag uses, so the per-loan label in STAT-01 and
    the file-level count in MEMO-01 can never disagree about the same row.
    ("resume before end" is invalid too, but it is advisory in memo_audit and
    is not part of that flag — keep the two definitions in step.)

    LABEL ONLY. Nothing branches on this: it does not gate the row, change a
    phase, or alter a delinquency value. Applying a date-anomaly gate to the
    calculation is a separate, client-confirmed decision."""
    if row is None:
        return False
    start = row.get("Deferral Start Date")
    end = row.get("Deferral End Date")
    resume = row.get("Regular Payment Resume Date")
    if pd.isna(start):
        return False
    start = pd.Timestamp(start)
    if pd.notna(end) and pd.Timestamp(end) < start:
        return True
    if pd.notna(resume) and pd.Timestamp(resume) < start:
        return True
    return False


def _deferral_evidence(status, row=None):
    """Shapes the evidence record every renderer reads. Always returns the
    same keys so a consumer never has to test for their presence, and always
    returns plain display-ready strings (never NaT/Timestamp), because these
    cross into JSON, Excel and the review page unchanged.

    READ-ONLY REVIEW EVIDENCE, per the rule in CLAUDE.md: nothing in this
    module's calculation logic may read these back."""
    def _date(col):
        v = row.get(col) if row is not None else None
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            return ""
        try:
            return pd.Timestamp(v).strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            return str(v).strip()

    action = row.get("Securitization Action") if row is not None else None
    return {
        "evidence_status": status,
        "deferral_start": _date("Deferral Start Date"),
        "deferral_end": _date("Deferral End Date"),
        "deferral_resume": _date("Regular Payment Resume Date"),
        "securitization_action": "" if action is None or (not isinstance(action, str) and pd.isna(action)) else str(action).strip(),
    }


_UNCLASSIFIED_TYPE_WARNING = (
    "The Servicer Memo's Servicer Action text does not state whether this is a "
    "full principal-and-interest deferral or a principal-only/interest-only "
    "one, so the deferral type could not be established from the memo. This "
    "cycle's Delinquency Status is therefore PROVISIONAL: it retains the "
    "established full P&I treatment (the rank is held, as it would be if no "
    "payment were due), which is a deliberate interim assumption rather than a "
    "confirmed result — a principal-only deferral would instead be evaluated "
    "against normal missed-payment math, because interest would still be due. "
    "The loan is excluded from automatic correction until a human confirms the "
    "deferral type, at which point the definitive status follows the correct "
    "rule."
)


def _lookup_deferral(memo_df, gp, servicer_key, cycle_month, soup_status=None):
    """Finds the memo row (if any) whose deferral window covers cycle_month
    for this loan, and returns
    (deferral_info, warning, warning_type, evidence):

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
      evidence: ALWAYS a dict (never None), from _deferral_evidence — the
               loan's evidence_status plus the matched memo row's raw
               Deferral Start/End/Resume dates and Securitization Action.
               Read-only review evidence: it is display material for
               corrections.py, report.py and the review page, and no
               calculation in this module may read it back.

    soup_status is SOUP's last recorded Loan Status for this loan — needed
    ONLY for the narrow CHANGE IN TERMS AGREEMENT resume exception below
    (condition 2: confirms an EXISTING deferral is ending, never used to
    establish a NEW one).

    servicer_key must be the canonical soup_key (e.g. "Millennium Bank",
    "Lendistry"), NOT a note-series-suffixed display label like "Millennium
    2025" — the memo file has no note-series field, so a suffixed label can
    never match it.
    """
    if memo_df is None or memo_df.empty or gp is None or cycle_month is None:
        return None, None, None, _deferral_evidence(EVIDENCE_NO_DEFERRAL)
    if "gp_sba" not in memo_df.columns and "gp_lender" not in memo_df.columns:
        return None, None, None, _deferral_evidence(EVIDENCE_NO_DEFERRAL)

    id_match = pd.Series(False, index=memo_df.index)
    if "gp_sba" in memo_df.columns:
        id_match = id_match | (memo_df["gp_sba"] == gp)
    if "gp_lender" in memo_df.columns:
        id_match = id_match | (memo_df["gp_lender"] == gp)
    # CLIENT-CONFIRMED (2026-08-10): a memo row attributed to THIS loan by
    # file_readers.resolve_memo_attribution -- the SBA-Loan-ID -> Lender-Loan-
    # Number fallback -- belongs to it. Purely additive: a row the two probes
    # above already matched has resolved_gp equal to gp anyway, and a row left
    # unresolved contributes nothing.
    if "resolved_gp" in memo_df.columns:
        id_match = id_match | (memo_df["resolved_gp"] == gp)

    if not id_match.any():
        # No memo entry exists for this loan at all. A FACT, not a defect --
        # deliberately distinct from lookup-failed below, which IS one.
        return None, None, None, _deferral_evidence(EVIDENCE_NO_DEFERRAL)

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

    # CLIENT-CONFIRMED (2026-08-10): for a row attributed by the SBA-Loan-ID ->
    # Lender-Loan-Number fallback, THE MATCHED LOAN DETERMINES THE SERVICER --
    # the memo's own Lender cell is not consulted. That cell is blank on every
    # row this fallback resolves (the workbook's INDEX/MATCH returned #N/A),
    # which is exactly why inferring the servicer from the memo cannot work.
    # Additive: it can only admit rows the loan index already tied to this
    # servicer, never reject one the alias table accepted.
    if "resolved_servicer" in memo_df.columns and servicer_key:
        resolved_here = (memo_df["resolved_servicer"].astype("string").str.strip().str.lower()
                         == str(servicer_key).strip().lower())
        lender_match = lender_match | resolved_here

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
        return None, warning, "lender_identity", _deferral_evidence(EVIDENCE_LOOKUP_FAILED)

    all_candidates = memo_df[id_match & lender_match].copy()

    # CLIENT-CONFIRMED (2026-08-04): a narrow, RESUME-ONLY exception for
    # Securitization Action = "CHANGE IN TERMS AGREEMENT". Checked before
    # the standard "Deferred"-only logic below, and deliberately never
    # treated as evidence of a NEW deferral — see _check_change_in_terms_resume
    # for the full set of required conditions. Does not generalize to
    # MODIFICATION, CHANGE MEMO, or any other Securitization Action value.
    cit_result, cit_row = _check_change_in_terms_resume(all_candidates, soup_status, cycle_month)
    if cit_result is not None:
        return cit_result, None, None, _deferral_evidence(EVIDENCE_DOCUMENTED, cit_row)

    candidates = all_candidates.copy()

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
        return None, None, None, _deferral_evidence(EVIDENCE_NO_DEFERRAL)
    # Rows the memo marks Deferred but left with no usable start date. They are
    # dropped from matching exactly as before -- the only change is that a loan
    # no longer renders as "no deferral" when THIS is the reason.
    claimed_without_dates = bool(candidates["Deferral Start Date"].isna().any())
    candidates = candidates.dropna(subset=["Deferral Start Date"])
    if candidates.empty:
        return None, None, None, _deferral_evidence(
            EVIDENCE_DATES_MISSING if claimed_without_dates else EVIDENCE_NO_DEFERRAL)

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
        return None, None, None, _deferral_evidence(EVIDENCE_NO_DEFERRAL)

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
    months_since_start = (cycle_period.year - chosen_start_month.year) * 12 + (cycle_period.month - chosen_start_month.month)
    if deferral_type in _FULL_PNI_CALC_TYPES:
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
    elif chosen_resume_month is not None and cycle_period == chosen_resume_month:
        # A PARTIAL deferral has no start/mid freeze -- it follows the normal
        # payment ladder throughout -- but it still ENDS, and its end is the
        # same event as a full deferral's: the cycle containing the Regular
        # Payment Resume Date. The resume boundary is a property of the
        # WINDOW, not of the deferral's type.
        #
        # Only "resume" is assigned here. start / mid / resume_unverified stay
        # full-only above: they describe the full-P&I freeze, which has no
        # partial equivalent. The full ladder above is untouched, including
        # its client-confirmed precedence (start before resume before mid), so
        # a full deferral's phase is decided exactly as before.
        phase = "resume"
    else:
        phase = None

    # The memo names a deferral but not its type. Surface it: the loan reaches
    # the reviewer and corrections.py's WARNING_REASONS excludes it from every
    # auto-correction. Only set when the phase logic above hasn't already
    # raised a more specific deferral_dates warning — that one says something
    # narrower about the same record, and replacing it would lose detail.
    if deferral_type == "unclassified" and warning is None:
        warning = _UNCLASSIFIED_TYPE_WARNING
        warning_type = "deferral_dates"

    # Evidence status for a deferral that DID match this cycle. Precedence is
    # deliberate: impossible dates outrank an unknown type, because a bad date
    # undermines the window itself (which cycle the deferral even covers),
    # whereas an unknown type leaves a valid window with one unknown attribute.
    if _memo_date_order_is_impossible(chosen):
        evidence_status = EVIDENCE_DATE_ANOMALY
    elif deferral_type == "unclassified":
        evidence_status = EVIDENCE_TYPE_UNCLASSIFIED
    else:
        evidence_status = EVIDENCE_DOCUMENTED

    return (
        {"deferral_type": deferral_type, "phase": phase, "months_since_start": months_since_start},
        warning, warning_type, _deferral_evidence(evidence_status, chosen),
    )


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

# CLIENT-CONFIRMED (2026-08-10): the other direction of the rule above.
#
# _VALID_EXIT_TRANSITION_KEYWORDS suppresses a flag when a loan ARRIVES at
# PAID-IN-FULL. Its own comment records the reverse case -- a loan LEAVING a
# closed status -- as "a different, unvalidated scenario" left flagged pending
# confirmation. That confirmation is now given:
#
#   When the client's OWN prior SOUP history already establishes that a loan
#   reached PAID-IN-FULL or LIQUIDATION with an Ending Actual Balance of 0,
#   the loan is CLOSED. It must not appear anywhere in the current
#   client-facing exception population, whatever Loan Status or balance the
#   servicer subsequently reports.
#
# This is an ELIGIBILITY gate, not a calculation: every value is still
# computed exactly as before, and audited_status_by_loan is still captured, so
# nothing about how the audit reasons changes. Only whether the loan is
# reportable changes.
#
# Zero is judged with the same BALANCE_TOLERANCE every other balance
# comparison in this module uses, and the whole SOUP history is scanned -- not
# just the latest row -- because a loan that closed and later reappeared is
# still closed on the client's own evidence.
_CLOSED_AT_ZERO_STATUS_KEYWORDS = ("PAID-IN-FULL", "PIF", "LIQUIDATION")


def loans_closed_at_zero(soup_servicer_df):
    """{gp: {closed_month, closed_status, closed_balance, closed_deal}} for
    every loan the client's own SOUP history establishes as CLOSED.

    Two qualifications, both learned from real data rather than assumed:

    1. NOT CONTRADICTED BY LATER ACTIVITY. A closed status at a zero balance
       only establishes closure if nothing afterwards says otherwise. Three
       Lendistry 2024 loans (4704279102, 4811009108, 4831479106) each carry a
       PAID-IN-FULL $0.00 row in 2025-08 and then keep amortising normally for
       nine more months. A loan that is still paying down is not closed, and
       the client's own later data is what proves it.

    2. DEAL-AWARE. Those same rows are a DEAL exit, not a loan closure: the
       PAID-IN-FULL row sits under SOUP2025_1 while the loan continues under
       SOUP2024_1, same GP, same DtbseID, same month. A closure of one Deal
       must never suppress a loan still active under another. So closure is
       resolved per (GP, Deal) and a GP qualifies only when EVERY one of its
       deals is closed.

    Compare GP 4338609104, which does qualify: a single deal, whose
    PAID-IN-FULL $0.00 row is its final SOUP activity with nothing after it.

    Returns {} when required columns are absent, so a file missing them
    degrades to "suppress nothing" rather than erroring. When SOUP carries no
    Deal column at all, every row falls into one partition and the rule
    reduces to qualification 1 alone.
    """
    if soup_servicer_df is None or soup_servicer_df.empty:
        return {}
    needed = {"SBA GP Number", "Loan Status", "Ending Actual Balance"}
    if not needed.issubset(soup_servicer_df.columns):
        return {}

    df = soup_servicer_df.copy()
    df["_gp"] = _norm_id(df["SBA GP Number"])
    df["_status"] = df["Loan Status"].astype("string").str.strip().str.upper()
    df["_bal"] = pd.to_numeric(df["Ending Actual Balance"], errors="coerce")
    df["_month"] = (pd.to_datetime(df["Payment Month"], errors="coerce")
                    if "Payment Month" in df.columns else pd.NaT)
    df["_deal"] = (df["Deal"].astype("string").str.strip()
                   if "Deal" in df.columns else "")
    df = df[df["_gp"].notna()]

    def _is_closed(status):
        return any(k in str(status) for k in _CLOSED_AT_ZERO_STATUS_KEYWORDS)

    out = {}
    for gp, gp_rows in df.groupby("_gp"):
        gp = str(gp).strip()
        deal_closures = {}
        all_deals_closed = True
        for deal, rows in gp_rows.groupby("_deal", dropna=False):
            rows = rows.sort_values("_month")
            closure = None
            for _, r in rows.iterrows():
                if _is_closed(r["_status"]) and pd.notna(r["_bal"]) and abs(r["_bal"]) <= BALANCE_TOLERANCE:
                    closure = r          # latest qualifying row in this deal
            if closure is None:
                all_deals_closed = False
                continue
            # Qualification 1: nothing later in THIS deal contradicts it.
            later = rows[rows["_month"] > closure["_month"]] if pd.notna(closure["_month"]) else rows.iloc[0:0]
            contradicted = any(
                (not _is_closed(r["_status"])) and pd.notna(r["_bal"]) and abs(r["_bal"]) > BALANCE_TOLERANCE
                for _, r in later.iterrows()
            )
            if contradicted:
                all_deals_closed = False
                continue
            deal_closures[deal] = closure

        # Qualification 2: every deal for this GP must be closed.
        if deal_closures and all_deals_closed:
            latest = max(deal_closures.values(),
                         key=lambda r: (pd.Timestamp.min if pd.isna(r["_month"]) else r["_month"]))
            out[gp] = {
                "closed_month": str(latest["_month"])[:10] if pd.notna(latest["_month"]) else "",
                "closed_status": str(latest["Loan Status"]).strip(),
                "closed_balance": round(float(latest["_bal"]), 2),
                "closed_deal": str(latest["_deal"]) if str(latest["_deal"]) else "",
            }
    return out


# CLIENT-CONFIRMED (2026-08-04): a terminal COLLECTION
# status appearing in EITHER the SERVICER'S REPORTED DELINQUENCY STATUS or
# the SERVICER'S REPORTED LOAN STATUS is a fundamentally different situation
# from a normal delinquency ranking disagreement. Two distinct consequences,
# both implemented below:
#
#   1. When found in Delinquency Status: that field holds a collection/
#      legal-process descriptor instead of a ranking value, so the
#      comparison against calculated_dq is short-circuited to REVIEW rather
#      than a misleading "DELINQUENCY STATUS MISMATCH" (see
#      status_match_check). The calculated_dq computation itself is
#      unaffected by this case.
#
#   2. When found in LOAN Status (client-confirmed 2026-08-04): FORECLOSURE/
#      LIQUIDATION/REO are ACTIVE, ongoing collection statuses, unlike
#      PAID-IN-FULL/CHARGE-OFF (which are genuinely FINISHED loans where a
#      delinquency ranking no longer applies). The generic terminal-status
#      short-circuit in _calculate_expected_status (_TERMINAL_STATUS_KEYWORDS,
#      below) would otherwise blank the ranking to CURRENT regardless of the
#      servicer's own reported value or any payment activity — this is
#      exactly the bug found on loan 5387239101 (Loan Status FORECLOSURE,
#      Delinquency Status 240 DAYS, blanked to CURRENT with the payment
#      amount having zero effect on the outcome, confirmed by testing $100
#      and $0 payments and getting identical results). Normal catch-up
#      logic must never IMPROVE the ranking for these loans; the servicer's
#      own reported ladder value is preserved instead of recalculated.
#
# Deliberately a SEPARATE, narrower constant from _TERMINAL_STATUS_KEYWORDS
# and _VALID_EXIT_TRANSITION_KEYWORDS above — those serve different,
# already-confirmed purposes (Loan Status terminal detection that DOES
# blank to CURRENT, for genuinely finished loans; a single confirmed valid
# EXIT transition). This one covers exactly the three values explicitly
# named by the client. Do NOT fold in PAID-IN-FULL, CHARGE-OFF, or
# BANKRUPTCY here without a separate confirmed rule — the client's request
# named FORECLOSURE/LIQUIDATION/REO specifically ("such as" examples, not
# necessarily this exact final list), and widening it should be a
# deliberate decision, not an assumption made here.
_TERMINAL_COLLECTION_KEYWORDS = ("FORECLOSURE", "LIQUIDATION", "REO")


def _calculate_expected_status(gp, srf_status, soup_status, soup_dq_raw, received, scheduled,
                                has_payment_cols, memo_df, servicer_name, cycle_month,
                                srf_dq=None,
                                scheduled_ambiguous=False,
                                servicing_fee=None):
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
    warning_type, is_terminal_this_cycle, reversal_info). expected_loan_status
    is None when normal comparison against SOUP's last status applies (no
    deferral override). reversal_info is None unless a negative Payment
    Received amount actually drove calculated_dq this cycle, in which case
    it's {"payment_equivalents_removed": N} — corrections.py uses this to
    require manual review for any reversal-driven result rather than
    auto-correcting it (client-confirmed 2026-08-03)."""
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

    # CLIENT-CONFIRMED (2026-08-04): checked against the SERVICER'S CURRENT
    # report specifically, not SOUP's history — this is about what the
    # servicer is reporting THIS cycle, per the client's explicit requirement
    # to detect this in the Servicer Loan Status field.
    srf_status_is_terminal_collection = any(k in srf_status_u for k in _TERMINAL_COLLECTION_KEYWORDS)

    # NARROWED 2026-08-08. A terminal status used to skip the memo lookup
    # entirely, which meant a servicer-reported LIQUIDATION/FORECLOSURE/REO
    # silently suppressed every piece of INDEPENDENT memo evidence for that
    # loan -- including evidence that would have established a different
    # Expected Loan Status (e.g. a documented deferral whose Regular Payment
    # Resume Date has passed, which sets expected to the resumed status rather
    # than SOUP's stale DEFERRED). The servicer's own claim must not be able
    # to switch off independent evidence about itself.
    #
    # Terminal COLLECTION loans now consult the memo exactly like any other
    # loan. Branch B statuses (PAID/PIF/CHARGE/CANCEL/WITHDRAWN) keep the
    # short-circuit unchanged -- they are being reviewed separately, and a
    # finished loan has no deferral timeline to place.
    #
    # The short-circuit literal widens with the call: a skipped loan's
    # evidence is stated as "no-deferral" rather than left a bare None for a
    # renderer to interpret.
    skip_memo_lookup = is_terminal_this_cycle and not srf_status_is_terminal_collection
    deferral, memo_warning, warning_type, deferral_evidence = (
        (None, None, None, _deferral_evidence(EVIDENCE_NO_DEFERRAL)) if skip_memo_lookup
        else _lookup_deferral(memo_df, gp, servicer_name, cycle_month, soup_status=soup_status)
    )

    prior_missed = _status_to_missed_count(soup_dq_raw) if has_payment_cols else None
    expected_loan_status = None
    reversal_info = None
    # Broadens the "prior_missed is not None" gate below just enough to
    # catch a negative Payment Received whose prior SOUP status couldn't be
    # parsed — that case must surface as UNVERIFIED (client-confirmed
    # safeguard), not silently fall through to the generic raw-text-echo
    # branch at the bottom of this function, which is what would otherwise
    # happen since prior_missed being None would skip past the branch that
    # would have handled it.
    received_is_negative = received is not None and pd.notna(received) and float(received) < 0

    # SUPERSEDED 2026-08-08 — the CLIENT-CONFIRMED (2026-08-04) rule that gave
    # FORECLOSURE/LIQUIDATION/REO their own delinquency calculation is no
    # longer in force. History, kept deliberately rather than deleted:
    #
    #   The 2026-08-04 rule computed these loans' Delinquency Status from a
    #   "degradation-only floor" (the prior missed count advanced, never
    #   reduced by catch-up credit) and then took max(floor, servicer's own
    #   reported DQ), trusting the servicer's value whenever it was worse and
    #   precise. It was introduced to fix loan 5387239101, where an open-ended
    #   "150+ Days" bucket parsed LOWER than the loan's own prior 240 DAYS and
    #   silently improved the ranking with no payment to justify it.
    #
    #   Superseded because terminal collection status and delinquency
    #   calculation are SEPARATE CONCEPTS. A servicer reporting LIQUIDATION is
    #   evidence to review, not a reason to compute this loan's delinquency
    #   differently from every other loan, and not a licence for the
    #   servicer's own reported DQ to determine the audited figure — that let
    #   the servicer's number validate itself.
    #
    # These loans now follow the SAME normal ladder as everything else:
    # calculated from prior SOUP delinquency status and this cycle's payment
    # activity, including catch-up credit. The servicer's reported DQ is
    # displayed for comparison and still raises DELINQUENCY STATUS MISMATCH
    # when it differs, but never enters the calculation. Note this also
    # retires the "150+ Days" hazard by construction rather than by special
    # case: a value that never enters the calculation cannot corrupt it.
    #
    # Everything INFORMATIONAL about these loans is unchanged: the TERMINAL
    # STATUS REVIEW reason, STAT-02, exclusion from auto-correction, and the
    # always-visible row.
    #
    # Branch B below (PAID/PIF/CHARGE/CANCEL/WITHDRAWN blanked to CURRENT) is
    # deliberately NOT changed here. Those statuses raise a different question
    # -- some may legitimately end the scheduled-payment relationship -- and
    # are being reviewed separately.
    if is_terminal_this_cycle and not srf_status_is_terminal_collection:
        calculated_dq = "CURRENT"
    # Keys on the PHASE, not the type: the reset at the end of a deferral is
    # the same event whatever kind of deferral it was. Reusing this branch
    # rather than adding a partial-specific copy keeps one implementation of
    # "reset to CURRENT, then evaluate the resume-month payment".
    elif deferral and deferral["phase"] == "resume":
        # Client-confirmed: the cycle containing Regular Payment Resume
        # Date resets the delinquency baseline to CURRENT, then applies
        # normal payment evaluation THIS cycle — the rank accumulated
        # during the deferral window is explicitly NOT carried forward.
        if scheduled_ambiguous:
            calculated_dq = "UNVERIFIED"
        elif has_payment_cols:
            new_missed, payment_equivalents_removed = _apply_payment_transition(0, received, scheduled, servicing_fee)
            calculated_dq = "UNVERIFIED" if new_missed is None else _missed_count_to_status(new_missed)
            if payment_equivalents_removed is not None:
                reversal_info = {"payment_equivalents_removed": payment_equivalents_removed}
        else:
            calculated_dq = "CURRENT"

        # BUG FIX (GP 5658839107): the resume cycle is the loan LEAVING
        # DEFERRED — that's the expected, correct outcome of a confirmed
        # resume, not a mismatch. Leaving expected_loan_status as None here
        # made it fall back to comparing against soup_status, which is
        # SOUP's PRIOR (still-DEFERRED) month — incorrectly flagging every
        # legitimate deferral exit as a Loan Status mismatch.
        #
        # CLIENT-CONFIRMED (2026-08-10): the expected outcome is ACTIVE,
        # DERIVED FROM THE MEMO — never copied from the servicer's own
        # current Loan Status. This branch previously did
        # `expected_loan_status = srf_status`, which was the last surviving
        # path where a servicer's assertion validated itself: the audit
        # adopted the servicer's status as the expected one, so the two
        # could never disagree and no LOAN STATUS MISMATCH was possible.
        # That is exactly the reasoning behind the SUPERSEDED 2026-08-08
        # block further down for terminal collection statuses; this branch
        # predates it (it was written for GP 5658839107) and was missed.
        # Found on GP 6717929105 (July 2026, Community Bank and Trust),
        # where a servicer-reported LIQUIDATION on a confirmed resume cycle
        # matched itself, produced no Loan Status recommendation, and left
        # the reviewer with no way to correct it.
        #
        # The independent evidence is the Servicer Memo's Regular Payment
        # Resume Date — a separate document from the SRF submission being
        # audited, and the very thing that put this cycle in the `resume`
        # phase. A loan whose confirmed deferral window has ended and whose
        # regular payment schedule has resumed is back in normal servicing,
        # which is ACTIVE. This branch already trusts that same date for
        # something more consequential (resetting the delinquency ladder to
        # CURRENT above, discarding the whole window's accumulated rank);
        # trusting it to move the number but not the status is incoherent.
        #
        # Deliberately type-agnostic, matching this branch's existing
        # phase-not-type keying: full, partial and unclassified deferrals
        # all end on the same memo-supplied resume date, and full-vs-partial
        # governs only the payment treatment DURING the window. An
        # unclassified deferral still carries DEFERRAL DATA QUALITY WARNING,
        # which corrections.WARNING_REASONS excludes from correction
        # candidacy entirely — so weak evidence is reported here but can
        # never produce a write candidate.
        #
        # Three arrivals are deliberately NOT handled here because they are
        # already handled independently, and none of them reads
        # expected_loan_status:
        #   PAID-IN-FULL/PIF — is_valid_exit_transition short-circuits the
        #       match test in status_match_check before the comparison.
        #   BANKRUPTCY       — the `if "BANKRUPT" in srf_status_u` override
        #       below runs AFTER this branch and wins as the final word.
        #   DEFERRED         — the contradiction path immediately below.
        #
        # The one case NOT auto-accepted: if the servicer is STILL
        # reporting DEFERRED on a cycle the memo confirms should have
        # resumed, that is a LOAN STATUS CONTRADICTION -- and it is reported as
        # one, because expected_loan_status is ACTIVE here too.
        #
        # SUPERSEDED 2026-08-10 (site 4). This sub-branch previously set
        # expected_loan_status = None and raised warning_type =
        # "deferral_dates", producing a DEFERRAL DATA QUALITY WARNING. Both
        # were wrong, for two separate reasons:
        #
        #   1. expected_loan_status = None fell through to
        #      `expected_loan_status or soup_status` in status_match_check,
        #      so the servicer's DEFERRED was compared against SOUP's STALE
        #      PRIOR month -- which still reads DEFERRED precisely because the
        #      resume had not happened yet. It "matched", and no LOAN STATUS
        #      MISMATCH was possible. Same self-validation defect as the
        #      `else` branch above, reached by a different route.
        #
        #   2. DEFERRAL DATA QUALITY WARNING means, in this codebase's own
        #      words, "found the right memo record, but its dates don't let us
        #      confidently place this cycle in the deferral timeline." That is
        #      false here: the dates place it exactly, evidence_status reads
        #      `documented`, and it is the SERVICER's report that disagrees
        #      with good evidence. Because the tag is in
        #      corrections.WARNING_REASONS it also excluded the loan from
        #      correction candidacy ENTIRELY -- silently suppressing the
        #      unrelated, deterministic Delinquency Status correction on the
        #      grounds that the evidence could not be trusted, when it could.
        #
        # Found on GP 6673679110 (July 2026, Community Bank and Trust): memo
        # row 183 confirms 2026-04-01 -> 2026-06-01 with Regular Payment
        # Resume Date 2026-07-01, the servicer still reported DEFERRED, and
        # the loan produced ZERO recommendations despite two real problems.
        #
        # The contradiction is not lost -- it is now expressed as LOAN STATUS
        # MISMATCH and carried to the reviewer as the review_reason on the
        # DEFERRED -> ACTIVE recommendation (corrections.
        # REVIEW_REASON_RESUME_CONTRADICTION), i.e. on the control where the
        # decision is actually made rather than in a column that suppressed it.
        #
        # Scoped to THIS call site only. The five other producers of
        # warning_type = "deferral_dates" are untouched and still suppress
        # corrections: blank Regular Payment Resume Date (resume_unverified),
        # indeterminate phase, unclassified deferral type, and the two
        # "reported DEFERRED with no memo deferral found at all" paths. Note
        # deliberately NOT implemented as a predicate on evidence_status ==
        # "documented": a resume_unverified loan with a BLANK resume date is
        # also `documented` (orderly dates, known type), so such a predicate
        # would silently disarm the most important warning in the set.
        #
        # No assignment to warning_type/memo_warning happens here at all now,
        # so a warning _lookup_deferral already raised for this loan (e.g. the
        # unclassified-type one) survives instead of being overwritten by this
        # message -- strictly more warning coverage than before, not less.
        expected_loan_status = "ACTIVE"
    elif deferral and deferral["deferral_type"] in _FULL_PNI_CALC_TYPES and deferral["phase"] == "resume_unverified":
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
    elif deferral and deferral["deferral_type"] in _FULL_PNI_CALC_TYPES:
        expected_loan_status = "DEFERRED"
        # CLIENT-CONFIRMED (2026-08-03): an external deferral (one initiated
        # by a servicer memo, not necessarily by a missed payment) does NOT
        # pause the delinquency ranking ladder. Every month spent in
        # deferral — including the first — is ranked as one more step up
        # that ladder, exactly like a normal missed payment would be under
        # _apply_payment_transition's "nothing received" branch (+1). This
        # replaces the prior behavior, which incorrectly held the rank
        # steady at whatever it was before the deferral began for the
        # deferral's first month.
        #
        # This is deliberately just "+1", the SAME single-step-per-cycle
        # formula every other branch in this function uses — NOT "+1 per
        # elapsed month since deferral start". prior_missed already comes
        # from SOUP's most recently recorded status, which in a real
        # multi-month audit is itself last month's calculated/corrected
        # value — SOUP's rolling history already carries the ladder forward
        # month over month on its own. Adding an elapsed-month count on top
        # of that would double-count every month after the first (caught by
        # testing against a simulated 3-month deferral before this was
        # deployed — month 3 came out at 120 DAYS instead of the correct 60
        # DAYS).
        calculated_dq = (
            _missed_count_to_status(prior_missed + 1)
            if prior_missed is not None else str(soup_dq_raw).strip().upper()
        )
    elif deferral and deferral["deferral_type"] == "partial":
        # Loan Status should read DEFERRED for the window, but DQ still
        # follows normal payment-vs-scheduled math below.
        expected_loan_status = "DEFERRED"
        if scheduled_ambiguous:
            calculated_dq = "UNVERIFIED"
        elif has_payment_cols and (prior_missed is not None or received_is_negative):
            if received_is_negative and prior_missed is None:
                calculated_dq = "UNVERIFIED"
            else:
                new_missed, payment_equivalents_removed = _apply_payment_transition(prior_missed, received, scheduled, servicing_fee)
                calculated_dq = "UNVERIFIED" if new_missed is None else _missed_count_to_status(new_missed)
                if payment_equivalents_removed is not None:
                    reversal_info = {"payment_equivalents_removed": payment_equivalents_removed}
        else:
            calculated_dq = str(soup_dq_raw).strip().upper()
    elif has_payment_cols and (prior_missed is not None or received_is_negative):
        if received_is_negative and prior_missed is None:
            calculated_dq = "UNVERIFIED"
        elif scheduled_ambiguous:
            calculated_dq = "UNVERIFIED"
        else:
            new_missed, payment_equivalents_removed = _apply_payment_transition(prior_missed, received, scheduled, servicing_fee)
            calculated_dq = "UNVERIFIED" if new_missed is None else _missed_count_to_status(new_missed)
            if payment_equivalents_removed is not None:
                reversal_info = {"payment_equivalents_removed": payment_equivalents_removed}
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

    # CLIENT-CONFIRMED (2026-08-03): the servicer's own reported BANKRUPTCY
    # status is accepted as authoritative for LOAN STATUS purposes. There is
    # no independent historical source (SOUP, the memo) that can confirm or
    # dispute WHEN a bankruptcy began, so the prior SOUP status (e.g. ACTIVE)
    # must never be used to contradict it — this is the same "arriving is not
    # a mismatch" principle already applied to PAID-IN-FULL
    # (_VALID_EXIT_TRANSITION_KEYWORDS), just implemented directly here by
    # setting expected_loan_status rather than only suppressing the match
    # boolean, so the displayed Expected Loan Status also reads BANKRUPTCY
    # rather than misleadingly showing the pre-bankruptcy SOUP status.
    #
    # Deliberately does NOT touch calculated_dq — Delinquency Status keeps
    # following its normal payment-history math regardless of bankruptcy,
    # since nothing here was asked to change that, only Loan Status.
    #
    # Deliberately one-directional, same as PAID-IN-FULL: only overrides
    # when the CURRENT servicer report says BANKRUPTCY. A loan LEAVING
    # BANKRUPTCY for something else is a materially different, unvalidated
    # scenario and stays subject to the normal comparison below.
    #
    # Placed last, after every deferral branch above, so it always wins as
    # the final word on Loan Status — a real bankruptcy report should not be
    # silently overridden by an old, no-longer-relevant deferral window.
    if "BANKRUPT" in srf_status_u:
        expected_loan_status = srf_status

    # SUPERSEDED 2026-08-08 — the CLIENT-CONFIRMED (2026-08-04) rule here
    # accepted a servicer-reported FORECLOSURE/LIQUIDATION/REO as
    # authoritative for Expected Loan Status, on the same "arriving is not a
    # mismatch" principle as BANKRUPTCY/PAID-IN-FULL above.
    #
    # Superseded because it let the servicer's assertion validate itself: the
    # audit adopted the servicer's status as the expected one, so the two
    # could never disagree and no LOAN STATUS MISMATCH was possible. A
    # servicer-reported terminal status is EVIDENCE TO REVIEW, not the source
    # of truth.
    #
    # Expected Loan Status is now determined from independent evidence, which
    # is what the existing fallback already does: SOUP's last recorded status
    # (`expected_loan_status or soup_status` in status_match_check). Where the
    # servicer's claim and that evidence disagree, LOAN STATUS MISMATCH is
    # raised — which is the point.
    #
    # The Servicer Memo is deliberately NOT used to infer or confirm
    # LIQUIDATION/FORECLOSURE/REO. Doing so requires an explicitly defined
    # business requirement establishing how such confirmation works; none
    # exists, so none is implemented. (The delivered memo carries zero rows
    # mentioning any of the three in Securitization Action or Servicer Action,
    # so there is nothing to read even if one wanted to.) Do not add memo-based
    # inference here without that requirement.
    if srf_status_is_terminal_collection and not expected_loan_status and not str(soup_status or "").strip():
        # The servicer claims a terminal status and there is NO independent
        # evidence to weigh it against. Do not guess, and do not adopt the
        # claim -- say plainly that it could not be established.
        expected_loan_status = "UNVERIFIED"

    return calculated_dq, expected_loan_status, deferral, memo_warning, warning_type, is_terminal_this_cycle, reversal_info, deferral_evidence


def status_match_check(servicer_inputs_df, soup_servicer_df, memo_df=None, servicer_name=None, closed_at_zero=None):
    records = []
    correction_data = []
    suppressed_closed_gps = []
    # CLIENT-CONFIRMED (2026-08-04): a record whose ONLY reason is the
    # informational "DEFERRAL RESUME INFORMATION" marker (both Loan Status
    # and DQ Status genuinely match, no real warning) must stay visible in
    # `records` for transparency, but must NOT count toward issues_found or
    # flip STAT-01's overall status to WARN — it isn't a mismatch, warning,
    # or audit error, it's a disclosure of how a clean result was reached.
    real_issue_count = 0
    # CLIENT-CONFIRMED (2026-08-04): a COMPLETE per-loan dataset, unlike
    # `records`/`correction_data` above which only ever contain the FLAGGED
    # subset of loans (a loan that matches cleanly hits `continue` below and
    # never reaches either list). Populated unconditionally for every loan,
    # captured BEFORE that skip-check runs, specifically so the Portfolio
    # Summary's Audited workbook can use STAT-01's real calculated_dq for
    # EVERY loan, not just the ones STAT-01 chose to flag. This is the
    # single authoritative source — see audited_portfolio_summary_record,
    # which consumes this dict rather than recomputing anything.
    audited_status_by_loan = {}

    required = {"SBA GP Number", "Loan Status", "Delinquency Status"}
    if not required.issubset(servicer_inputs_df.columns) or not required.issubset(soup_servicer_df.columns):
        result = _result("STAT-01", "Status / Delinquency Match", "SBA Reconciliation", "WARN", records,
                          lambda n: "")
        result["correction_data"] = []
        result["_cycle_month"] = None
        result["audited_status_by_loan"] = {}
        return result

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
        scheduled_payment_updated_during_cycle = bool(row.get("scheduled_payment_updated_during_cycle"))

        calculated_dq, expected_loan_status, deferral, memo_warning, warning_type, _, reversal_info, deferral_evidence = _calculate_expected_status(
            gp, srf_status, soup_status, soup_dq_raw, row.get("received"), row.get("scheduled"),
            has_payment_cols, memo_df, servicer_name, cycle_month,
            srf_dq=srf_dq, scheduled_ambiguous=scheduled_ambiguous,
            servicing_fee=row.get("servicing_fee"),
        )

        # Captured HERE — unconditionally, before any skip-check below can
        # short-circuit this iteration — so audited_status_by_loan ends up
        # complete for every loan STAT-01 was able to evaluate at all (a
        # loan absent from SOUP entirely is skipped even earlier than this
        # point and simply has no entry here; callers fall back to the raw
        # reported status for those, the same conservative "no information,
        # don't guess" default used elsewhere in this codebase).
        audited_status_by_loan[gp] = calculated_dq

        # CLOSED-AT-ZERO SUPPRESSION (see _CLOSED_AT_ZERO_STATUS_KEYWORDS).
        # Deliberately placed AFTER the capture above: the audited status is
        # still computed and recorded for this loan, because suppression is an
        # ELIGIBILITY decision about reporting, not a change to the audit's
        # reasoning. The Audited Portfolio Summary excludes the loan through
        # its own explicit exclusion set instead -- an explicitly excluded loan
        # and a loan with no audited status are different states, and the
        # fallback for the latter would wrongly bucket it on the servicer's
        # own reported value.
        #
        # Placed BEFORE the skip block below so no record and no
        # correction_data entry is produced -- which is what makes both the
        # Loan Status and Delinquency Status findings, and every downstream
        # recommendation, disappear by construction rather than by a second
        # filter in corrections.py.
        if closed_at_zero and str(gp).strip() in closed_at_zero:
            suppressed_closed_gps.append(str(gp).strip())
            continue

        # CLIENT-CONFIRMED (2026-08-04): a terminal collection status
        # (FORECLOSURE/LIQUIDATION/REO) can appear in EITHER the SERVICER'S
        # reported Delinquency Status OR the SERVICER'S reported Loan
        # Status. Either one is flagged as its own distinct review item
        # below, separately from a normal calculation disagreement. When
        # SUPERSEDED 2026-08-08: when found in Loan Status this no longer
        # affects calculated_dq at all -- those loans now use the same normal
        # ladder as every other loan. This flag is purely for the Reason tag,
        # STAT-02 visibility and exclusion from auto-correction.
        srf_dq_is_terminal_collection = any(k in srf_dq.upper() for k in _TERMINAL_COLLECTION_KEYWORDS)
        srf_status_is_terminal_collection = any(k in srf_status.upper() for k in _TERMINAL_COLLECTION_KEYWORDS)

        # UNVERIFIED means the calculation branch that ran actually needed
        # Scheduled Payment and couldn't trust it — this is NOT the same as
        # a confirmed delinquency mismatch, so it gets its own match state
        # ("REVIEW") rather than a hard YES/NO. Do not let the prior status
        # happening to differ from what was reported get relabeled as
        # DELINQUENCY STATUS MISMATCH just because calculated_dq isn't a
        # real status string.
        if calculated_dq == "UNVERIFIED":
            dq_match = "REVIEW"
        elif srf_dq_is_terminal_collection:
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

        # THE audit's effective Expected Loan Status: the independently
        # derived override when one applies (confirmed deferral, resume,
        # bankruptcy, UNVERIFIED), otherwise SOUP's last recorded status.
        # Computed ONCE and used by every consumer -- the match test below,
        # the display column, and corrections.py via correction_data -- so no
        # consumer re-derives it and they cannot drift apart. Previously this
        # expression appeared twice here and nowhere else, which is why
        # corrections.py had no way to see it.
        effective_expected_loan_status = expected_loan_status or soup_status
        status_match = is_valid_exit_transition or (srf_status.upper() == effective_expected_loan_status.upper())

        # CLIENT-CONFIRMED (2026-08-04): whenever the CHANGE IN TERMS
        # AGREEMENT resume exception (see _check_change_in_terms_resume)
        # produced this cycle's result, that must stay visible for review
        # even if Loan Status and DQ both come out clean — this is a
        # narrow, exception-driven acceptance, not a routine match, and the
        # client explicitly wants it distinguishable from an ordinary
        # resume sourced from a standard "Deferred" memo record.
        resume_via_change_in_terms_exception = bool(deferral.get("resume_via_change_in_terms_exception")) if deferral else False

        # A memo-related warning is surfaced even when the calculated status
        # happens to match what was reported — the point is that this
        # loan's deferral/memo context couldn't be verified at all, so a
        # coincidental match is not the same as a confirmed one. Same
        # principle for an ambiguous Scheduled Payment: the row must stay
        # visible even if Loan Status happens to match, since the DQ side
        # is explicitly unresolved (dq_match == "REVIEW", never treated as
        # a silent pass).
        if (
            status_match and dq_match is True and not memo_warning and not scheduled_ambiguous
            and not resume_via_change_in_terms_exception and not scheduled_payment_updated_during_cycle
            and not srf_status_is_terminal_collection
        ):
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
        # CLIENT-CONFIRMED (2026-08-04): distinct from DELINQUENCY STATUS
        # MISMATCH on purpose -- a terminal collection status in the DQ
        # field isn't a ranking disagreement, it's a different kind of
        # anomaly needing manual attention. Unlike the two informational
        # tags below, this DOES count toward real_issue_count (left out of
        # that exclusion set deliberately) and IS added to corrections.py's
        # WARNING_REASONS, since a loan with this flag should not have any
        # field auto-corrected without a human looking at it first.
        # CLIENT-CONFIRMED (2026-08-04): distinct from
        # DELINQUENCY STATUS MISMATCH on purpose -- a terminal collection
        # status in either field isn't a ranking disagreement, it's a
        # different kind of anomaly needing manual attention. Unlike the
        # two informational tags below, this DOES count toward
        # real_issue_count (left out of that exclusion set deliberately)
        # and IS added to corrections.py's WARNING_REASONS, since a loan
        # with this flag should not have any field auto-corrected without
        # a human looking at it first. Also appears in the separate STAT-02
        # Terminal Status Review output when triggered via Loan Status.
        if srf_dq_is_terminal_collection or srf_status_is_terminal_collection:
            reasons.append("TERMINAL STATUS REVIEW")
        if memo_warning and warning_type == "lender_identity":
            reasons.append("MEMO DATA QUALITY WARNING")
        if memo_warning and warning_type == "deferral_dates":
            reasons.append("DEFERRAL DATA QUALITY WARNING")
        if scheduled_ambiguous:
            reasons.append("SCHEDULED PAYMENT AMBIGUOUS")
        # Deliberately NOT added to corrections.py's WARNING_REASONS set —
        # this is an informational marker for a CONFIRMED, working-as-
        # designed exception, not a problem or uncertainty, so it must
        # never block auto-correction eligibility the way a real warning
        # reason does.
        if resume_via_change_in_terms_exception:
            reasons.append("DEFERRAL RESUME INFORMATION")
        # CLIENT-CONFIRMED (2026-08-04): same principle as the deferral
        # resume exception above — a resolved step-pattern update is a
        # confirmed, working-as-designed outcome, not a problem, so it must
        # never count as an issue/warning or block auto-correction.
        if scheduled_payment_updated_during_cycle:
            reasons.append("SCHEDULED PAYMENT UPDATED DURING CYCLE")

        if set(reasons) - {"DEFERRAL RESUME INFORMATION", "SCHEDULED PAYMENT UPDATED DURING CYCLE"}:
            real_issue_count += 1

        if srf_dq_is_terminal_collection and srf_status_is_terminal_collection:
            terminal_status_review_detail = (
                f"Both the Servicer Loan Status ({srf_status!r}) and Servicer Delinquency Status "
                f"({srf_dq!r}) contain a terminal collection status. This is a REVIEW SIGNAL "
                f"about the loan's collection state, not an input to the delinquency calculation: "
                f"the Calculated DQ Status above was computed by the same normal ladder used for "
                f"every other loan, from the prior SOUP delinquency status and this cycle's payment "
                f"activity. The servicer's reported Delinquency Status is shown for comparison only "
                f"and does not determine the calculated value. This loan needs manual review of the "
                f"reported terminal status itself."
            )
        elif srf_status_is_terminal_collection:
            terminal_status_review_detail = (
                f"Servicer Loan Status ({srf_status!r}) contains a terminal collection status. "
                f"This is a REVIEW SIGNAL that the servicer is reporting this status; it is not "
                f"treated as authoritative and does not change the delinquency calculation. The "
                f"Calculated DQ Status above was computed by the same normal ladder used for every "
                f"other loan, from the prior SOUP delinquency status and this cycle's payment "
                f"activity, and the Expected Loan Status was determined from independent evidence "
                f"rather than adopted from this report. Review whether the reported status is "
                f"supported."
            )
        elif srf_dq_is_terminal_collection:
            terminal_status_review_detail = (
                f"Servicer Delinquency Status ({srf_dq!r}) contains a terminal collection status "
                f"rather than a delinquency ranking. The Calculated DQ Status above was computed "
                f"normally from payment history and is unaffected; this loan needs manual review "
                f"rather than an automatic ranking comparison."
            )
        else:
            terminal_status_review_detail = ""

        records.append({
            "SBA GP Number": gp,
            "Reason": " + ".join(reasons),
            "Servicer Loan Status": srf_status,
            "Expected Loan Status": effective_expected_loan_status,
            "Loan Status Match": "YES" if status_match else "NO",
            "Servicer DQ Status": srf_dq,
            "Prior (SOUP) DQ Status": str(soup_dq_raw).strip(),
            "Calculated DQ Status": calculated_dq,
            "DQ Status Match": dq_match if dq_match == "REVIEW" else ("YES" if dq_match else "NO"),
            "Deferral Type": deferral["deferral_type"] if deferral else "",
            # --- Reviewer context: the six facts that explain WHY a deferral
            # row was flagged, as discrete columns rather than prose. STAT-01
            # is the only place all six coexist for one loan — the memo has no
            # SOUP or servicer status, and SOUP has no memo label — which is
            # why they are assembled here and merely rendered downstream.
            # A label conflict (a "Modification" Securitization Action against
            # a full-P&I window) is unreadable without the memo label sitting
            # beside the prior and current statuses.
            "Deferral Evidence Status": deferral_evidence["evidence_status"],
            "Memo Securitization Action": deferral_evidence["securitization_action"],
            "Deferral Start Date": deferral_evidence["deferral_start"],
            "Deferral End Date": deferral_evidence["deferral_end"],
            "Regular Payment Resume Date": deferral_evidence["deferral_resume"],
            "Prior (SOUP) Loan Status": str(soup_status).strip() if soup_status is not None else "",
            "Memo Lookup Warning": memo_warning if warning_type == "lender_identity" else "",
            "Deferral Data Quality Warning": memo_warning if warning_type == "deferral_dates" else "",
            "Deferral Resume Source": "CHANGE IN TERMS AGREEMENT exception (not a standard Deferred record)" if resume_via_change_in_terms_exception else "",
            "Scheduled Payment Update Detail": (
                f"Distinct values reported this cycle: {row.get('scheduled_distinct_values')} — "
                f"resolved as a one-directional step pattern; final chronological value "
                f"${row.get('scheduled'):,.2f} used for the delinquency calculation."
            ) if scheduled_payment_updated_during_cycle else "",
            "Terminal Status Review Detail": terminal_status_review_detail,
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

        # Raw facts only — no eligibility judgment here. Whether any of this
        # actually qualifies for auto-correction is decided entirely in
        # backend/corrections.py, kept deliberately separate from this
        # detection logic. expected_loan_status (NOT the "or soup_status"
        # display fallback above) is the field that lets corrections.py tell
        # "a confirmed deferral override" apart from "no override, happens
        # to read the same as SOUP" — those must not be conflated.
        correction_data.append({
            "gp": gp,
            "excel_row": row.get("_excel_row"),
            # CLIENT-CONFIRMED (2026-08-05): every posting row for this
            # loan/cycle, not just the last -- see consolidate_cycle_postings'
            # _all_excel_rows. corrections.py uses this to repeat an approved
            # correction across every posting row for the loan rather than
            # only the one `excel_row` (last posting) points at. `excel_row`
            # above is left in place unchanged for any other consumer.
            "excel_rows": row.get("_all_excel_rows") or [],
            "sheet_name": row.get("_sheet_name"),
            "header_excel_row": row.get("_header_excel_row"),
            "reasons": reasons,
            "dq_match": dq_match,
            "calculated_dq": calculated_dq,
            "srf_dq": srf_dq,
            "status_match": status_match,
            "expected_loan_status_raw": expected_loan_status,
            # The EFFECTIVE expected status (override, else SOUP), computed
            # once above. corrections.py's terminal-review branch proposes
            # THIS -- never soup_status directly, which would ignore every
            # deferral/resume/bankruptcy branch that legitimately changes the
            # answer. expected_loan_status_raw above stays as it is: the
            # deferral branch needs to tell a confirmed override apart from a
            # SOUP value that merely happens to read the same.
            "expected_loan_status_effective": effective_expected_loan_status,
            "srf_status": srf_status,
            "soup_status": soup_status,
            "deferral_type": deferral["deferral_type"] if deferral else None,
            "deferral_phase": deferral["phase"] if deferral else None,
            # The shared deferral-evidence contract, computed once above and
            # passed through verbatim. corrections.py, report.py and the review
            # page all read these; none of them recomputes or reinterprets.
            # Read-only review evidence, same rule as the block below: nothing
            # in this module's calculation logic may read them back.
            "evidence_status": deferral_evidence["evidence_status"],
            "deferral_start": deferral_evidence["deferral_start"],
            "deferral_end": deferral_evidence["deferral_end"],
            "deferral_resume": deferral_evidence["deferral_resume"],
            "securitization_action": deferral_evidence["securitization_action"],
            # CLIENT-CONFIRMED (2026-08-03): whenever a negative Payment
            # Received amount actually drove calculated_dq this cycle, that
            # result must require manual review before correction rather
            # than being auto-correctable — controlled here by exposing the
            # fact to corrections.py's eligibility logic (None when no
            # reversal was involved), NOT by giving STAT-01 a different
            # calculated_dq than NEG-01 would show for the same loan/cycle.
            "reversal_equivalents": reversal_info["payment_equivalents_removed"] if reversal_info else None,
            # CLIENT-CONFIRMED (2026-08-12): the offsetting-reversal companion
            # to reversal_equivalents above. reversal_equivalents answers "did a
            # NEGATIVE NET payment drive this cycle's figure"; this answers "was
            # there a reversal at all, even one an offsetting positive posting
            # cancelled out to exactly $0". The two are deliberately separate
            # facts, not one merged flag: only the first means the deficiency
            # formula ran, and corrections.py has to be able to tell them apart
            # to state the right reason and to leave the BR-002 ratio displayed
            # for a genuine $0 cycle. Raw facts only — neither is read by any
            # calculation in this module.
            "has_negative_posting": bool(row.get("has_negative_posting")),
            "offsetting_zero_net_reversal": _is_offsetting_zero_net_reversal(row),
            # CLIENT-CONFIRMED (2026-08-12): the CURRENT PERIOD's own
            # terminal-event dates, passed through verbatim from
            # consolidate_cycle_postings. corrections.py routes Loan Status
            # eligibility on bankruptcy_date_populated so that a BANKRUPTCY
            # recommendation rests on this cycle's evidence rather than on
            # soup_status, which is a carry-forward and is explicitly NOT
            # sufficient on its own. paid_in_full/liquidation carry the client's
            # zero-balance exception and the contradictory-evidence flag.
            #
            # Raw facts only. expected_loan_status, effective_expected_loan_
            # status, status_match and every delinquency figure above are
            # computed without reading any of these, and are unchanged by their
            # presence. Never feed one back into a calculation here.
            "bankruptcy_date": row.get("bankruptcy_date") or "",
            "bankruptcy_date_populated": bool(row.get("bankruptcy_date_populated")),
            "paid_in_full_date": row.get("paid_in_full_date") or "",
            "paid_in_full_date_populated": bool(row.get("paid_in_full_date_populated")),
            "liquidation_date": row.get("liquidation_date") or "",
            "liquidation_date_populated": bool(row.get("liquidation_date_populated")),
            # CLIENT-CONFIRMED (2026-08-13): DISPLAY-ONLY review evidence for the
            # Human Review loan card. Passed straight through from
            # consolidate_cycle_postings (_cycle_review_evidence); read by nothing
            # in this module's calculation logic and by nothing in corrections.py's
            # eligibility allow-list. A Loan Status contradiction is judged by a
            # human, and this is the evidence that human is shown -- it is never
            # an input to a rule.
            "foreclosure_date": row.get("foreclosure_date") or "",
            "reo_date": row.get("reo_date") or "",
            "liquidation_proceeds": row.get("liquidation_proceeds"),
            "liquidation_loss": row.get("liquidation_loss"),
            "subsequent_recoveries": row.get("subsequent_recoveries"),
            # CLIENT-CONFIRMED (2026-08-13): the PRIOR FINALISED reporting period's
            # Ending Balance -- SOUP's last recorded "Ending Actual Balance" for
            # this loan, off the SAME soup_latest row that already supplies
            # soup_status and soup_dq_raw above. The client's confirmed
            # PAID-IN-FULL / LIQUIDATION validity rule is keyed on this figure, so
            # a reviewer judging a terminal-status contradiction needs it on screen.
            #
            # DISPLAY ONLY, and deliberately so: this is deliberately NOT wired into
            # any terminal-status rule here. BAL-01 independently derives the same
            # figure for its own continuity check (see prior_ending there); this is
            # a second READ of one shared source, never a second derivation, and
            # nothing about STAT-01's calculated values changes by its presence.
            "prior_ending_balance": (
                _round2(soup_latest.loc[gp].get("Ending Actual Balance"))
                if gp in soup_latest.index
                and "Ending Actual Balance" in soup_latest.columns
                else None
            ),
            # --- Read-only review evidence (NOT used in any calculation) ---
            # Exposed purely so the correction review page can show, per
            # recommendation, the prior recorded status and the BR-002
            # payment evidence [(received + servicing_fee) / scheduled] that
            # drove calculated_dq, without the frontend having to recompute
            # anything. All sourced from the SAME consolidated `row` and
            # `soup_dq_raw` the flagged `records` entry above already uses.
            "prior_dq": str(soup_dq_raw).strip(),
            "payment_received": _round2(row.get("received")) if has_payment_cols else None,
            "scheduled_payment": _round2(row.get("scheduled")) if has_payment_cols else None,
            "servicing_fee": _round2(row.get("servicing_fee")),
            "ending_balance": _round2(row.get("ending_balance")),
            # consolidate_cycle_postings ALWAYS produces posting_count, so
            # None here should be unreachable. Defaulting to None (not 1) on
            # purpose: if it ever does go missing, the review page shows the
            # count as absent rather than silently claiming "1 posting" for a
            # loan that actually has several — misleading evidence on the very
            # screen a human uses to judge the correction.
            "posting_count": int(row["posting_count"]) if row.get("posting_count") is not None else None,
        })

    result = _result(
        "STAT-01", "Status / Delinquency Match", "SBA Reconciliation", "WARN", records,
        lambda n: f"{n} loan(s) where Loan Status doesn't match SOUP's last recorded status, or "
                  f"where Delinquency Status doesn't match the status calculated from "
                  f"missed-payment history (Payment Received vs. Scheduled Payment against the "
                  f"loan's prior recorded status).",
        issues_found=real_issue_count,
    )
    # Sibling key, deliberately NOT inside `records`: report.py's detail-tab
    # builder derives Excel columns from `list(records[0].keys())`, so
    # anything added to a record dict becomes a new visible column in the
    # existing audit report. Keeping this on a sibling key means report.py,
    # chatbot.py, and the frontend's rendering of `r.records` are completely
    # unaffected — main.py reads this key internally to build corrected
    # workbooks, then removes it before the API response is returned.
    result["correction_data"] = correction_data
    result["_cycle_month"] = cycle_month
    result["audited_status_by_loan"] = audited_status_by_loan
    # Loans withheld from this check's client-facing output by the
    # closed-at-zero rule. Internal only -- main.py logs it and pops it.
    result["suppressed_closed_gps"] = suppressed_closed_gps
    return result


# ---------------------------------------------------------------------------
# STAT-02 — Terminal Status Review
# CLIENT-CONFIRMED (2026-08-04): a separate, dedicated manual-review output
# for any loan whose SERVICER-reported Loan Status this cycle is FORECLOSURE,
# LIQUIDATION, or REO. This is deliberately independent of STAT-01 — STAT-01
# may or may not also flag the same loan (via its own "TERMINAL STATUS
# REVIEW" Reason tag, which fires for the identical condition plus the
# Delinquency-Status-field variant), but this check exists so these loans
# are visible in their own dedicated section for the client, not just buried
# as one Reason among several in a general reconciliation tab.
#
# Deliberately does NOT duplicate BAL-02's balance-validation logic — this
# check never looks at Ending Balance requirements or reports a PASS/FAIL
# verdict on the balance; it only surfaces these loans for review. BAL-02
# remains the sole check responsible for "does a LIQUIDATION loan report a
# $0 Ending Balance."
# ---------------------------------------------------------------------------

def terminal_status_review_check(servicer_inputs_df, soup_servicer_df, servicer_name=None, closed_at_zero=None):
    records = []

    required = {"SBA GP Number", "Loan Status"}
    if not required.issubset(servicer_inputs_df.columns):
        return _result("STAT-02", "Terminal Status Review", "SBA Reconciliation", "WARN", records,
                        lambda n: "")

    si_grouped = consolidate_cycle_postings(servicer_inputs_df)

    soup = soup_servicer_df.copy() if soup_servicer_df is not None else pd.DataFrame()
    soup_latest = pd.DataFrame()
    if not soup.empty and "SBA GP Number" in soup.columns:
        soup["SBA GP Number"] = _norm_id(soup["SBA GP Number"])
        sort_month = "Payment Month" if "Payment Month" in soup.columns else None
        if sort_month:
            soup[sort_month] = pd.to_datetime(soup[sort_month], errors="coerce")
        soup_latest = (soup.sort_values(["SBA GP Number", sort_month]) if sort_month else soup).groupby("SBA GP Number").tail(1)
        soup_latest = soup_latest.set_index("SBA GP Number")

    suppressed_closed_gps = []
    for gp, row in si_grouped.iterrows():
        # CLOSED-AT-ZERO SUPPRESSION (see _CLOSED_AT_ZERO_STATUS_KEYWORDS).
        # Prior SOUP already establishes this loan as closed at a zero
        # balance, so it is not eligible to be a client-facing exception
        # whatever the servicer now reports.
        if closed_at_zero and str(gp).strip() in closed_at_zero:
            suppressed_closed_gps.append(str(gp).strip())
            continue
        srf_status = str(row.get("loan_status", "")).strip()
        srf_status_u = srf_status.upper()

        matched_keyword = next((k for k in _TERMINAL_COLLECTION_KEYWORDS if k in srf_status_u), None)
        if matched_keyword is None:
            continue

        srf_dq = str(row.get("dq_status", "")).strip()
        soup_dq_raw = str(soup_latest.loc[gp].get("Delinquency Status", "")).strip() if gp in soup_latest.index else ""

        records.append({
            "SBA GP Number": gp,
            "Servicer": servicer_name or "",
            "Loan Status": srf_status,
            "Servicer DQ Status": srf_dq,
            "Prior SOUP DQ Status": soup_dq_raw,
            "Payment Received": round(float(row.get("received")), 2) if pd.notna(row.get("received")) else "",
            "Scheduled Payment": round(float(row.get("scheduled")), 2) if pd.notna(row.get("scheduled")) else "",
            "Ending Balance": round(float(row.get("ending_balance")), 2) if pd.notna(row.get("ending_balance")) else "",
            "Review Reason": f"{matched_keyword} REPORTED BY SERVICER",
        })

    records.sort(key=lambda r: r["SBA GP Number"])

    result = _result(
        "STAT-02", "Terminal Status Review", "SBA Reconciliation", "WARN", records,
        lambda n: f"{n} loan(s) reported by the servicer as FORECLOSURE, LIQUIDATION, or REO this "
                  f"cycle — surfaced for manual review. A servicer-reported terminal status is "
                  f"evidence to review, not the source of truth: these loans are excluded from "
                  f"automatic corrections, and their Calculated DQ Status is computed by the same "
                  f"normal delinquency ladder as every other loan (prior SOUP status plus this "
                  f"cycle's payment activity). The servicer's reported Delinquency Status is shown "
                  f"for comparison only and never determines the calculated value.",
    )
    result["suppressed_closed_gps"] = suppressed_closed_gps
    return result


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
    reversal_info_by_loan = {}
    for gp, row in si_grouped.iterrows():
        soup_status = str(soup_latest.loc[gp].get("Loan Status", "")).strip() if gp in soup_latest.index else ""
        soup_dq_raw = soup_latest.loc[gp].get("Delinquency Status", "") if gp in soup_latest.index else ""
        prior_status_by_loan[gp] = str(soup_dq_raw).strip()

        srf_status = str(row.get("loan_status", "")).strip()
        srf_dq = str(row.get("dq_status", "")).strip()
        calculated_dq, _, _, _, _, _, reversal_info, _ = _calculate_expected_status(
            gp, srf_status, soup_status, soup_dq_raw, row.get("received"), row.get("scheduled"),
            has_payment_cols, memo_df, servicer_name, cycle_month,
            srf_dq=srf_dq, scheduled_ambiguous=bool(row.get("scheduled_ambiguous")),
            servicing_fee=row.get("servicing_fee"),
        )
        # NEG-01 only ever DISPLAYS calculated_dq for manual-review context
        # (Current Period Status (Calculated)) — it never re-derives a
        # match/mismatch verdict from it, so "UNVERIFIED" safely passes
        # through as plain text here rather than risking being treated as
        # a real status label.
        calculated_dq_by_loan[gp] = calculated_dq
        reversal_info_by_loan[gp] = reversal_info

    # scan the RAW, un-aggregated rows -- a negative posting can be masked by
    # a positive one in the same cycle's summed total, so each posting needs
    # to be checked individually rather than relying on the aggregate sum
    negative_rows = si[si["Payment Received"] < 0]
    for _, row in negative_rows.iterrows():
        gp = row["SBA GP Number"]
        period = _fmt_date(row.get("Payment Receipt Date")) if "Payment Receipt Date" in si.columns else ""
        reversal_info = reversal_info_by_loan.get(gp)
        # CLIENT-CONFIRMED (2026-08-03): state exactly how many payment
        # equivalents this negative amount represented and how many
        # delinquency levels that produced, rather than a generic "review
        # manually" note with no numbers behind it. reversal_info is None
        # when this loan's CONSOLIDATED (netted) cycle total wasn't actually
        # negative -- e.g. a same-cycle positive posting fully offset this
        # one -- in which case the normal (non-reversal) calculation ran
        # instead, and the note reflects that rather than implying a
        # reversal calculation happened when it didn't.
        if reversal_info is not None:
            n = reversal_info["payment_equivalents_removed"]
            net_received = si_grouped.loc[gp].get("received")
            scheduled_this_cycle = si_grouped.loc[gp].get("scheduled")
            # CLIENT-CONFIRMED (2026-08-04): the calculation is based on the
            # loan's CONSOLIDATED net Payment Received for the whole cycle,
            # never an individual posting's own amount — a loan can have
            # several negative postings this cycle, or a positive posting
            # that partly offsets a negative one, and this row's own
            # "Negative Payment Amount" would misdescribe the calculation if
            # quoted here instead of the consolidated total that was
            # actually used. The deficiency includes the current cycle's own
            # unpaid scheduled payment, not just the removed amount — see
            # _apply_payment_transition's docstring for the exact formula.
            note = (
                f"The consolidated net Payment Received for this loan this cycle was "
                f"${net_received:,.2f} against a Scheduled Payment of ${scheduled_this_cycle:,.2f}, "
                f"a total deficiency of {n} scheduled payment equivalent(s) at the "
                f"{PAYMENT_SUFFICIENCY_THRESHOLD:.0%} sufficiency threshold (including the current "
                f"cycle's own unpaid scheduled payment). The delinquency ranking advanced {n} "
                f"level(s) from the prior SOUP status. Manual review is required before correction."
            )
        elif calculated_dq_by_loan.get(gp) == "UNVERIFIED":
            note = (
                "Reversed payment — the delinquency impact could not be determined (Scheduled Payment "
                "or the prior SOUP delinquency status is missing/unusable for this loan). Flagged for "
                "manual review."
            )
        else:
            note = (
                "Negative posting occurred this cycle, but the consolidated (netted) cycle total for "
                "this loan was not negative — reviewed for visibility only; the standard calculation "
                "applied."
            )
        records.append({
            "SBA GP Number": gp,
            "Period / Month": period,
            "Negative Payment Amount": round(float(row["Payment Received"]), 2),
            "Consolidated Net Payment This Cycle": round(float(si_grouped.loc[gp].get("received")), 2) if pd.notna(si_grouped.loc[gp].get("received")) else "",
            "Prior Period Status": prior_status_by_loan.get(gp, ""),
            "Current Period Status (Calculated)": calculated_dq_by_loan.get(gp, ""),
            "Note": note,
        })

    records.sort(key=lambda r: (r["SBA GP Number"], r["Period / Month"]))

    return _result(
        "NEG-01", "Negative Payment Flags", "SBA Reconciliation", "WARN", records,
        lambda n: f"{n} payment posting(s) with a negative Payment Received amount this cycle "
                  f"(typically a reversal) — flagged for manual review. The calculated status "
                  f"reflects a reversal-aware delinquency calculation (converting the negative "
                  f"amount into scheduled-payment equivalents), but is NOT automatically written "
                  f"to the corrected workbook; a reviewer should confirm the correct status by hand.",
    )


# ---------------------------------------------------------------------------
# PORT-01 — Portfolio Summary
# CLIENT-CONFIRMED (2026-08-04, revised): a multi-dimensional dollar-balance
# summary. Deliberately NOT a reconciliation check -- it never flags a
# mismatch, never drives a correction, and does not touch or depend on any
# delinquency/loan-status reconciliation logic elsewhere in this file. It
# only reads the SAME per-loan consolidated record
# (consolidate_cycle_postings) that STAT-01/NEG-01 already use.
#
# Two different dollar measures are used, by column, per explicit client
# revision (2026-08-04) of the original "Ending Balance for every column"
# spec:
#   - Outstanding, every delinquency bucket, Not Making Payment, Deferred,
#     Bankruptcy, and Foreclosure all use ENDING BALANCE. (Confirmed for Not
#     Making Payment against Lendistry July 2026: 60 loans summing to
#     $3,593,725.78 of Ending Balance, matching the client's expected figure.
#     Beginning Balance happens to give the same total there because every one
#     of those 60 loans has beginning == ending, so that reconciliation
#     confirms the loan set and the magnitude, not the field choice.)
#   - Paid In Full and Liquidation use BEGINNING BALANCE instead -- these
#     are treated as separate ACTIVITY metrics (how much left the portfolio
#     this cycle), not part of the remaining outstanding balance. Their
#     Ending Balance is expected to be $0 (confirmed against real CB&T
#     data: all 3 PAID-IN-FULL loans had Ending Balance = $0, while their
#     Beginning Balance summed to $679,563.79, matching the client's known
#     target of $679,564) -- Beginning Balance is what actually answers
#     "how much did we pay off / liquidate this cycle."
#
# Outstanding is ALSO no longer "every loan" -- it explicitly EXCLUDES
# PAID-IN-FULL and LIQUIDATION loans, since those balances are activity
# that has LEFT the remaining portfolio, not part of it. This makes
# Outstanding + Paid In Full + Liquidation NOT simply "the whole
# portfolio" in a single Ending-Balance sense -- Outstanding is Ending
# Balance for what remains, Paid In Full/Liquidation are Beginning
# Balance for what left.
#
# "Overlapping" still applies everywhere else: a loan's Delinquency
# Status, payment activity, and Loan Status describe independent
# characteristics, so a loan's Ending Balance can land in multiple
# columns at once (e.g. "30 Days" AND "Deferred" AND "Not Making
# Payment" simultaneously). A PAID-IN-FULL or LIQUIDATION loan is
# EXCLUDED from Outstanding specifically, but is NOT excluded from the
# other Ending-Balance columns (delinquency buckets, Not Making Payment,
# Deferred/Bankruptcy/Foreclosure) if it happens to independently satisfy
# one of them -- Loan Status is a reporting dimension, not a universal
# filter, except for the one explicit Outstanding exclusion above.
# ---------------------------------------------------------------------------

def _classify_delinquency_bucket(dq_status_upper):
    """Pure classification helper — no side effects, no balance math. Given
    an already-uppercased Delinquency Status string, returns the exact
    aging-bucket category name it belongs to, or None if unclassifiable
    (blank, unrecognized text). Used for BOTH the Original and Audited
    Portfolio Summary aging buckets via _build_portfolio_summary_record —
    this is the ONLY place this classification logic exists, specifically
    so the two workbooks can never diverge on how a given status string
    gets bucketed, only on WHICH status string each one is given."""
    if dq_status_upper == "CURRENT":
        return "Current"
    if dq_status_upper == "DELINQUENT":
        return "Delinquent"
    if dq_status_upper == "30 DAYS":
        return "30 Days"
    if dq_status_upper == "60 DAYS":
        return "60 Days"
    if dq_status_upper == "90 DAYS":
        return "90 Days"
    if dq_status_upper == "120 DAYS":
        return "120 Days"
    # "150 DAYS or greater, including labels such as 150+ DAYS, 180 DAYS,
    # 210 DAYS, 240 DAYS, and higher." Uses the SAME shared ladder parser
    # as every other check in this file for classification purposes only —
    # this is a descriptive bucket ("is this loan at least 150 days
    # delinquent"), not a numeric override calculation, so an open-ended
    # "150+ DAYS" label is legitimately treated as >= 150 days here (unlike
    # STAT-01's terminal-collection branch, which never trusts a
    # "+"-suffixed value as a PRECISE number to compute a new ranking from
    # — a different question than this one).
    missed = _status_to_missed_count(dq_status_upper)
    if missed is not None and missed >= 6:  # missed=6 == "150 DAYS"
        return "150+ Days"
    return None


_PORTFOLIO_SUMMARY_CATEGORY_ORDER = [
    "Current", "Delinquent", "30 Days", "60 Days", "90 Days", "120 Days", "150+ Days",
    "Not Making Payment", "Deferred", "Bankruptcy", "Foreclosure", "Liquidation", "Paid In Full",
]

# Internal passthrough column, written ONLY by _unit_weight_frame and read ONLY
# by _build_portfolio_summary_record's Not Making Payment rule. Carries
# "Beginning Balance == Ending Balance" evaluated on the ORIGINAL balances,
# because _unit_weight_frame replaces both with unit weights and the comparison
# would otherwise be trivially true in the counts pass. Not a consolidated fact
# and never part of any record the caller sees.
_NMP_BALANCES_UNCHANGED = "_nmp_balances_unchanged"


def _build_portfolio_summary_record(si_grouped, servicer_name, dq_status_for_aging_fn):
    """The SINGLE shared implementation behind both portfolio_summary_check
    (Original) and audited_portfolio_summary_record (Audited) — every rule
    lives here exactly once. dq_status_for_aging_fn(gp, row) -> str is the
    ONLY thing that differs between the two callers: which Delinquency
    Status value feeds the Current-through-150+ aging buckets. Every other
    category (Outstanding, Not Making Payment, Deferred, Bankruptcy,
    Foreclosure, Liquidation, Paid In Full) is computed identically
    regardless, which is exactly why those columns are guaranteed identical
    between the two workbooks — they never touch dq_status_for_aging_fn."""
    totals = {c: 0.0 for c in _PORTFOLIO_SUMMARY_CATEGORY_ORDER}
    outstanding = 0.0

    for gp, row in si_grouped.iterrows():
        ending_balance_raw = row.get("ending_balance")
        ending_balance = float(ending_balance_raw) if ending_balance_raw is not None and pd.notna(ending_balance_raw) else 0.0
        beginning_balance_raw = row.get("beginning_balance")
        beginning_balance = float(beginning_balance_raw) if beginning_balance_raw is not None and pd.notna(beginning_balance_raw) else 0.0

        # CLIENT-CONFIRMED (2026-08-05): Portfolio Summary eligibility gate.
        # A consolidated loan with Beginning Balance = 0 AND Ending Balance
        # = 0 does not represent an active portfolio loan for this cycle
        # and is excluded ENTIRELY, before any categorization -- it must
        # not contribute to Outstanding, any aging bucket, Not Making
        # Payment, Deferred, Bankruptcy, Foreclosure, Liquidation, or Paid
        # In Full. Deliberately checked BEFORE loan_status/dq_status are
        # even inspected, per the client's explicit preference for one
        # single eligibility gate over scattering the condition into each
        # category below. A loan with either balance greater than zero
        # remains fully eligible and flows through the existing rules
        # unchanged.
        if beginning_balance == 0 and ending_balance == 0:
            continue

        loan_status = str(row.get("loan_status", "")).strip().upper()
        # CLIENT-CONFIRMED (2026-08-04, corrected): exact matching only —
        # the prior substring check ("PAID" in loan_status) was too broad
        # and could incorrectly classify values like PARTIALLY PAID,
        # UNPAID, or PAID DOWN as Paid In Full.
        paid_in_full_statuses = {"PAID-IN-FULL", "PAID IN FULL", "PIF"}
        is_paid_in_full = loan_status in paid_in_full_statuses
        is_liquidation = loan_status == "LIQUIDATION"

        # CLIENT-CONFIRMED (2026-08-04, revised): Outstanding represents the
        # remaining portfolio only -- excludes PAID-IN-FULL and LIQUIDATION
        # loans entirely (their balances are tracked separately below).
        # Identical for Original and Audited — never depends on
        # dq_status_for_aging_fn.
        if not is_paid_in_full and not is_liquidation:
            outstanding += ending_balance

        # Delinquency Status aging buckets — the ONLY category that differs
        # between Original and Audited. dq_status_for_aging_fn decides
        # WHICH status string to classify; _classify_delinquency_bucket
        # decides WHAT bucket it lands in — the same classifier either way.
        dq_for_aging = str(dq_status_for_aging_fn(gp, row) or "").strip().upper()
        bucket = _classify_delinquency_bucket(dq_for_aging)
        if bucket is not None:
            totals[bucket] += ending_balance

        # CLIENT-CONFIRMED (2026-08-04): "Not Making Payment" is based
        # SOLELY on consolidated Payment Received = 0. Loan Status is
        # intentionally ignored. A deferred, bankruptcy, foreclosure,
        # liquidation, or other special-status loan may still contribute to
        # this metric if no payment was received during the cycle. This was
        # explicitly re-confirmed after investigating why an earlier
        # candidate target value ($665,889 for CB&T) appeared to exclude
        # DEFERRED loans — traced down to the raw postings and found no
        # data-quality explanation (no duplicate postings, no blank-vs-zero
        # ambiguity, no reporting-period anomaly) for excluding them; the
        # client confirmed the exclusion-based number reflected a
        # different, no-longer-used business definition. Identical for
        # Original and Audited.
        #
        # CLIENT-CONFIRMED (2026-08-04, corrected): a blank/missing/NaN
        # Payment Received value must NOT be treated as $0 here — the prior
        # implementation silently defaulted a missing value to 0.0, which
        # incorrectly classified a loan with no payment DATA at all
        # (unknown) the same as a loan CONFIRMED to have received nothing.
        # Only an explicitly present value that equals exactly 0 counts.
        #
        # CLIENT-CONFIRMED (2026-08-12): a SECOND, narrower way in, for a blank
        # Payment Received that is CORROBORATED as no payment activity rather
        # than as absent data. The 2026-08-04 rule above stands unchanged and
        # still decides every present value; this only reaches rows it would
        # otherwise drop.
        #
        # WHY. Measured on Lendistry's July 2026 file: 37 loans reported a
        # single filed posting with Payment Received, Principal, Net Interest,
        # Scheduled Payment and Payment Receipt Date ALL empty, no balance
        # movement, and 120+ day delinquency on 35 of them. 36 of those 37 were
        # filed by the same servicer as an explicit 0 the month before, for the
        # same loans in the same condition — the servicer changed how it writes
        # "no payment this cycle", not what happened. The blank-is-unknown rule
        # then dropped $2,241,118.06 the audit had counted in June. Excluding
        # them is the defect; the client's expected 60 loans / $3,593,725.78
        # reproduces exactly with this branch and only with it.
        #
        # THE DISCRIMINATOR IS has_filed_posting, NOT THE SERVICER. A blank
        # payment field on a posting the servicer actually filed asserts "no
        # payment"; the same blank on a row carrying no posting identity at all
        # asserts nothing, and stays unknown. Measured across all 20 available
        # servicer/period files, this separates perfectly: the 37 Lendistry-2025
        # rows all carry posting identity, while 2 Lendistry-2024 rows (GP
        # 4490649110, GP 5462759105) carry none and are correctly left out —
        # they are unvalidated by the client and unsupported by the data.
        #
        # Deliberately NOT part of the corroboration: Payment Receipt Date and
        # Net Interest. Both are blank on all 39 blank-payment loans in every
        # available file, so neither discriminates anything, and the receipt
        # date is not even carried on the consolidated frame.
        received_raw = row.get("received")
        received_present_zero = (
            received_raw is not None
            and pd.notna(received_raw)
            and float(received_raw) == 0
        )

        principal_raw = row.get("principal")
        principal_blank_or_zero = (
            principal_raw is None or pd.isna(principal_raw) or float(principal_raw) == 0
        )
        # _unit_weight_frame REWRITES both balance columns to unit weights, so
        # comparing them here would be trivially true in the counts pass and
        # would let counts classify a loan the dollars pass rejected — the exact
        # drift that function's docstring promises is impossible. It therefore
        # precomputes this predicate from the ORIGINAL balances and passes it
        # through; the fallback covers the dollars pass, where the real
        # balances are still present and the column is absent.
        balances_unchanged = row.get(_NMP_BALANCES_UNCHANGED)
        if balances_unchanged is None or (
            not isinstance(balances_unchanged, bool) and pd.isna(balances_unchanged)
        ):
            balances_unchanged = beginning_balance == ending_balance

        received_blank_corroborated = (
            (received_raw is None or pd.isna(received_raw))
            and principal_blank_or_zero
            and bool(balances_unchanged)
            and bool(row.get("has_filed_posting"))
        )

        if received_present_zero or received_blank_corroborated:
            totals["Not Making Payment"] += ending_balance

        # Loan Status buckets. Deferred/Bankruptcy/Foreclosure are mutually
        # exclusive with each other and use Ending Balance, identical for
        # Original and Audited. Paid In Full and Liquidation are handled
        # separately below using Beginning Balance instead — activity
        # metrics, not a "current standing" bucket.
        if loan_status == "DEFERRED":
            totals["Deferred"] += ending_balance
        elif "BANKRUPT" in loan_status:
            totals["Bankruptcy"] += ending_balance
        elif "FORECLOS" in loan_status:
            totals["Foreclosure"] += ending_balance

        if is_liquidation:
            totals["Liquidation"] += beginning_balance
        if is_paid_in_full:
            totals["Paid In Full"] += beginning_balance

    # CLIENT-CONFIRMED (2026-08-13): OUTSTANDING ENDING BALANCE IS THE
    # DENOMINATOR for every Portfolio Summary percentage, Paid In Full and
    # Liquidation included.
    #
    # Paid In Full and Liquidation are CURRENT PERIOD TERMINAL ACTIVITY
    # measures, and their dollar numerators use Beginning Balance under their
    # own existing business rules (see the terminal-activity block just above),
    # while this denominator is the Ending Balance of the loans that REMAIN --
    # which explicitly excludes the PIF and Liquidation loans themselves. So the
    # numerator and the denominator are drawn from different balance points and
    # from disjoint sets of loans.
    #
    # THE MIXED BASIS IS INTENTIONAL. It is not an oversight, and it must NOT be
    # "normalized" -- neither by moving this denominator to Beginning Balance nor
    # by switching the terminal-activity numerators to Ending Balance -- without
    # a NEW client-approved business rule. Pinned by
    # test_portfolio_summary_percentage_basis.py, which fails if either basis
    # moves.
    def pct(amount):
        return round((amount / outstanding) * 100, 2) if outstanding else 0.0

    record = {"Servicer": servicer_name or "", "Outstanding ($)": round(outstanding, 2), "Outstanding (%)": 100.0 if outstanding else 0.0}
    for c in _PORTFOLIO_SUMMARY_CATEGORY_ORDER:
        record[f"{c} ($)"] = round(totals[c], 2)
        record[f"{c} (%)"] = pct(totals[c])
    return record


def portfolio_summary_check(servicer_inputs_df, servicer_name=None):
    """PORT-01 — the ORIGINAL Portfolio Summary: aging buckets use the
    consolidated, servicer-reported Delinquency Status as-is (dq_status
    from consolidate_cycle_postings), unaudited.

    CLIENT-CONFIRMED (2026-08-10): displayed as "Original Portfolio Summary".
    This is the LENDER-REPORTED STARTING POSITION and is deliberately RETAINED
    -- it has audit value as the before picture. It is NOT the audited figure
    and NOT the post-review figure. Three distinct layers exist:

        Original Portfolio Summary       this function -- what the lender said
        Audited Portfolio Summary        audited_portfolio_summary_record --
                                         what the audit calculated
        Final Audited Portfolio Summary  verified_portfolio_summary_record --
                                         what the reviewer verified

    The label matters: these values were previously shown on the results
    screen as a bare "Portfolio Summary", indistinguishable from the audited
    workbook the client downloads, which is exactly how the two came to be
    read as the same report disagreeing with itself rather than as two
    different reports. Naming is the whole fix here -- no value changes."""
    si_grouped = consolidate_cycle_postings(servicer_inputs_df)
    record = _build_portfolio_summary_record(
        si_grouped, servicer_name, lambda gp, row: row.get("dq_status", "")
    )

    result = _result(
        "PORT-01", "Original Portfolio Summary", "SBA Reconciliation", "INFO", [record],
        lambda n: "",
        issues_found=0,  # a summary, never an "issue" -- stays visible via the informational override, doesn't count toward the audit's issue tally
    )
    return result


def audited_portfolio_summary_record(servicer_inputs_df, audited_status_by_loan, servicer_name=None,
                                     excluded_gps=None):
    """CLIENT-CONFIRMED (2026-08-04): the AUDITED Portfolio Summary record
    — identical to portfolio_summary_check in every respect EXCEPT the
    aging-bucket source, which uses STAT-01's real calculated_dq for every
    loan (audited_status_by_loan, from status_match_check's
    audited_status_by_loan — see that function's docstring for how it's
    built as a COMPLETE per-loan dataset, not just the flagged subset).
    Falls back to the loan's own raw dq_status when a loan has no entry in
    audited_status_by_loan (e.g. it was never found in SOUP at all, so
    STAT-01 could never evaluate it) — the same conservative "no
    information, don't guess" default used elsewhere in this codebase.

    Returns the raw record dict directly, NOT wrapped in _result() — this
    isn't its own separate check shown in the reconciliation results
    table, it exists purely to feed the Audited Portfolio Summary
    workbook."""
    si_grouped = consolidate_cycle_postings(servicer_inputs_df)
    # CLIENT-CONFIRMED (2026-08-10): a loan the closed-at-zero rule suppressed
    # is EXPLICITLY EXCLUDED from the audited population -- it contributes
    # neither its count nor its (stale) current servicer balance to any
    # audited category.
    #
    # Done by dropping the row, NOT by removing the loan from
    # audited_status_by_loan: the lambda below falls back to the servicer's own
    # reported dq_status when a loan has no audited entry, so omission would
    # bucket the loan on the very value the audit rejected. An explicitly
    # excluded loan and a loan with no audited status are different states.
    if excluded_gps:
        excluded = {str(g).strip() for g in excluded_gps}
        si_grouped = si_grouped[~si_grouped.index.astype(str).str.strip().isin(excluded)]
    return _build_portfolio_summary_record(
        si_grouped, servicer_name,
        lambda gp, row: audited_status_by_loan.get(gp, row.get("dq_status", "")),
    )


# Which consolidated column each reviewable field writes to. This mapping is
# the ONLY route a reviewer decision has into the Portfolio Summary, and it is
# deliberately limited to the two fields Manual Review may correct.
#
# CLIENT-CONFIRMED (2026-08-10): financial values are NOT manually editable.
# beginning_balance, ending_balance and received have no entry here and cannot
# be reached from a review decision -- a financial value can only ever enter
# the verified frame from the lender's own submission. If a financial value is
# wrong, that requires servicer correction or resubmission, not a silent edit
# through this feature. The correction allow-list (corrections.
# build_eligible_corrections) never proposes a financial field in the first
# place, so this is the second of two independent barriers.
_VERIFIED_OVERRIDE_COLUMNS = {
    "Loan Status": "loan_status",
    "Delinquency Status": "dq_status",
}


def _apply_verified_overrides(si_grouped, verified_overrides):
    """Returns a COPY of the consolidated frame with reviewer-approved values
    written over the lender's. verified_overrides is
    {(gp, field): approved_value} exactly as corrections.
    generate_reviewed_workbook returns it.

    Unknown GPs and unmapped fields are ignored rather than raising: the map
    comes from a completed review, and a loan or field that isn't in this
    frame simply has nothing to override. Never mutates the input."""
    if not verified_overrides:
        return si_grouped
    out = si_grouped.copy()
    index = {str(gp).strip(): gp for gp in out.index}
    for (gp, field), value in verified_overrides.items():
        column = _VERIFIED_OVERRIDE_COLUMNS.get(field)
        key = index.get(str(gp).strip())
        if column is None or key is None or column not in out.columns:
            continue
        out.at[key, column] = value
    return out


def _unit_weight_frame(si_grouped):
    """A COPY of the consolidated frame with every balance replaced by a unit
    weight, so the SAME builder that sums dollars sums COUNTS instead.

    Every Portfolio Summary category is a pure membership test followed by a
    sum, so weighting each loan 1.0 turns each dollar total into the number of
    loans in that category -- with no second classification engine and no
    change to _build_portfolio_summary_record. The eligibility gate, the
    Outstanding exclusion, _classify_delinquency_bucket, the Not Making
    Payment rule and the Deferred/Bankruptcy/Foreclosure chain are all reused
    verbatim, so counts and dollars cannot drift apart in what they classify.

    The (beginning == 0 AND ending == 0) eligibility gate is PRESERVED by
    weighting only the loans that pass it -- an ineligible loan keeps 0/0 and
    is skipped by the builder exactly as it is for dollars. Weighting every
    row unconditionally would silently re-admit loans the client confirmed
    must be excluded before any categorisation.

    Any predicate the builder derives FROM the balances must also be preserved
    here, for the same reason and it is not optional: once both columns are 1.0,
    "Beginning Balance == Ending Balance" is trivially true, so a rule using it
    would classify in the counts pass what it rejected in the dollars pass --
    precisely the drift this function exists to make impossible. The Not Making
    Payment corroboration rule uses that comparison, so it is evaluated here on
    the ORIGINAL values and passed through on _NMP_BALANCES_UNCHANGED. Every
    other input that rule reads (received, principal, has_filed_posting) is left
    untouched by this function and needs no such treatment.

    Never mutates the input."""
    out = si_grouped.copy()
    beginning = pd.to_numeric(out.get("beginning_balance"), errors="coerce").fillna(0.0)
    ending = pd.to_numeric(out.get("ending_balance"), errors="coerce").fillna(0.0)
    eligible = ~((beginning == 0) & (ending == 0))
    out[_NMP_BALANCES_UNCHANGED] = (beginning == ending)
    out["beginning_balance"] = eligible.astype(float)
    out["ending_balance"] = eligible.astype(float)
    return out


# CLIENT-CONFIRMED (2026-08-13): the loan count of the population the summary
# actually reports on, carried beside the categories on the counts record and
# rendered in the Final Audited summary's Current Period Terminal Activity
# section. Named as a key rather than added to
# _PORTFOLIO_SUMMARY_CATEGORY_ORDER on purpose: it is NOT a category. Adding it
# there would give it a column in COUNT / BALANCE % / BALANCE ($) and a
# meaningless "Servicing Population ($)" dollar figure, and would change the
# three existing sections this change must leave alone.
SERVICING_POPULATION = "Servicing Population"


def _eligible_loan_count(si_grouped):
    """The number of loans that pass the Portfolio Summary eligibility gate.

    CLIENT-CONFIRMED (2026-08-13): computed INDEPENDENTLY from the population,
    never as Outstanding + Paid In Full + Liquidation. That arithmetic happens
    to be right on the June 2026 accepted SOUP 2025 files only because every
    $0-Ending-Balance loan there is a Paid In Full, and it breaks in two ways
    that occur in real data: it double-counts a PAID-IN-FULL or LIQUIDATION
    loan carrying a positive balance (Mission Valley GP 4778059109, $276,702.10)
    and it misses an ACTIVE loan at $0. Categories overlap by design, so no sum
    over them is a population count.

    Derived through _unit_weight_frame so the eligibility gate is encoded in
    exactly ONE place: that function already weights each eligible loan 1.0 and
    each ineligible loan 0.0, so summing the weights IS the eligible-loan
    count. A loan excluded by the gate, or dropped upstream by closed-at-zero
    suppression (it never reaches this frame), contributes nothing -- which is
    what keeps this figure consistent with every other number in the same
    report rather than being a second, competing population."""
    if si_grouped is None or len(si_grouped) == 0:
        return 0
    weights = pd.to_numeric(
        _unit_weight_frame(si_grouped).get("ending_balance"), errors="coerce"
    ).fillna(0.0)
    return int(round(float(weights.sum())))


def _verified_population(servicer_inputs_df, verified_overrides, excluded_gps=None):
    """THE verified loan population, built ONCE:

        original lender data + verified_overrides + closed-at-zero suppression

    CLIENT-CONFIRMED (2026-08-10): every reported view -- Count, Balance % and
    Balance ($) -- must be derived from this one frame. Three independently
    constructed populations is precisely the failure mode this exists to
    prevent."""
    si_grouped = _apply_verified_overrides(
        consolidate_cycle_postings(servicer_inputs_df), verified_overrides
    )
    if excluded_gps:
        excluded = {str(g).strip() for g in excluded_gps}
        si_grouped = si_grouped[~si_grouped.index.astype(str).str.strip().isin(excluded)]
    return si_grouped


def verified_portfolio_summary_views(servicer_inputs_df, verified_overrides,
                                     servicer_name=None, excluded_gps=None):
    """Both reported views for one servicer, from ONE verified population.

    Returns {"dollars": <record>, "counts": <record>} where "dollars" is
    byte-identical to what verified_portfolio_summary_record returns (that
    function now delegates here, so there is one construction path, not two)
    and "counts" carries integer loan counts under the same category keys.

    The percentages live on the dollars record and are unchanged: pct() there
    divides by Outstanding and returns 0.0 when Outstanding is zero."""
    population = _verified_population(servicer_inputs_df, verified_overrides, excluded_gps)
    aging_source = lambda gp, row: row.get("dq_status", "")

    dollars = _build_portfolio_summary_record(population, servicer_name, aging_source)
    weighted = _build_portfolio_summary_record(
        _unit_weight_frame(population), servicer_name, aging_source
    )
    counts = {"Servicer": servicer_name or ""}
    for category in ["Outstanding"] + _PORTFOLIO_SUMMARY_CATEGORY_ORDER:
        counts[category] = int(round(weighted[f"{category} ($)"]))
    # From THIS population -- the verified one, after verified_overrides and
    # closed-at-zero exclusions -- so it cannot describe a different set of
    # loans than the categories beside it. See _eligible_loan_count.
    counts[SERVICING_POPULATION] = _eligible_loan_count(population)
    return {"dollars": dollars, "counts": counts}


def original_portfolio_summary_views(servicer_inputs_df, servicer_name=None):
    """Both reported views for one servicer, from the LENDER'S OWN SUBMISSION
    and nothing else.

    CLIENT-CONFIRMED (2026-08-11): the Original Portfolio Summary adopts the
    Final Audited summary's presentation -- three stacked sections, per-SOUP-
    deal totals, loan counts alongside dollars. That is a PRESENTATION AND
    AGGREGATION change only: every figure stays exactly what PORT-01 already
    reports, and stays invariant to everything the audit or a reviewer does.

    Deliberately NOT verified_portfolio_summary_views called with empty
    arguments. That function is provably equivalent when handed {} and None --
    confirmed record-for-record against all eight June servicers -- but it
    ACCEPTS verified_overrides and excluded_gps, and reusing it would reduce
    this summary's guarantee from "impossible: there is no parameter" to
    "correct: this one call site passes empty values". The signature IS the
    guarantee here, exactly as it already is for portfolio_summary_check
    (whose two signature tests exist for this reason), and one future call
    site passing a non-empty value is all it would take to break it.

    So, by construction: no verified_overrides, no excluded_gps, no
    audited_status_by_loan, no reviewer decisions, no correction inputs. The
    only input is the lender's submitted Servicer Inputs frame.

    Not a new calculation either. Dollars come from the same
    _build_portfolio_summary_record every layer uses, with the same aging
    source portfolio_summary_check uses -- the row's own servicer-reported
    dq_status. Counts come from that same builder over _unit_weight_frame,
    exactly as the Final Audited summary derives its counts, so the two
    summaries cannot drift apart in what they classify.

    Closed-at-zero suppression is an AUDIT judgement and is deliberately
    absent: a loan the audit would suppress still belongs in what the lender
    submitted, which is what OriginalIgnoresClosedAtZeroTests already requires
    of PORT-01.

    Returns {"dollars": <record>, "counts": <record>}. "dollars" is equal to
    portfolio_summary_check's record for the same input -- asserted against
    the real June files by test_original_summary_final_format.py."""
    population = consolidate_cycle_postings(servicer_inputs_df)
    aging_source = lambda gp, row: row.get("dq_status", "")

    dollars = _build_portfolio_summary_record(population, servicer_name, aging_source)
    weighted = _build_portfolio_summary_record(
        _unit_weight_frame(population), servicer_name, aging_source
    )
    counts = {"Servicer": servicer_name or ""}
    for category in ["Outstanding"] + _PORTFOLIO_SUMMARY_CATEGORY_ORDER:
        counts[category] = int(round(weighted[f"{category} ($)"]))
    # Carried here too so the two summaries expose the same counts contract.
    # This population is the lender's submission alone -- no verified_overrides
    # and no closed-at-zero suppression, per this function's own guarantee -- so
    # the figure can legitimately differ from the Final Audited one for the same
    # servicer, exactly as Outstanding already can.
    counts[SERVICING_POPULATION] = _eligible_loan_count(population)
    return {"dollars": dollars, "counts": counts}


# SOUP deal identifiers follow SOUP's own Deal-column convention,
# "SOUP<year>_<series>" -- produced by
# file_readers.extract_deal_from_note_series from each SRF's "Note Series:"
# cell (e.g. "2025 FBC-1" -> "SOUP2025_1") and already used to scope POP-01.
# Verified against all eight June files: every servicer's declared deal is
# present in SOUP for that servicer.
_DEAL_YEAR_PATTERN = re.compile(r"^SOUP(\d{4})_(\d+)$")

# CLIENT-CONFIRMED (2026-08-10): a servicer whose SRF declares no usable Note
# Series is grouped HERE rather than dropped or guessed at. A missing deal is
# a source-data problem and must stay visible in the report.
UNASSIGNED_DEAL_LABEL = "Unassigned Total"


def _deal_year_and_series(deal):
    """(year, series) as ints, or None when the identifier is missing or
    malformed. Deliberately strict: no filename inference, no servicer-name
    lookup, no guessing at a year."""
    match = _DEAL_YEAR_PATTERN.match(str(deal or "").strip())
    return (int(match.group(1)), int(match.group(2))) if match else None


def _deal_total_label(deal, all_deals):
    """Display label for one deal's Total row.

    "SOUP2024_1" -> "2024 SOUP Deal Total". The raw identifier stays internal.

    If two deals ever share a year (e.g. SOUP2024_1 alongside SOUP2024_2) the
    year alone stops identifying the group, so the full identifier is shown
    instead -- ambiguity surfaces rather than two deals silently reading as
    the same label."""
    parsed = _deal_year_and_series(deal)
    if parsed is None:
        return UNASSIGNED_DEAL_LABEL
    year = parsed[0]
    same_year = [d for d in all_deals
                 if (_deal_year_and_series(d) or (None,))[0] == year]
    if len(same_year) > 1:
        return f"{str(deal).strip()} Total"
    return f"{year} SOUP Deal Total"


def _deal_sort_key(deal):
    """Deals ascend by year then series; an unassigned group sorts LAST so a
    source-data problem never displaces real reporting."""
    parsed = _deal_year_and_series(deal)
    return (1, 9999, 9999) if parsed is None else (0, parsed[0], parsed[1])


def group_portfolio_summary_by_deal(dollar_records, count_records, deals):
    """CLIENT-CONFIRMED (2026-08-10): Totals are reported PER SOUP DEAL. There
    is deliberately NO combined total across deals -- a 2024-deal servicer and
    a 2025-deal servicer belong to different securitizations and their
    balances are not meaningfully additive.

    dollar_records / count_records / deals are parallel lists, one entry per
    servicer in upload order; deals[i] is that servicer's SOUP deal identifier
    from file_readers.extract_deal_from_note_series, or None when its SRF
    declared none.

    Returns blocks in display order (2024, then 2025, then any unassigned):

        [{"deal": "SOUP2024_1", "label": "2024 SOUP Deal Total",
          "totals": {"dollars": {...}, "counts": {...}},
          "dollars": [<servicer records>], "counts": [<servicer records>]}]

    Each block's totals come from portfolio_summary_totals over THAT block's
    servicers only, so every percentage is recomputed against that deal's own
    Outstanding -- which is why each deal reads Outstanding = 100.00%.
    Servicers keep upload order within their block.

    A GROUPING AND AGGREGATION CHANGE ONLY. No category rule, no
    classification, no population construction and no per-servicer figure is
    touched: the same records that were previously summed into one total are
    now summed into two."""
    grouped = {}
    for index, deal in enumerate(deals):
        key = str(deal).strip() if deal is not None and str(deal).strip() else None
        grouped.setdefault(key, []).append(index)

    ordered_keys = sorted(grouped, key=_deal_sort_key)
    blocks = []
    for key in ordered_keys:
        indices = grouped[key]
        block_dollars = [dollar_records[i] for i in indices]
        block_counts = [count_records[i] for i in indices]
        label = _deal_total_label(key, ordered_keys)
        blocks.append({
            "deal": key,
            "label": label,
            "totals": portfolio_summary_totals(block_dollars, block_counts, label=label),
            "dollars": block_dollars,
            "counts": block_counts,
        })
    return blocks


def portfolio_summary_totals(dollar_records, count_records, label="Total"):
    """The Total row across servicers.

    Counts and dollars are SUMMED -- each servicer's population is disjoint,
    so summation is exact. Categories still overlap with each other, so the
    Total row does not reconcile internally any more than a servicer row does.

    Percentages are RECOMPUTED from the aggregate dollars, never averaged.
    CLIENT-CONFIRMED (2026-08-10): averaging servicer percentages weights a
    23-loan servicer equally with a 1,234-loan one. Measured on the June set,
    that error reaches 0.91 points on Current and 0.83 on Not Making Payment.

    Zero aggregate Outstanding follows the same convention _build_portfolio_
    summary_record uses for a servicer: 0.00% everywhere, including
    Outstanding itself."""
    categories = ["Outstanding"] + _PORTFOLIO_SUMMARY_CATEGORY_ORDER
    dollars = {c: round(sum(r.get(f"{c} ($)", 0.0) or 0.0 for r in dollar_records), 2)
               for c in categories}
    counts = {c: sum(int(r.get(c, 0) or 0) for r in count_records) for c in categories}

    outstanding = dollars["Outstanding"]
    total_dollars = {"Servicer": label}
    for c in categories:
        total_dollars[f"{c} ($)"] = dollars[c]
        if not outstanding:
            total_dollars[f"{c} (%)"] = 0.0
        elif c == "Outstanding":
            total_dollars["Outstanding (%)"] = 100.0
        else:
            total_dollars[f"{c} (%)"] = round(dollars[c] / outstanding * 100, 2)

    total_counts = {"Servicer": label}
    total_counts.update(counts)
    # Summed like any other count: each servicer's population is disjoint, so
    # the deal's servicing population is exactly the sum of its servicers'.
    # Absent from a record (e.g. a caller that built counts before this key
    # existed) contributes 0 rather than raising.
    total_counts[SERVICING_POPULATION] = sum(
        int(r.get(SERVICING_POPULATION, 0) or 0) for r in count_records
    )
    return {"dollars": total_dollars, "counts": total_counts}


def verified_portfolio_summary_record(servicer_inputs_df, verified_overrides,
                                      servicer_name=None, excluded_gps=None):
    """CLIENT-CONFIRMED (2026-08-10): the FINAL AUDITED Portfolio Summary
    record — the third and last of the three layers:

        Original Portfolio Summary        portfolio_summary_check -- what the
                                          lender reported
        Audited Portfolio Summary         audited_portfolio_summary_record --
                                          what the audit calculated
        Final Audited Portfolio Summary   THIS -- what the reviewer verified

    This is NOT a third calculation. It is PORT-01's calculation applied to
    post-review values, and the aging lambda below is deliberately IDENTICAL
    to portfolio_summary_check's — because after review the row's own values
    ARE the verified ones. Deliberately NOT audited_status_by_loan: that would
    reinstate the audit's recommendation on loans the reviewer REJECTED, which
    is precisely what the three-layer model exists to prevent.

    That identity is also what makes this reproduce the historical workflow.
    Before this feature the client downloaded a corrected workbook, re-uploaded
    it and re-ran PORT-01 on it; this removes the round trip while computing
    the same thing from the same values.

    excluded_gps carries the closed-at-zero suppression set. Confirmed against
    all eight banks' historical corrected (vF) workbooks: the ONLY rows the
    client ever removed by hand were exactly the closed-at-zero loans, and
    every one of them was removed (Community Bank 4338609104; Peoples
    1271409104 / 1328729108 / 1430308405 / 3179359109). Suppression mechanises
    that practice rather than diverging from it — and since the new workflow
    cannot delete rows, it is the only mechanism that reproduces it.

    Returns the raw record dict directly, not wrapped in _result()."""
    return verified_portfolio_summary_views(
        servicer_inputs_df, verified_overrides, servicer_name, excluded_gps
    )["dollars"]


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


# ---------------------------------------------------------------------------
# DQ-01 — Final Posting Incomplete
# CLIENT-CONFIRMED (2026-08-06): surfaces every loan where
# consolidate_cycle_postings' incomplete-final-posting fallback actually
# engaged -- i.e. the posting flagged as this cycle's final/authoritative
# one (chronologically last, or explicitly marked "Last Payment Instance?
# = Y") had no real Loan Status, Delinquency Status, or Ending Balance at
# all, and an earlier, complete posting had to be used instead. This is
# never silent: every downstream number derived from this loan this cycle
# (Outstanding, its delinquency bucket, Loan Status classification, any
# check that reads consolidate_cycle_postings) was affected by which row
# won, so this check exists specifically so a human sees that a
# substitution happened, not just its consequence.
#
# Deliberately does NOT re-derive the blank-detection logic -- it reads
# incomplete_final_posting/original_last_excel_row/effective_last_excel_row
# directly off the SAME consolidated record every other check uses, so
# this can never disagree with what actually happened during consolidation.
# ---------------------------------------------------------------------------

def final_posting_incomplete_check(servicer_inputs_df, servicer_name=None):
    records = []

    if "SBA GP Number" not in servicer_inputs_df.columns:
        return _result("DQ-01", "Final Posting Incomplete", "SBA Reconciliation", "WARN", records,
                        lambda n: "")

    si_grouped = consolidate_cycle_postings(servicer_inputs_df)

    for gp, row in si_grouped.iterrows():
        if not row.get("incomplete_final_posting"):
            continue

        records.append({
            "SBA GP Number": gp,
            "Servicer": servicer_name or "",
            "Loan Status Used": row.get("loan_status", ""),
            "Delinquency Status Used": row.get("dq_status", ""),
            "Ending Balance Used": round(float(row["ending_balance"]), 2) if pd.notna(row.get("ending_balance")) else "",
            "Original Last Posting Excel Row": row.get("original_last_excel_row"),
            "Posting Used Instead (Excel Row)": row.get("effective_last_excel_row"),
            "Detail": (
                "The posting marked as the final payment instance contained blank Loan Status, "
                "Delinquency Status, and Ending Balance. The last complete posting was used for "
                "portfolio reporting. Servicer correction is recommended."
            ),
        })

    records.sort(key=lambda r: r["SBA GP Number"])

    return _result(
        "DQ-01", "Final Posting Incomplete", "SBA Reconciliation", "WARN", records,
        lambda n: f"{n} loan(s) had a blank final posting this cycle (no Loan Status, Delinquency "
                  f"Status, or Ending Balance at all) — the last complete posting was used instead "
                  f"for every downstream total. Servicer correction of the source file is recommended.",
    )


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
        # CLIENT-CONFIRMED (2026-08-14): the SAME cycle-opening rule
        # consolidate_cycle_postings uses, so DUP-01 and BAL-01 can never
        # report different Beginning Balances for the same loan. This is the
        # one value DUP-01 shares with the audit; the posting sequence above
        # (and therefore Ending Balance and the payment list) is unchanged.
        opening = _cycle_opening_posting(group)
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
            "Beginning Balance": opening.get("Beginning Balance", ""),
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
    ("CONT-01", "Prior Period Loan Continuity"),
    ("BAL-01", "Balance Continuity"),
    ("BAL-02", "Liquidation Balance Validation"),
    ("STAT-01", "Status / Delinquency Match"),
    ("STAT-02", "Terminal Status Review"),
    ("PAY-01", "Missing Payment Periods"),
    ("PAY-02", "Payment Shortfall Check"),
    ("DUP-01", "Duplicate Postings"),
    ("NEG-01", "Negative Payment Flags"),
    ("PORT-01", "Original Portfolio Summary"),
    ("DQ-01", "Final Posting Incomplete"),
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

    deal (from file_readers.extract_deal_from_note_series) scopes POP-01 and
    CONT-01 to Servicer + Deal + normalized GP Number instead of Servicer + GP
    Number alone — see population_check's docstring-level comment for why.
    Affects only those two; BAL-01, BAL-02, STAT-01, STAT-02, PAY-01/02, DUP-01, NEG-01, and PORT-01
    are unchanged by this parameter."""
    # CLOSED-AT-ZERO ELIGIBILITY, resolved ONCE per servicer and shared by
    # every client-facing exception check. Computed here rather than inside
    # each check so all four apply an identical rule to an identical set --
    # the alternative, four independent derivations, is exactly the drift
    # consolidate_cycle_postings exists to prevent.
    closed_at_zero = loans_closed_at_zero(soup_servicer_df)

    results = [
        population_check(loan_terms_df, soup_servicer_df, servicer_name, deal=deal),
        prior_period_continuity_check(
            servicer_inputs_df, soup_servicer_df, loan_terms_df=loan_terms_df,
            servicer_name=servicer_name, deal=deal, closed_at_zero=closed_at_zero,
        ),
        balance_continuity_check(servicer_inputs_df, soup_servicer_df, closed_at_zero=closed_at_zero),
        liquidation_balance_check(servicer_inputs_df, closed_at_zero=closed_at_zero),
        status_match_check(servicer_inputs_df, soup_servicer_df, memo_df=memo_df, servicer_name=memo_lender_key or servicer_name, closed_at_zero=closed_at_zero),
        terminal_status_review_check(servicer_inputs_df, soup_servicer_df, servicer_name=servicer_name, closed_at_zero=closed_at_zero),
        missing_payment_periods_check(soup_servicer_df, loan_terms_df),
        payment_shortfall_check(servicer_inputs_df),
        duplicate_postings_check(servicer_inputs_df),
        negative_payment_check(servicer_inputs_df, soup_servicer_df, memo_df=memo_df, servicer_name=memo_lender_key or servicer_name),
        portfolio_summary_check(servicer_inputs_df, servicer_name=servicer_name),
        final_posting_incomplete_check(servicer_inputs_df, servicer_name=servicer_name),
    ]
    # tag every result with which servicer it belongs to, so multi-servicer
    # runs can be told apart in the combined results list -- always the
    # display-facing servicer_name (e.g. "Millennium 2025"), never the memo
    # lookup key, so report tabs/summary still distinguish note series.
    for r in results:
        r["servicer"] = servicer_name
        r["id"] = f"{r['id']}::{servicer_name}"

    # ---- suppression trace -------------------------------------------------
    # One entry per suppressed loan, carrying the evidence that justified the
    # suppression plus which checks it was withheld from. Attached to STAT-01
    # as a sibling key -- the same pattern correction_data/audited_status_by_loan
    # already use -- so main.py can log it and strip it before the response.
    # Suppression must be auditable; silent exclusion is what this whole
    # session's memo-coverage work was about.
    suppressed_by_gp = {}
    for r in results:
        for gp in r.pop("suppressed_closed_gps", []) or []:
            suppressed_by_gp.setdefault(gp, []).append(r["id"].split("::")[0])
    trace = []
    if suppressed_by_gp:
        srf = consolidate_cycle_postings(servicer_inputs_df)
        for gp, checks_suppressed in sorted(suppressed_by_gp.items()):
            evidence = closed_at_zero.get(gp, {})
            row = srf.loc[gp] if gp in srf.index else None
            trace.append({
                "SBA GP Number": gp,
                "Servicer": servicer_name,
                "Prior Closed Status": evidence.get("closed_status", ""),
                "Prior Closed Month": evidence.get("closed_month", ""),
                "Prior Closed Deal": evidence.get("closed_deal", ""),
                "Prior Ending Actual Balance": evidence.get("closed_balance"),
                "Current Servicer Loan Status": str(row.get("loan_status", "")).strip() if row is not None else "",
                "Current Servicer Ending Balance": (
                    round(float(row.get("ending_balance")), 2)
                    if row is not None and row.get("ending_balance") is not None
                    and pd.notna(row.get("ending_balance")) else None
                ),
                "Checks Suppressed": sorted(checks_suppressed),
            })
    stat01 = next((r for r in results if r["id"].split("::")[0] == "STAT-01"), None)
    if stat01 is not None:
        stat01["suppressed_closed_loans"] = trace

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
