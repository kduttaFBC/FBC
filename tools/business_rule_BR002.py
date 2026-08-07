#!/usr/bin/env python3
"""
MONTHLY PAYMENTS MADE  --  BUSINESS RULE BR-002   (loan / reporting period)
==========================================================================

Rule ID       : BR-002        (BR = approved business rule)
Rule version  : 1.0
Formerly      : PR-002 v2.0 proposed; PR-001 v1.0 withdrawn
Applies to    : SRF Servicer Inputs, column AE "Monthly Payments Made?"

APPROVAL STATUS -- PENDING
    The rule is implemented and tested. It is NOT marked approved, because
    approval requires a reference to the approval and APPROVAL[] is empty.
    Populate APPROVAL from the client's written confirmation. See the
    APPROVAL GATE below for what the confirmation must cover.

    The question to put to the client is NOT "may we deploy 0.80". The
    audited summaries are already fully consistent with 0.80 and
    inconsistent with 0.95 for the June dataset, so the question is:

      "The application's audited results are consistent with an 80% payment
       qualification threshold, while the workbook formula uses 95%. Please
       confirm the intended business rule is 80%, and that the workbook
       formula should be updated to match."


THE RULE -- ONE DECISION PER LOAN PER REPORTING PERIOD

    Adjusted Payment Total = SUM(Payment Received over all postings)
                           + SUM(Servicing Fee    over all postings)
    Payment Units          = Adjusted Payment Total / Scheduled Payment
    Monthly Payments Made  = whole units + 1 when the fractional
                             remainder >= 0.80

    Postings are never evaluated independently. A loan with four postings
    of 30% each makes one payment, not zero.

    The result is an integer count. 0.80 units -> 1, 1.79 -> 1, 1.80 -> 2,
    5.80 -> 6.


DOWNSTREAM CHAIN

    postings -> aggregate -> PR-002 count -> delinquency movement (AG)
             -> new delinquency status (AH) -> final loan status
             -> portfolio summary

    PR-002 is not wired into the portfolio summary. portfolio_summary()
    consumes recalculated loan statuses, never payment counts, and refuses
    to produce totals when any loan in the set is undecidable.


DELIBERATE DIVERGENCE FROM THE TEMPLATE

    The workbook's column AE qualifies the fraction at 0.95:

        =IF(D12<>"",INT(AD12/R12)
            +IF(ROUND(AD12/R12-INT(AD12/R12),10)>=0.95,1,0),"")

    PR-002 uses 0.80. Until the template is updated to:

        =IF(D12<>"",INT(AD12/R12)
            +IF(ROUND(AD12/R12-INT(AD12/R12),10)>=0.80,1,0),"")

    the application will intentionally report a different count, and so a
    different delinquency status, for any loan whose fractional remainder
    falls in [0.80, 0.95). Every reconciliation must treat this as expected.

    The outer IF(D12<>"","") guard is retained. A blank Delinquency Status
    yields no decision, exactly as the template yields "".


THE DELINQUENCY LOOKUP TABLE IS CORRUPT -- READ BEFORE DEPLOYING

    Column AH resolves the movement string through an EXTERNAL workbook:
    VLOOKUP(AG, [1]DQ_Lookup!$A$2:$B$64, 2, FALSE). That file was not
    supplied, but its contents are cached inside the .xlsx external-link
    part, and DQ_LOOKUP_CACHED below is that cache verbatim.

    Seven of its 60 keys are malformed. The day count has been autofilled
    to follow the "+ n" suffix:

        '130 DAYS + 3 Status'   should be '120 DAYS + 3 Status'
        '40 DAYS + 4 Status'    should be  '30 DAYS + 4 Status'
        '140 DAYS + 4 Status'   should be '120 DAYS + 4 Status'
        '50 DAYS + 5 Status'    should be  '30 DAYS + 5 Status'
        '160 DAYS + 6 Status'   should be '120 DAYS + 6 Status'
        '70 DAYS + 7 Status'    should be  '30 DAYS + 7 Status'
        '170 DAYS + 7 Status'   should be '120 DAYS + 7 Status'

    Consequently ten legitimate status/count combinations are absent and
    return #N/A -- see MISSING_COMBINATIONS. Neither the June nor the July
    2026 submission reaches one, but LOWERING THE THRESHOLD RAISES COUNTS
    AND SO INCREASES EXPOSURE. July already contains a loan at
    'CURRENT + 5 Status'; the same loan at '30 DAYS' would have failed.

    This module does not repair the table. new_delinquency_status() fails
    closed on a missing key, and proposed_ladder_status() offers the
    first-principles answer clearly marked as unauthorised.
"""

from __future__ import annotations

__version__ = "3.0.0"

RULE_ID = "BR-002"                 # BR = approved business rule
RULE_VERSION = "1.0"
RULE_GRAIN = "loan-reporting-period"
FORMERLY = "PR-002 v2.0 (proposed); PR-001 v1.0 (withdrawn, boolean)"

# --------------------------------------------------------------------------
# APPROVAL GATE
#
# A business rule is approved when there is a reference to the approval, not
# when someone says it was approved. This module refuses to report itself as
# approved until every field below is populated, for the same reason the
# validator refuses to apply a $15.00 balance correction without written
# authorisation: an unevidenced adjustment of $711,164.92 in reclassified
# balance is not made safer by being a rule rather than a number.
#
# Populate from the actual confirmation, then commit.
# --------------------------------------------------------------------------
APPROVAL = {
    "status": "PENDING",           # PENDING | APPROVED
    "reference": None,             # e.g. "email, K. Dutta -> M. <client>, 2026-08-05"
    "approved_by": None,           # named individual with authority
    "approved_on": None,           # ISO date
    "scope_confirmed": None,       # must cover BOTH limbs below
}
APPROVAL_REQUIRED_SCOPE = (
    "(a) qualification is Payment Received + Servicing Fee >= 0.80 x "
    "Scheduled Payment, at loan and reporting-period grain; "
    "(b) Monthly Payments Made remains an INTEGER COUNT driving delinquency "
    "movement and portfolio summary values; "
    "(c) the workbook template formula is updated from 0.95 to 0.80."
)


def is_approved() -> bool:
    return (APPROVAL["status"] == "APPROVED"
            and all(APPROVAL[k] for k in
                    ("reference", "approved_by", "approved_on",
                     "scope_confirmed")))


def approval_state() -> str:
    if is_approved():
        return (f"APPROVED per {APPROVAL['reference']} "
                f"({APPROVAL['approved_by']}, {APPROVAL['approved_on']})")
    missing = [k for k in ("status", "reference", "approved_by",
                           "approved_on", "scope_confirmed")
               if not APPROVAL[k] or APPROVAL[k] == "PENDING"]
    return f"NOT APPROVED -- missing: {', '.join(missing)}"

RULE_THRESHOLD = "0.80"
TEMPLATE_THRESHOLD = "0.95"   # superseded once TEMPLATE_ALIGNMENT is ALIGNED

# --------------------------------------------------------------------------
# TEMPLATE ALIGNMENT -- time-bounded, not permanent design
#
# The divergence from the workbook formula is a transitional state, not an
# architectural choice. Once the template formula is officially changed to
# 0.80, set status to "ALIGNED" with the reference. divergence_note() then
# reports the divergence in the past tense, and the section below becomes
# historical record rather than a description of current behaviour.
#
# Leaving this PENDING indefinitely would let a temporary condition harden
# into apparent design intent.
# --------------------------------------------------------------------------
TEMPLATE_ALIGNMENT = {
    "status": "PENDING",       # PENDING | ALIGNED
    "template_version": None,  # version in which the formula became 0.80
    "aligned_on": None,
    "reference": None,
}


def is_template_aligned() -> bool:
    return (TEMPLATE_ALIGNMENT["status"] == "ALIGNED"
            and bool(TEMPLATE_ALIGNMENT["reference"]))


def divergence_note() -> str:
    if is_template_aligned():
        return (f"HISTORICAL. The workbook formula qualified at "
                f"{TEMPLATE_THRESHOLD} until template "
                f"{TEMPLATE_ALIGNMENT['template_version']} aligned it to "
                f"{RULE_THRESHOLD} on {TEMPLATE_ALIGNMENT['aligned_on']} "
                f"({TEMPLATE_ALIGNMENT['reference']}). No divergence "
                f"remains; reconciliations should agree.")
    return (f"CURRENT AND TRANSITIONAL. The workbook formula qualifies at "
            f"{TEMPLATE_THRESHOLD}; this rule qualifies at {RULE_THRESHOLD}. "
            f"Every reconciliation between the application and the workbook "
            f"must treat the difference as expected for any loan whose "
            f"fractional unit falls in [{RULE_THRESHOLD}, "
            f"{TEMPLATE_THRESHOLD}). This state is transitional and must not "
            f"be read as design intent -- set TEMPLATE_ALIGNMENT when the "
            f"template formula is corrected.")

EXCEL_EQUIVALENT = (
    '=IF(D{row}<>"",INT(AD{row}/R{row})'
    '+IF(ROUND(AD{row}/R{row}-INT(AD{row}/R{row}),10)>=0.80,1,0),"")'
)

import argparse
import sys
from decimal import Decimal, ROUND_FLOOR, ROUND_HALF_UP, InvalidOperation

THRESHOLD = Decimal(RULE_THRESHOLD)
CENT = Decimal("0.01")
TEN_DP = Decimal("1E-10")

ABSENT = "absent"
CORRUPT = "corrupt"
OK = "ok"

STATUS_LADDER = ["CURRENT", "DELINQUENT", "30 DAYS", "60 DAYS",
                 "90 DAYS", "120 DAYS", "150 DAYS"]

# Verbatim from xl/externalLinks/externalLink1.xml, sheetId 6 (DQ_Lookup),
# cached inside the June 2026 Peoples Bank submission. Reproduced as found,
# including the seven malformed keys.
DQ_LOOKUP_CACHED = {
    'Current': 'Current', 'DELINQUENT': 'DELINQUENT', '30 DAYS': '30 DAYS',
    '60 DAYS': '60 DAYS', '90 DAYS': '90 DAYS', '120 DAYS': '120 DAYS',
    '150 DAYS': '150 DAYS',
    'Current - 1 Status': 'DELINQUENT', 'DELINQUENT - 1 Status': '30 DAYS',
    '30 DAYS - 1 Status': '60 DAYS', '60 DAYS - 1 Status': '90 DAYS',
    '90 DAYS - 1 Status': '120 DAYS', '120 DAYS - 1 Status': '150 DAYS',
    '150 DAYS - 1 Status': '150 DAYS',
    'CURRENT + 1 Status': 'Current', 'DELINQUENT + 1 Status': 'Current',
    '30 DAYS + 1 Status': 'DELINQUENT', '60 DAYS + 1 Status': '30 DAYS',
    '90 DAYS + 1 Status': '60 DAYS', '120 DAYS + 1 Status': '90 DAYS',
    '150 DAYS + 1 Status': '120 DAYS',
    'CURRENT + 2 Status': 'Current', 'DELINQUENT + 2 Status': 'Current',
    '30 DAYS + 2 Status': 'Current', '60 DAYS + 2 Status': 'DELINQUENT',
    '90 DAYS + 2 Status': '30 DAYS', '120 DAYS + 2 Status': '60 DAYS',
    '150 DAYS + 2 Status': '90 DAYS',
    'CURRENT + 3 Status': 'Current', 'DELINQUENT + 3 Status': 'Current',
    '30 DAYS + 3 Status': 'Current', '60 DAYS + 3 Status': 'Current',
    '90 DAYS + 3 Status': 'DELINQUENT',
    '130 DAYS + 3 Status': '30 DAYS',          # malformed: should be 120 DAYS
    '150 DAYS + 3 Status': '60 DAYS',
    'CURRENT + 4 Status': 'Current', 'DELINQUENT + 4 Status': 'Current',
    '40 DAYS + 4 Status': 'Current',           # malformed: should be 30 DAYS
    '60 DAYS + 4 Status': 'Current', '90 DAYS + 4 Status': 'Current',
    '140 DAYS + 4 Status': 'DELINQUENT',       # malformed: should be 120 DAYS
    '150 DAYS + 4 Status': '30 DAYS',
    'CURRENT + 5 Status': 'Current', 'DELINQUENT + 5 Status': 'Current',
    '50 DAYS + 5 Status': 'Current',           # malformed: should be 30 DAYS
    '60 DAYS + 5 Status': 'Current', '90 DAYS + 5 Status': 'Current',
    '150 DAYS + 5 Status': '30 DAYS',
    'CURRENT + 6 Status': 'Current', 'DELINQUENT + 6 Status': 'Current',
    '60 DAYS + 6 Status': 'Current', '90 DAYS + 6 Status': 'Current',
    '160 DAYS + 6 Status': 'Current',          # malformed: should be 120 DAYS
    '150 DAYS + 6 Status': 'DELINQUENT',
    'CURRENT + 7 Status': 'Current', 'DELINQUENT + 7 Status': 'Current',
    '70 DAYS + 7 Status': 'Current',           # malformed: should be 30 DAYS
    '90 DAYS + 7 Status': 'Current',
    '170 DAYS + 7 Status': 'Current',          # malformed: should be 120 DAYS
    '150 DAYS + 7 Status': 'Current',
}

_DQ = {k.upper(): v for k, v in DQ_LOOKUP_CACHED.items()}

MALFORMED_KEYS = [k for k in DQ_LOOKUP_CACHED
                  if k.upper().split(" + ")[0].split(" - ")[0]
                  not in STATUS_LADDER]

MISSING_COMBINATIONS = [
    (s, n) for n in range(1, 8) for s in STATUS_LADDER
    if f"{s} + {n} STATUS" not in _DQ
]


# --------------------------------------------------------------------------
# release gate
# --------------------------------------------------------------------------

def is_release_authorised() -> bool:
    """Alias retained for callers; approval is the single gate."""
    return is_approved()


# --------------------------------------------------------------------------
# coercion
# --------------------------------------------------------------------------

def to_money(v):
    """
    Returns (Decimal, OK) | (None, ABSENT) | (None, CORRUPT).

    ABSENT means nothing was reported; a blank servicing fee legitimately
    means zero. CORRUPT means a value is present and is not a number -- the
    literal 'Y' in June row 59, for instance. Coercing that to zero would
    report a missing payment on the strength of a typo, so it fails closed.
    """
    if v is None or isinstance(v, bool):
        return None, ABSENT
    if isinstance(v, str):
        t = v.strip().replace(",", "").replace("$", "")
        if t == "":
            return None, ABSENT
        try:
            return Decimal(t).quantize(CENT, rounding=ROUND_HALF_UP), OK
        except InvalidOperation:
            return None, CORRUPT
    if isinstance(v, (int, float, Decimal)):
        try:
            return Decimal(str(v)).quantize(CENT, rounding=ROUND_HALF_UP), OK
        except InvalidOperation:
            return None, CORRUPT
    return None, CORRUPT


def aggregate_postings(postings):
    """
    Sum Payment Received and Servicing Fee across a loan's postings.

    postings: iterable of (payment_received, servicing_fee).

    Original values are never modified. Returns the two totals, per-component
    status, and the raw inputs of any corrupt posting so the audit record can
    show what was actually reported.
    """
    pr_total, sf_total = Decimal("0.00"), Decimal("0.00")
    pr_status, sf_status = OK, OK
    corrupt = []
    n = 0
    for i, (pr_raw, sf_raw) in enumerate(postings):
        n += 1
        d, st = to_money(pr_raw)
        if st == CORRUPT:
            pr_status = CORRUPT
            corrupt.append({"posting_index": i, "field": "payment_received",
                            "raw": repr(pr_raw)})
        elif d is not None:
            pr_total += d
        d, st = to_money(sf_raw)
        if st == CORRUPT:
            sf_status = CORRUPT
            corrupt.append({"posting_index": i, "field": "servicing_fee",
                            "raw": repr(sf_raw)})
        elif d is not None:
            sf_total += d
    return {
        "postings_aggregated": n,
        "payment_received_total": pr_total,
        "servicing_fee_total": sf_total,
        "payment_received_status": pr_status,
        "servicing_fee_status": sf_status,
        "corrupt_inputs": corrupt,
    }


# --------------------------------------------------------------------------
# PR-002
# --------------------------------------------------------------------------

def monthly_payments_made(adjusted_payment_total, scheduled_payment) -> int:
    """
    Apply PR-002 to a loan's aggregated total. Returns an integer count.

    Raises ValueError when the inputs cannot support a decision. Callers
    processing a batch must catch per loan; use recalculate_loan() for the
    non-raising form.
    """
    adjusted, ad_st = to_money(adjusted_payment_total)
    sched, sc_st = to_money(scheduled_payment)
    if ad_st == CORRUPT:
        raise ValueError("Adjusted payment total is not numeric.")
    if adjusted is None:
        adjusted = Decimal("0.00")
    if sc_st == CORRUPT:
        raise ValueError(
            f"Scheduled payment is not numeric ({scheduled_payment!r}).")
    if sched is None:
        raise ValueError("Scheduled payment is missing; it must be > 0.")
    if sched <= 0:
        raise ValueError(f"Scheduled payment must be > 0, got {sched}.")
    if adjusted < 0:
        raise ValueError(
            f"Adjusted payment total is negative ({adjusted}); flooring a "
            "negative quotient yields a negative count, which the "
            "delinquency logic cannot interpret.")

    units = adjusted / sched
    whole = int(units.to_integral_value(rounding=ROUND_FLOOR))
    fraction = (units - Decimal(whole)).quantize(
        TEN_DP, rounding=ROUND_HALF_UP)        # mirrors Excel ROUND(...,10)
    return whole + int(fraction >= THRESHOLD)


# --------------------------------------------------------------------------
# delinquency chain
# --------------------------------------------------------------------------

def delinquency_movement(prior_status, count):
    """Reproduce column AG. Returns the lookup key, or None."""
    if prior_status is None or str(prior_status).strip() == "":
        return None
    s = str(prior_status).strip()
    if count == 0:
        return f"{s} - 1 Status"
    if count == 1:
        return s
    return f"{s} + {count - 1} Status"


def new_delinquency_status(movement_key):
    """
    Reproduce column AH via the cached DQ_Lookup. Returns (status, None) or
    (None, reason). Fails closed on a missing key rather than guessing --
    Excel would return #N/A here.
    """
    if movement_key is None:
        return None, "no movement key (blank delinquency status)"
    v = _DQ.get(movement_key.upper())
    if v is None:
        return None, (f"'{movement_key}' is absent from DQ_Lookup; Excel "
                      f"returns #N/A. The cached table has {len(MALFORMED_KEYS)} "
                      f"malformed keys and {len(MISSING_COMBINATIONS)} missing "
                      f"combinations.")
    return v, None


def proposed_ladder_status(prior_status, count):
    """
    First-principles answer, derived from the status ladder rather than the
    corrupt table. UNAUTHORISED -- provided so the gap can be quantified and
    put to the business owner. Never call this in production.
    """
    if prior_status is None or str(prior_status).strip() == "":
        return None
    s = str(prior_status).strip().upper()
    if s not in STATUS_LADDER:
        return None
    i = STATUS_LADDER.index(s)
    j = i + 1 if count == 0 else max(0, i - (count - 1))
    return STATUS_LADDER[min(j, len(STATUS_LADDER) - 1)]


# --------------------------------------------------------------------------
# one loan, end to end
# --------------------------------------------------------------------------

def recalculate_loan(loan_id, postings, scheduled_payment,
                     delinquency_status, prior_delinquency_status) -> dict:
    """
    Aggregate a loan's postings, apply PR-002, then recalculate delinquency
    movement and final status. Non-raising: monthly_payments_made is None
    when no decision is possible, with the reason stated.

    Corrupt inputs are preserved as raw values with a status, never shown
    as 0.00.
    """
    agg = aggregate_postings(postings)
    pr_bad = agg["payment_received_status"] == CORRUPT
    sf_bad = agg["servicing_fee_status"] == CORRUPT
    adjusted = (agg["payment_received_total"]
                + agg["servicing_fee_total"]).quantize(CENT,
                                                       rounding=ROUND_HALF_UP)
    sched, sched_st = to_money(scheduled_payment)

    rec = {
        "loan_id": loan_id,
        "rule_id": RULE_ID,
        "rule_version": RULE_VERSION,
        "grain": RULE_GRAIN,
        "postings_aggregated": agg["postings_aggregated"],
        # corrupt components are reported as None with the raw value kept
        "payment_received": (None if pr_bad
                             else float(agg["payment_received_total"])),
        "payment_received_status": agg["payment_received_status"],
        "servicing_fee": (None if sf_bad
                          else float(agg["servicing_fee_total"])),
        "servicing_fee_status": agg["servicing_fee_status"],
        "corrupt_inputs": agg["corrupt_inputs"],
        "adjusted_payment_total": None if (pr_bad or sf_bad) else float(adjusted),
        "scheduled_payment": None if sched is None else float(sched),
        "scheduled_payment_status": sched_st,
        "payment_units": None,
        "whole_units": None,
        "fractional_unit": None,
        "threshold": float(THRESHOLD),
        "monthly_payments_made": None,
        "delinquency_status_reported": delinquency_status,
        "prior_delinquency_status": prior_delinquency_status,
        "delinquency_movement": None,
        "new_delinquency_status": None,
        "final_loan_status": None,
        "decidable": False,
        "reason": None,
    }

    if pr_bad or sf_bad:
        raw = ", ".join(f"{c['field']}={c['raw']}"
                        for c in agg["corrupt_inputs"])
        rec["reason"] = (f"corrupt input, cannot be treated as zero: {raw}")
        return rec

    # The template's outer guard: a blank Delinquency Status yields "".
    if delinquency_status is None or str(delinquency_status).strip() == "":
        rec["reason"] = ("Delinquency Status is blank; the template returns "
                         "an empty string and no count is produced")
        return rec

    try:
        count = monthly_payments_made(adjusted, scheduled_payment)
    except ValueError as e:
        rec["reason"] = str(e)
        return rec

    units = adjusted / sched
    whole = int(units.to_integral_value(rounding=ROUND_FLOOR))
    rec["payment_units"] = float(units)
    rec["whole_units"] = whole
    rec["fractional_unit"] = float(
        (units - Decimal(whole)).quantize(TEN_DP, rounding=ROUND_HALF_UP))
    rec["monthly_payments_made"] = count

    key = delinquency_movement(prior_delinquency_status, count)
    rec["delinquency_movement"] = key
    status, why = new_delinquency_status(key)
    if status is None:
        rec["reason"] = why
        return rec
    rec["new_delinquency_status"] = status
    rec["final_loan_status"] = status
    rec["decidable"] = True
    return rec


def portfolio_summary(loan_records, balances=None) -> dict:
    """
    Rebuild portfolio summaries from RECALCULATED LOAN STATUSES.

    Consumes final_loan_status, never payment counts -- PR-002 is not wired
    into the summary. Refuses to produce totals if any loan is undecidable,
    because a summary that silently omits loans is the failure mode that
    cost 650,128.27 on the June submission.
    """
    undecidable = [r["loan_id"] for r in loan_records if not r["decidable"]]
    if undecidable:
        return {"status": "REFUSED", "undecidable_loans": undecidable,
                "reason": "every loan must have a recalculated status before "
                          "a portfolio summary can be produced"}
    by_status = {}
    for r in loan_records:
        by_status[r["final_loan_status"]] = \
            by_status.get(r["final_loan_status"], 0) + 1
    out = {"status": "OK", "loan_count": len(loan_records),
           "loans_by_status": dict(sorted(by_status.items()))}
    if balances:
        tot = sum(Decimal(str(balances[r["loan_id"]]))
                  for r in loan_records if r["loan_id"] in balances)
        bal = {}
        for r in loan_records:
            if r["loan_id"] in balances:
                bal[r["final_loan_status"]] = bal.get(
                    r["final_loan_status"], Decimal("0")) \
                    + Decimal(str(balances[r["loan_id"]]))
        out["outstanding_balance"] = float(tot)
        out["balance_by_status"] = {k: float(v) for k, v in sorted(bal.items())}
        out["percentage_by_status"] = {
            k: (float(v / tot * 100) if tot else 0.0)
            for k, v in sorted(bal.items())}
    return out


def template_count(adjusted_payment_total, scheduled_payment,
                   threshold: str = TEMPLATE_THRESHOLD):
    """Reproduce column AE at a given threshold, for comparison."""
    ad, _ = to_money(adjusted_payment_total)
    r, _ = to_money(scheduled_payment)
    if ad is None or r is None or r <= 0:
        return None
    units = ad / r
    whole = int(units.to_integral_value(rounding=ROUND_FLOOR))
    fraction = (units - Decimal(whole)).quantize(TEN_DP,
                                                rounding=ROUND_HALF_UP)
    return whole + int(fraction >= Decimal(threshold))


# --------------------------------------------------------------------------
# regression suite
# --------------------------------------------------------------------------

# These EXERCISE THE THRESHOLD CHANGE. Neither the June nor the July file
# contains a fractional remainder in [0.80, 0.95), so without these cases the
# override is entirely untested.
THRESHOLD_EXERCISING = [
    # adjusted, scheduled, PR-002 count, template@0.95 count
    ("850.00",  "1000.00", 1, 0, "fraction 0.85 -- qualifies at .80 only"),
    ("1850.00", "1000.00", 2, 1, "fraction 0.85 on top of 1 whole unit"),
    ("5800.00", "1000.00", 6, 5,
     "fraction 0.80 -- 6 at .80 vs 5 at .95, count preserved either way"),
    ("940.00",  "1000.00", 1, 0, "fraction 0.94 -- just under .95"),
    ("950.00",  "1000.00", 1, 1, "fraction 0.95 -- both qualify"),
    ("2850.00", "1000.00", 3, 2, "fraction 0.85 on top of 2 whole units"),
]

SPECIFIED_UNITS = [
    ("560.80",  "701.00",  1, "units 0.80 exactly"),
    ("560.79",  "701.00",  0, "units just below 0.80"),
    ("800.00",  "1000.00", 1, "units 0.80"),
    ("1790.00", "1000.00", 1, "units 1.79"),
    ("1800.00", "1000.00", 2, "units 1.80"),
    ("5800.00", "1000.00", 6, "units 5.80  COUNT PRESERVED"),
]

# (payment_received, servicing_fee) per posting
MULTI_POSTING = [
    # postings, scheduled, expected count, label
    ([("300.00", "0.00"), ("300.00", "0.00"), ("250.00", "0.00")],
     "1000.00", 1,
     "no single posting reaches 80%; 30+30+25 = 85% combined -> 1"),
    ([("200.00", "0.00"), ("200.00", "0.00"), ("200.00", "0.00"),
      ("200.00", "0.00")],
     "1000.00", 1, "four postings of 20% each = 80% -> 1"),
    ([("250.00", "20.00"), ("250.00", "20.00"), ("250.00", "20.00")],
     "1000.00", 1,
     "payments alone are 75%; fees of 60.00 carry it to 81% -> 1"),
    ([("100.00", "0.00"), ("100.00", "0.00")],
     "1000.00", 0, "20% combined -> 0, advances delinquency"),
    ([("1000.00", "0.00"), ("850.00", "0.00")],
     "1000.00", 2, "1.85 units across two postings -> 2"),
    ([("990.00", "10.00")], "1000.00", 1, "single posting, fee completes it"),
]


def self_test() -> int:
    fails = 0
    print("=" * 74)
    print(f"REGRESSION SUITE -- {RULE_ID} v{RULE_VERSION}  grain={RULE_GRAIN}")
    print(f"threshold {RULE_THRESHOLD} (template {TEMPLATE_THRESHOLD})")
    print("=" * 74)

    print("\nTHRESHOLD-EXERCISING SYNTHETIC CASES")
    print(f"  {'adjusted':>10} {'sched':>10} {'fraction':>9} "
          f"{'PR002':>6} {'tmpl':>5} {'exp':>4}{'':>3}result")
    for ad, sch, exp, exp_t, label in THRESHOLD_EXERCISING:
        got = monthly_payments_made(ad, sch)
        got_t = template_count(ad, sch)
        u = to_money(ad)[0] / to_money(sch)[0]
        fr = float(u - int(u.to_integral_value(rounding=ROUND_FLOOR)))
        ok = got == exp and got_t == exp_t
        fails += 0 if ok else 1
        print(f"  {ad:>10} {sch:>10} {fr:>9.4f} {got:>6} {got_t:>5} "
              f"{exp:>4}   {'PASS' if ok else 'FAIL'}  {label}")

    print("\nSPECIFIED UNIT CASES")
    for ad, sch, exp, label in SPECIFIED_UNITS:
        got = monthly_payments_made(ad, sch)
        ok = got == exp
        fails += 0 if ok else 1
        print(f"  {ad:>10} / {sch:<10} -> {got:>3} (exp {exp:>3})  "
              f"{'PASS' if ok else 'FAIL'}  {label}")

    print("\nMULTI-POSTING AGGREGATION  (never evaluated independently)")
    for postings, sch, exp, label in MULTI_POSTING:
        r = recalculate_loan("TEST", postings, sch, "CURRENT", "CURRENT")
        got = r["monthly_payments_made"]
        each = [monthly_payments_made(
            (to_money(p)[0] + to_money(f)[0]), sch) for p, f in postings]
        ok = got == exp
        fails += 0 if ok else 1
        print(f"  {r['postings_aggregated']} postings, adjusted "
              f"{r['adjusted_payment_total']:>9,.2f} / {sch:<9} -> {got} "
              f"(exp {exp})  {'PASS' if ok else 'FAIL'}")
        print(f"      per-posting counts would be {each} -> summing those "
              f"gives {sum(each)}")
        print(f"      {label}")

    print("\nRETURN TYPE IS AN INTEGER COUNT")
    v = monthly_payments_made("5800.00", "1000.00")
    for label, ok in [("int and not bool",
                       isinstance(v, int) and not isinstance(v, bool)),
                      ("value 6, not True", v == 6)]:
        fails += 0 if ok else 1
        print(f"  {label:<40} {'PASS' if ok else 'FAIL'}")

    print("\nCORRUPT INPUT IS PRESERVED, NOT SHOWN AS 0.00")
    r = recalculate_loan("L1", [("Y", "10.00")], "532.41", "CURRENT",
                         "CURRENT")
    checks = [("payment_received is None", r["payment_received"] is None),
              ("status is 'corrupt'",
               r["payment_received_status"] == CORRUPT),
              ("raw value retained",
               r["corrupt_inputs"] and r["corrupt_inputs"][0]["raw"] == "'Y'"),
              ("count is None", r["monthly_payments_made"] is None),
              ("adjusted total not fabricated",
               r["adjusted_payment_total"] is None),
              ("reason stated", bool(r["reason"]))]
    for label, ok in checks:
        fails += 0 if ok else 1
        print(f"  {label:<40} {'PASS' if ok else 'FAIL'}")

    print("\nBLANK DELINQUENCY STATUS GUARD  (template returns \"\")")
    for ds, expect_none in [(None, True), ("", True), ("   ", True),
                            ("CURRENT", False)]:
        r = recalculate_loan("L2", [("900.00", "0.00")], "1000.00", ds,
                             "CURRENT")
        got_none = r["monthly_payments_made"] is None
        ok = got_none == expect_none
        fails += 0 if ok else 1
        print(f"  delinquency_status={ds!r:<10} count="
              f"{str(r['monthly_payments_made']):<5} "
              f"{'PASS' if ok else 'FAIL'}")

    print("\nDELINQUENCY CHAIN")
    for prior, count, exp_key, exp_status in [
            ("CURRENT", 0, "CURRENT - 1 Status", "DELINQUENT"),
            ("CURRENT", 1, "CURRENT", "Current"),
            ("CURRENT", 2, "CURRENT + 1 Status", "Current"),
            ("30 DAYS", 2, "30 DAYS + 1 Status", "DELINQUENT"),
            ("CURRENT", 6, "CURRENT + 5 Status", "Current")]:
        key = delinquency_movement(prior, count)
        st, why = new_delinquency_status(key)
        ok = key == exp_key and st == exp_status
        fails += 0 if ok else 1
        print(f"  {prior:<11} count={count} -> {str(key):<22} -> "
              f"{str(st):<12} {'PASS' if ok else 'FAIL'}")

    print("\nCORRUPT LOOKUP TABLE FAILS CLOSED")
    for prior, count in [("30 DAYS", 5), ("120 DAYS", 4), ("30 DAYS", 7)]:
        key = delinquency_movement(prior, count)
        st, why = new_delinquency_status(key)
        proposed = proposed_ladder_status(prior, count)
        ok = st is None and why is not None
        fails += 0 if ok else 1
        print(f"  {prior:<11} count={count} -> {key:<22} #N/A, "
              f"unauthorised proposal would be {proposed:<11} "
              f"{'PASS' if ok else 'FAIL'}")

    print("\nPORTFOLIO SUMMARY CONSUMES STATUS, REFUSES ON UNDECIDABLE")
    good = [recalculate_loan("A", [("900.00", "0.00")], "1000.00", "CURRENT",
                             "CURRENT"),
            recalculate_loan("B", [("100.00", "0.00")], "1000.00", "CURRENT",
                             "CURRENT")]
    s = portfolio_summary(good, balances={"A": "100000.00", "B": "50000.00"})
    ok = s["status"] == "OK" and s["loan_count"] == 2
    fails += 0 if ok else 1
    print(f"  all decidable -> {s['status']}, by status "
          f"{s.get('loans_by_status')}  {'PASS' if ok else 'FAIL'}")
    bad = good + [recalculate_loan("C", [("Y", "0.00")], "1000.00", "CURRENT",
                                   "CURRENT")]
    s2 = portfolio_summary(bad)
    ok = s2["status"] == "REFUSED" and s2["undecidable_loans"] == ["C"]
    fails += 0 if ok else 1
    print(f"  one undecidable -> {s2['status']}, loans "
          f"{s2.get('undecidable_loans')}  {'PASS' if ok else 'FAIL'}")

    print("\nAPPROVAL GATE")
    ok = is_approved() is False
    fails += 0 if ok else 1
    print(f"  {approval_state()}")
    print(f"  is_approved()={is_approved()}  "
          f"{'PASS' if ok else 'FAIL'}  (must be False until referenced)")

    print("\n" + "=" * 74)
    print("ALL TESTS PASSED" if not fails else f"{fails} FAILURE(S)")
    print("=" * 74)
    return 1 if fails else 0


def main() -> int:
    ap = argparse.ArgumentParser(
        description=f"{RULE_ID} v{RULE_VERSION} ({RULE_GRAIN})")
    ap.add_argument("--approval", action="store_true",
                    help="print the approval state and exit non-zero if "
                         "not approved")
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--lookup-audit", action="store_true",
                    help="report the corrupt DQ_Lookup table")
    ap.add_argument("--version", action="version",
                    version=f"business_rule {__version__} "
                            f"({RULE_ID} v{RULE_VERSION}, {RULE_GRAIN})")
    a = ap.parse_args()
    rc = 0
    if a.self_test:
        rc |= self_test()
    if a.lookup_audit:
        print("\n" + "=" * 74)
        print("DQ_LOOKUP AUDIT -- cached from the external workbook")
        print("=" * 74)
        print(f"entries              : {len(DQ_LOOKUP_CACHED)}")
        print(f"malformed keys       : {len(MALFORMED_KEYS)}")
        for k in MALFORMED_KEYS:
            print(f"    {k}")
        print(f"missing combinations : {len(MISSING_COMBINATIONS)}")
        for s, n in MISSING_COMBINATIONS:
            print(f"    '{s} + {n} Status'  (reached at a count of {n + 1})")
        print("\nLowering the threshold raises counts, so these become more")
        print("reachable. The table must be corrected in the external")
        print("workbook by its owner; this module will not guess.")
        print("=" * 74)
    if a.approval:
        print(f"{RULE_ID} v{RULE_VERSION}: {approval_state()}")
        print(f"template alignment: {divergence_note()}")
        print(f"required scope: {APPROVAL_REQUIRED_SCOPE}")
        rc |= (0 if is_approved() else 1)
    if not a.self_test and not a.lookup_audit and not a.approval:
        ap.print_help()
    return rc


if __name__ == "__main__":
    sys.exit(main())
