"""
Audit Intelligence — corrected servicer workbook generation (review-gated).

This module is the ONLY place that decides which audit findings are
eligible to correct, and the only place that ever writes to a copy of an
uploaded servicer workbook. It is deliberately isolated from the
audit/reconciliation logic in checks.py:

  - checks.py (status_match_check / STAT-01) computes raw facts about each
    loan — what the servicer reported, what the engine calculated, and which
    warning conditions apply — and exposes them on result["correction_data"].
    It makes NO decision about what's "correctable."
  - This module reads those raw facts and applies a strict, mechanical
    allow-list (build_eligible_corrections) to decide what is a CANDIDATE.
  - Nothing in this file ever changes an audit result, a report, or a
    reconciliation formula. A bug here cannot silently change what the audit
    engine concludes; a bug in the audit engine cannot silently widen what
    this module is willing to write.

WORKFLOW (review-then-generate — see BRIEF_correction_workflow_restructure):
  - The audit produces RECOMMENDATIONS only. build_recommendations() shapes
    the eligible candidates for the reviewer; NO workbook is written during
    /api/audit.
  - A reviewer then decides each recommendation (approve / reject / return),
    optionally substituting an approved value from that field's own domain.
  - generate_reviewed_workbook() opens a FRESH copy of the original uploaded
    bytes, writes ONLY the approved values, and records every decision —
    approved, rejected and returned — in the Correction Log.

Version 1 candidate scope (deliberately narrow — the eligibility allow-list):
  - Delinquency Status: a candidate only when STAT-01 produced a real,
    unambiguous calculated value and flagged no warning of any kind for that
    loan.
  - Loan Status: a candidate only for a CONFIRMED Full P&I deferral in its
    "start" or "mid" phase (calculated value DEFERRED). Partial (principal-
    only/interest-only) deferrals and the deferral "resume" phase are
    explicitly OUT of scope for v1 — see build_eligible_corrections for why.
  - Every other field (balances, payments, population, liquidation,
    negative payments, missing periods, duplicate postings, everything
    else) is untouched in this version.
  - The original uploaded workbook is never modified — a new workbook is
    always produced from a fresh copy of the original bytes.
  - The reviewer's approved value is domain-restricted (see
    validate_approved_value): a substitution can never reintroduce a
    cross-domain value. A reviewer MAY select any status in the client's
    authoritative Loan Status domain, including a terminal one -- but the
    APPLICATION never generates a terminal collection status itself, and no
    Loan Status is written without explicit reviewer approval.
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

REVIEW_REASON_TERMINAL_STATUS = (
    "The servicer reports a terminal collection status (FORECLOSURE / "
    "LIQUIDATION / REO) for this loan. That report is evidence to review, not "
    "a confirmed fact: it does not validate itself and it does not change the "
    "delinquency calculation. Requires explicit reviewer approval before "
    "anything is written."
)

# Carries the contradiction that checks.py's site-4 DEFERRAL DATA QUALITY
# WARNING used to state (superseded 2026-08-10). Same information, moved onto
# the control where the decision is made instead of a column that suppressed
# every correction on the loan.
REVIEW_REASON_RESUME_CONTRADICTION = (
    "The Servicer Memo Summary confirms this cycle is the deferral resume "
    "cycle (Regular Payment Resume Date falls in this reporting month), but "
    "the current servicer file still reports Loan Status = DEFERRED. The memo "
    "evidence is documented and internally consistent, so this is a Loan "
    "Status contradiction rather than a deferral data-quality problem. "
    "Requires explicit reviewer approval before anything is written."
)

# SUPERSEDED 2026-08-12 (the routing, NOT the rule). The reversal rule itself is
# unchanged and still CLIENT-CONFIRMED (2026-08-03): a Delinquency Status driven
# by a negative Payment Received is never automatically correctable. What changed
# is how that is enforced. It used to be an eligibility-layer `continue`, which
# did not merely block auto-correction -- it removed the loan from the
# recommendation list, the session candidate set and therefore Human Review
# entirely, so no reviewer could see the mismatch at all. Confirmed on Lendistry
# 2025, July 2026, loan 5951029110: STAT-01 flagged DELINQUENCY STATUS MISMATCH
# (reported 60 DAYS, calculated DELINQUENT off a consolidated net payment of
# -$10.92) and the recommendation was silently dropped.
#
# It is now routed the way every other must-be-judged case in this module is
# routed: review_required=True with the defect named, so the reviewer sees it and
# decides. Nothing is auto-applied -- _apply_one_correction's reviewer_approved
# gate is unchanged and still governs every write.
REVIEW_REASON_REVERSAL = (
    "This delinquency figure was driven by a REVERSAL — the consolidated net "
    "Payment Received for this cycle is negative. It is calculated by the "
    "client-confirmed deficiency formula (scheduled + |negative amount|, "
    "measured in 80%-of-scheduled units), which deliberately excludes the "
    "servicing fee, so the BR-002 payment-units ratio shown for ordinary "
    "recommendations does not apply here and is not displayed. A reversal is "
    "never automatically correctable: requires explicit reviewer approval "
    "before anything is written."
)

# CLIENT-CONFIRMED (2026-08-12). The offsetting-reversal half of the same
# 2026-08-03 rule. REVIEW_REASON_REVERSAL above is keyed on
# reversal_equivalents, which is only ever set when the CONSOLIDATED payment is
# negative. A reversal exactly cancelled by an offsetting positive posting nets
# to $0.00 and was therefore treated as an ordinary no-payment cycle: eligible
# for auto-correction, with no reversal reason shown to anyone. Measured on the
# real June 2026 fixtures: 33 recommendations, every one a clean +X/-X pair.
#
# Deliberately its OWN text rather than reusing REVIEW_REASON_REVERSAL, because
# the two are not the same fact and the reviewer must not be told something
# untrue: the net here is $0.00, not negative, and the deficiency formula did
# NOT run — the ordinary no-payment branch produced this figure. Keeping the
# texts separate also leaves every existing single-reason reversal string
# byte-identical.
REVIEW_REASON_REVERSAL_OFFSET = (
    "This cycle contains one or more REVERSED (negative) Payment Received "
    "postings that are offset by positive postings to a consolidated net of "
    "exactly $0.00. The delinquency figure was calculated by the ordinary "
    "no-payment branch — the reversal deficiency formula did not run — but a "
    "reversal cancelled out by an offsetting posting is still a reversal, and "
    "the client-confirmed rule is that reversal-driven delinquency figures are "
    "never automatically correctable: requires explicit reviewer approval "
    "before anything is written."
)

# CLIENT-CONFIRMED (2026-08-12). BANKRUPTCY is valid for a reporting period when
# the CURRENT period's Bankruptcy Date column contains a date. Before this,
# a servicer reporting ACTIVE against a prior SOUP BANKRUPTCY produced a
# LOAN STATUS MISMATCH that matched none of the three Loan Status branches below
# and whose expected value (BANKRUPTCY) is deliberately absent from
# _PROPOSABLE_LOAN_STATUS -- so no recommendation existed, the loan never
# reached Human Review for that field, and no reviewer could act on it.
# Confirmed on Lendistry 2025, July 2026, loans 7126119101, 5917099105,
# 6288639102, 6649779105 and 6666709108.
#
# The gate is the CURRENT period's Bankruptcy Date, never soup_status: the client
# was explicit that prior SOUP BANKRUPTCY alone must not support the proposal.
# That is also exactly why _PROPOSABLE_LOAN_STATUS stays unchanged -- it is
# shared with the terminal-contradiction branch, which proposes
# expected_loan_status_effective (a SOUP carry-forward when no override applies),
# so widening it would let BANKRUPTCY be proposed off prior SOUP alone. The
# marker on the candidate is what carries the evidence to the write gate instead.
REVIEW_REASON_BANKRUPTCY_CURRENT_PERIOD = (
    "The servicer reports a non-BANKRUPTCY Loan Status, but this reporting "
    "period's own Bankruptcy Date column is populated and the loan still carries "
    "a positive Ending Balance. Under the client-confirmed rule, a populated "
    "current-period Bankruptcy Date supports BANKRUPTCY for this period; prior "
    "SOUP status alone does not and played no part in this recommendation. "
    "Requires explicit reviewer approval before anything is written."
)

# CLIENT-CONFIRMED (2026-08-12): flag-and-propose, explicitly NOT suppress. A
# Liquidation Date alongside a POSITIVE Ending Balance does not trigger the
# client's zero-balance exception (that exception is keyed on Ending Balance =
# $0), so the Bankruptcy Date still stands and the recommendation is still made.
# But the row then carries three mutually inconsistent current-period signals --
# a non-BANKRUPTCY reported status, a bankruptcy date, and a liquidation date on
# a live balance -- and the client's instruction is that the reviewer sees that
# contradiction prominently rather than having it silently resolved either way.
# Suppressing the recommendation instead would put the loan straight back into
# the invisible state this whole change exists to end.
REVIEW_REASON_BANKRUPTCY_LIQUIDATION_CONTRADICTION = (
    "CONTRADICTORY CURRENT-PERIOD EVIDENCE: this period also reports a "
    "Liquidation Date, yet the Ending Balance is greater than $0 and the "
    "reported Loan Status is neither LIQUIDATION nor PAID-IN-FULL. The client's "
    "zero-balance exception does NOT apply here (it requires an Ending Balance "
    "of $0), so the Bankruptcy Date still supports BANKRUPTCY -- but the "
    "liquidation date and the live balance cannot both be right. Weigh both "
    "before approving."
)

# CLIENT-CONFIRMED (2026-08-12): the marker-based write gate. Set on a candidate
# ONLY by the current-period Bankruptcy Date branch, and the ONLY thing that lets
# validate_approved_value accept BANKRUPTCY as a value this application
# generated. Keyed on the marker rather than on the value so that BANKRUPTCY
# remains ungeneratable everywhere else in this module -- adding it to
# _PROPOSABLE_LOAN_STATUS would have widened all three other branches at once.
#
# Deliberately NOT trusted from client input: generate_reviewed_workbook reads it
# off the SERVER-SIDE candidate held in the session, never off the decision
# payload, so a caller cannot mint the marker to unlock a BANKRUPTCY write.
CURRENT_PERIOD_EVIDENCE_MARKER = "current_period_bankruptcy_date"

# SUPERSEDED 2026-08-08 (same day, later clarification). An earlier version of
# this module carried _NEVER_WRITABLE_LOAN_STATUS = (LIQUIDATION, FORECLOSURE,
# REO), refusing those values unconditionally at three layers. That conflated
# two different things, and the replacement rule separates them:
#
#   The APPLICATION must never independently GENERATE or automatically apply a
#   terminal collection status  ->  enforced by _PROPOSABLE_LOAN_STATUS below,
#   which is what the recommendation engine may emit.
#
#   A HUMAN REVIEWER may explicitly select and approve ANY legitimate Loan
#   Status  ->  governed by _PERMITTED_LOAN_STATUS, the client's authoritative
#   domain. A reviewer who examines the evidence and concludes the servicer's
#   LIQUIDATION report is correct must be able to say so.
#
# Nothing is written without explicit approval either way (see
# _apply_one_correction's reviewer_approved gate), so "never automatically"
# survives intact without a blanket value ban.


# Loan Status values the terminal-review branch may PROPOSE from independent
# evidence. Deliberately SEPARATE from _PERMITTED_LOAN_STATUS, which governs
# reviewer SUBSTITUTIONS and stays {ACTIVE, DEFERRED}: PAID-IN-FULL may be
# proposed when SOUP independently supports it, but may never be typed in by
# hand. An allow-list, not a deny-list, so blank, UNVERIFIED, malformed and
# cross-domain values ("CURRENT", "60 DAYS") are rejected by construction.
# NOT every valid SOUP Loan Status belongs here: BANKRUPTCY, SUBSTITUTION-IN
# and SUBSTITUTION-OUT are real SOUP values deliberately left out.
_PROPOSABLE_LOAN_STATUS = ("ACTIVE", "DEFERRED", "PAID-IN-FULL")


def _canonical_in(domain, value):
    """Resolves a value to its CANONICAL spelling within a domain, or None if
    it isn't in that domain. Validity check and canonicaliser in one.

    Canonicalising matters because the sources disagree on formatting: SOUP
    carries "Paid-in-Full" beside "PAID-IN-FULL" and "Active"/"ACTiVE", while
    the servicer workbooks this tool writes into are uniformly upper-hyphenated.
    Resolving through the domain guarantees the workbook's own representation
    however the value arrived."""
    return {_norm(v): v for v in domain}.get(_norm(value))


def _canonical_proposable_loan_status(value):
    """Validity check and canonicaliser in one. Returns the CANONICAL spelling
    of a proposable Loan Status, or None if the value may not be proposed.

    Canonicalising matters: SOUP is mixed-case and mixed-format ("Paid-in-Full"
    alongside "PAID-IN-FULL", "Active"/"ACTiVE"), while the servicer workbooks
    this tool writes into are uniformly "PAID-IN-FULL"/"ACTIVE". Uppercasing
    whatever SOUP happened to contain works only by luck -- a SOUP value of
    "Paid in Full" would write "PAID IN FULL" into a workbook whose canon is
    hyphenated. Resolving through this table guarantees the workbook's own
    representation regardless of how SOUP spelled it."""
    return _canonical_in(_PROPOSABLE_LOAN_STATUS, value)


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
#   6. Every other Loan Status mismatch is left for manual review.
#
# NARROWED 2026-08-08 — rules 2, 3 and 4 were written when this module
# AUTO-APPLIED corrections. Since the review-then-generate restructure there
# is no automatic-apply path: /api/audit only recommends, and
# generate_reviewed_workbook writes solely what a reviewer explicitly
# approved. Left unchanged, those rules vetoed a HUMAN's decision rather than
# an automatic write, which was never their purpose.
#
# They now separate the two things they had conflated:
#   GENERATING a recommendation is permitted when INDEPENDENT evidence (SOUP's
#     last recorded status) contradicts the servicer's current terminal claim.
#   AUTOMATICALLY APPLYING one remains impossible -- every such correction is
#     review_required and writes only on explicit approval.
#
# UNCHANGED: cross-domain protection; the deferral branch and its guards; and
# the rule that nothing is written without explicit reviewer approval.
#
# TWO DIFFERENT NOTIONS OF "TERMINAL" live in this module, and conflating them
# is a real bug rather than a nicety -- they overlap only on LIQUIDATION:
#
#   _COLLECTION_STATUS_KEYWORDS  FORECLOSURE / LIQUIDATION / REO. ACTIVE,
#     ongoing collection proceedings. Matches checks._TERMINAL_COLLECTION_
#     KEYWORDS exactly, which is what raises the TERMINAL STATUS REVIEW
#     reason, so the review flag and the reason can never disagree about the
#     same loan. This is the set the terminal-review logic uses.
#
#   _LOAN_STATUS_TERMINAL_KEYWORDS  LIQUIDATION / PAID-IN-FULL / PIF. Used
#     ONLY by the deferral safety guards, which is a different question: may
#     the application AUTO-DERIVE a "-> DEFERRED" Loan Status for this loan?
#     PAID-IN-FULL belongs there (never re-defer a paid-off loan) and does NOT
#     belong in the collection set (a paid-off loan is finished, not in
#     collection). Keep them separate.
_COLLECTION_STATUS_KEYWORDS = ("FORECLOSURE", "LIQUIDATION", "REO")


def _is_collection_status(status_value):
    """True when the servicer reports an ACTIVE collection proceeding. The
    trigger for terminal-status REVIEW routing."""
    s = str(status_value or "").strip().upper()
    return any(k in s for k in _COLLECTION_STATUS_KEYWORDS)


def _is_deferred_status(status_value):
    """True when the servicer reports the loan as still DEFERRED. The trigger
    for confirmed-resume contradiction routing."""
    return "DEFERRED" in str(status_value or "").strip().upper()


# The shared deferral-evidence contract value (checks.EVIDENCE_DOCUMENTED /
# DEFERRAL_EVIDENCE_STATUSES) meaning: deferral confirmed, dates consistent,
# type known. Spelled as a literal because this module deliberately does not
# import checks -- STAT-01 decides what the evidence says and this module only
# reads the verdict.
_EVIDENCE_DOCUMENTED = "documented"


_LOAN_STATUS_TERMINAL_KEYWORDS = ("LIQUIDATION", "PAID-IN-FULL", "PIF")


def _is_terminal_loan_status(status_value):
    """True for a loan that is liquidating or already paid off. Used ONLY by
    the deferral safety guards -- never for review routing, which uses
    _is_collection_status."""
    s = str(status_value or "").strip().upper()
    return any(k in s for k in _LOAN_STATUS_TERMINAL_KEYWORDS)


def _is_bankruptcy_status(status_value):
    """True when the servicer already reports the loan as BANKRUPTCY. Used only
    to stop the current-period Bankruptcy Date branch from proposing a value
    identical to the one already reported."""
    return "BANKRUPTCY" in str(status_value or "").strip().upper()


def _is_positive_balance(balance_value):
    """True when Ending Balance is greater than $0 at CENT precision -- the
    precision every figure in this application is reported and compared at, so a
    balance that is zero to the cent is zero here even with floating-point
    residue behind it.

    A missing or unparseable balance is NOT positive. That is deliberate: this
    guards a Loan Status proposal, and 'we could not read the balance' must never
    be treated as 'the loan has money outstanding'."""
    try:
        return round(float(balance_value), 2) > 0.00
    except (TypeError, ValueError):
        return False


# Loan Status values a reviewer may SELECT as an approved value. A value that
# is actually a Delinquency Status is still rejected as cross-domain.
# WIDENED 2026-08-08. Sourced from the SRF template's own "Reference Tables"
# sheet, which carries the client's Loan Status validation list -- identical
# across every servicer workbook inspected (CBT, Lendistry, Peoples). Using
# the client's own domain rather than an invented subset is what lets Loan
# Status review work like Delinquency Status review: the reviewer picks from
# the real field domain instead of a two-value stub.
#
# CHARGE-OFF is deliberately absent: it is not in the template's Loan Status
# list. It lives in the Servicer Memo's Securitization Action field, which is
# a different field entirely.
#
# This is the REVIEWER's domain. What the APPLICATION may generate is the
# narrower _PROPOSABLE_LOAN_STATUS -- the two must not be merged.
_PERMITTED_LOAN_STATUS = {
    "ACTIVE", "DEFERRED", "PAID-IN-FULL", "REO", "FORECLOSURE",
    "BANKRUPTCY", "SUBSTITUTION-IN", "SUBSTITUTION-OUT", "LIQUIDATION",
}


def validate_approved_value(field, value, substituted=True, evidence_marker=None):
    """The approved-value domain restriction. A reviewer's substituted value
    must come only from the target field's own domain, so a substitution can
    never reintroduce a cross-domain value (the exact defect that flags a
    recommendation for review in the first place) or turn a loan terminal.

    evidence_marker is the CANDIDATE's server-side marker (see
    CURRENT_PERIOD_EVIDENCE_MARKER), never a client-supplied field. It widens the
    GENERATED Loan Status domain by exactly one value, for exactly the
    recommendation path that established the evidence for it.

    Returns (True, normalized_value) if the value is writable for this field,
    else (False, reason)."""
    v = str(value if value is not None else "").strip().upper()
    if not v:
        return False, "Approved value is blank."
    if field == "Delinquency Status":
        if v not in _PERMITTED_DQ_STATUS:
            return False, (
                f"'{value}' is not a valid Delinquency Status — an approved value must come "
                f"from the delinquency ladder, not another field."
            )
        return True, v
    if field == "Loan Status":
        if not substituted:
            # The reviewer approved the value the APPLICATION derived. It must
            # come from the narrow generated domain -- this is the layer that
            # keeps the application from ever proposing a terminal collection
            # status off its own bat.
            # CLIENT-CONFIRMED (2026-08-12): the marker-based exception, checked
            # BEFORE the general domain so BANKRUPTCY can be written through the
            # current-period Bankruptcy Date path and nowhere else. Keyed on the
            # marker AND the value together: the marker does not unlock the whole
            # Loan Status domain, and BANKRUPTCY stays ungeneratable without it.
            #
            # This is why _PROPOSABLE_LOAN_STATUS was not widened. That set is
            # shared with the terminal-contradiction branch, whose proposal is a
            # SOUP carry-forward -- adding BANKRUPTCY there would have let it be
            # generated from prior SOUP alone, which the client ruled out.
            if evidence_marker == CURRENT_PERIOD_EVIDENCE_MARKER and v == "BANKRUPTCY":
                return True, "BANKRUPTCY"
            canonical = _canonical_proposable_loan_status(v)
            if canonical is None:
                return False, (
                    f"'{value}' is not a Loan Status this application may generate."
                )
            return True, canonical
        # The reviewer explicitly SELECTED this value, having looked at the
        # evidence. Any status in the client's authoritative domain is allowed,
        # including LIQUIDATION -- a reviewer who concludes the servicer's
        # terminal report is correct must be able to say so. Still gated: the
        # write itself requires explicit approval (_apply_one_correction).
        if v in _PERMITTED_DQ_STATUS:
            return False, (
                f"'{value}' is a Delinquency Status — an approved Loan Status must come "
                f"from the Loan Status domain, not another field."
            )
        canonical = _canonical_in(_PERMITTED_LOAN_STATUS, v)
        if canonical is None:
            return False, (
                f"'{value}' is not a permitted Loan Status substitution."
            )
        return True, canonical
        return True, v
    return False, f"Field '{field}' is not a correctable field."


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
# Existing columns are frozen — DO NOT reorder or rename them (external
# consumers read the log by position/name). New review-workflow columns are
# APPENDED after them:
#   - Calculated Value      : what STAT-01 computed for this loan/field.
#   - Reviewer Approved Value: what the reviewer authorised — equals
#                             Calculated Value unless they substituted.
#   - Decision              : Approved and applied / Rejected / Returned to servicer.
#   - Reviewer Note         : the reviewer's free-text note.
#   - Reviewed By (self-declared): the reviewer's name. Labelled "self-declared"
#                             because the app has no authentication — a typed
#                             name must never be mistaken for an authenticated
#                             one. Replace this column's source (not just its
#                             label) if/when real auth is added upstream.
#   - Reviewed On           : when the review was submitted.
#   - Source File SHA-256   : hash of the ORIGINAL uploaded workbook these
#                             corrections were reviewed against, so it is
#                             provable which source artifact was reviewed.
#   - Notify Servicer       : "Yes" when the reviewer flagged that the
#                             servicer should be told about this (a source-
#                             data defect that will recur unless raised),
#                             blank otherwise. This flag is the reason the
#                             review exists for cross-domain defects, so it
#                             must be recorded, not dropped.
# `Original Value` stays the reported value read from the cells; `Corrected
# Value` stays the value actually written (blank for a rejected/returned row).
CORRECTION_LOG_HEADERS = [
    "Institution", "Source Worksheet", "Excel Row Numbers", "Rows Updated", "SBA GP Number",
    "Field Updated", "Original Value", "Corrected Value", "Audit Check ID",
    "Eligibility Rule", "Timestamp", "Skipped Row References",
    "Review Status", "Review Reason",
    "Calculated Value", "Reviewer Approved Value", "Decision",
    "Reviewer Note", "Reviewed By (self-declared)", "Reviewed On",
    "Source File SHA-256", "Notify Servicer",
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
        #
        # ROUTING CORRECTED 2026-08-12 (see REVIEW_REASON_REVERSAL): "requires
        # manual review" is now implemented as review_required on the
        # Delinquency Status recommendation, not as a `continue` that deleted
        # the loan from Human Review altogether. The scope of the block is
        # unchanged for LOAN STATUS — see the guard below the DQ branch.
        # WIDENED 2026-08-12 (see REVIEW_REASON_REVERSAL_OFFSET): the rule's
        # SCOPE is unchanged — reversal-driven delinquency figures require
        # human review — but the test for "is this a reversal" no longer relies
        # on the consolidated net being negative, which a +X/-X pair defeats.
        # The two facts stay separate so the right reason can be stated and so
        # only a genuine net-negative reversal suppresses the BR-002 ratio.
        _reversal_net_negative = c.get("reversal_equivalents") is not None
        _reversal_offset = (
            bool(c.get("offsetting_zero_net_reversal")) and not _reversal_net_negative
        )
        _reversal = _reversal_net_negative or _reversal_offset

        gp = c.get("gp")
        excel_rows, skipped_row_refs = _clean_excel_rows(c.get("excel_rows"))

        # Which fields this loan has already had a recommendation emitted for.
        # _recommendation_id is (servicer, gp, field), so two recommendations for
        # the same loan/field would share one id and generate_reviewed_workbook's
        # `by_id` dict would silently keep only the last -- a reviewer's decision
        # applying to a recommendation they were not shown. The Loan Status
        # branches below are ordered by precedence, and this is what makes the
        # first match win instead of the last.
        _emitted_fields = set()

        def _emit(field, corrected_value, eligibility_rule,
                  review_required=False, review_reason=None,
                  evidence_marker=None):
            _emitted_fields.add(field)
            eligible.append({
                "gp": gp,
                "excel_rows": excel_rows,
                "skipped_row_refs": skipped_row_refs,
                "sheet_name": c.get("sheet_name"),
                "header_excel_row": c.get("header_excel_row"),
                "field": field,
                # corrected_value IS STAT-01's calculated value — the default
                # the reviewer sees, and what "Calculated Value" logs.
                "corrected_value": corrected_value,
                "eligibility_rule": eligibility_rule,
                # A recommendation flagged for review is never AUTO-applied.
                # It still reaches the reviewer (with the defect named) so it
                # can be judged, rather than disappearing the way an
                # eligibility-layer `continue` would make it disappear.
                "review_required": bool(review_required),
                "review_reason": review_reason,
                # The value actually reported by the servicer, carried so the
                # review page can show what needs reviewing without having
                # opened the workbook to read the cells.
                "reported_value": c.get("srf_dq") if field == "Delinquency Status"
                                  else c.get("srf_status"),
                # Read-only review evidence, passed straight through from
                # STAT-01's correction_data (see checks.py). Never used in any
                # decision here — purely for display on the review page.
                "prior_value": c.get("prior_dq") if field == "Delinquency Status" else c.get("soup_status"),
                "payment_received": c.get("payment_received"),
                "scheduled_payment": c.get("scheduled_payment"),
                "servicing_fee": c.get("servicing_fee"),
                # Passed through so the review page can tell a reversal row from
                # an ordinary one. The review page's BR-002 payment-units ratio
                # ((received + fee) / scheduled) is NOT the arithmetic that
                # produced calculated_dq for these rows -- the reversal branch of
                # checks._apply_payment_transition excludes the fee by design --
                # so this is the flag that suppresses that figure rather than
                # showing a reviewer a number that did not drive the result.
                # Display/branching only; never an input to eligibility.
                "reversal_equivalents": c.get("reversal_equivalents"),
                # The offsetting-reversal companion. Deliberately the COMPUTED
                # _reversal_offset, not the raw fact: the two reversal kinds are
                # mutually exclusive, and this must agree with which review
                # reason was actually stated. Display/branching only, never an
                # input to eligibility.
                #
                # Note what this does NOT do: it does not suppress the review
                # page's BR-002 ratio the way reversal_equivalents does. For an
                # offsetting reversal the net IS $0.00 and the ordinary
                # no-payment branch DID produce the figure, so the ratio is the
                # arithmetic that ran and stays on screen. This flag only adds
                # the fact that reversed postings are hidden inside that $0.00.
                "offsetting_zero_net_reversal": _reversal_offset,
                # The shared deferral-evidence contract (checks.
                # DEFERRAL_EVIDENCE_STATUSES). Passed through verbatim and
                # NEVER recomputed here: STAT-01 is the only place that
                # decides what the evidence says, so the review page, the
                # workbook and this module cannot drift into three readings
                # of the same facts. Together with prior_value and
                # reported_value above, these are the six columns that let a
                # reviewer see WHY a deferral row was flagged — most of all a
                # label conflict, where a "Modification" Securitization
                # Action sits against a full-P&I window and is unreadable
                # without the prior and current statuses beside it.
                "evidence_status": c.get("evidence_status"),
                "deferral_start": c.get("deferral_start"),
                "deferral_end": c.get("deferral_end"),
                "deferral_resume": c.get("deferral_resume"),
                "securitization_action": c.get("securitization_action"),
                "servicer_loan_status": c.get("srf_status"),
                "prior_loan_status": c.get("soup_status"),
                "balance": c.get("ending_balance"),
                "posting_count": c.get("posting_count"),
                # CLIENT-CONFIRMED (2026-08-12): the current-period evidence
                # marker. None on every recommendation except one produced by the
                # Bankruptcy Date branch below. This is what validate_approved_
                # value checks to allow BANKRUPTCY as a generated value, so it is
                # a WRITE GATE, not display -- see CURRENT_PERIOD_EVIDENCE_MARKER.
                "evidence_marker": evidence_marker,
                # Read-only current-period evidence, passed straight through from
                # STAT-01. Display only on the review page; the eligibility
                # decision was already made below and is never re-derived from
                # these on the client side.
                "bankruptcy_date": c.get("bankruptcy_date"),
                "paid_in_full_date": c.get("paid_in_full_date"),
                "liquidation_date": c.get("liquidation_date"),
            })

        # --- Delinquency Status -------------------------------------------------
        if (
            "DELINQUENCY STATUS MISMATCH" in reasons
            and c.get("dq_match") is False  # never True/"REVIEW" — REVIEW means UNVERIFIED, already excluded upstream
            and _is_real_dq_value(c.get("calculated_dq"))
        ):
            _cross_domain = _is_cross_domain_dq(c.get("srf_dq"))
            # A servicer-reported terminal status forces HUMAN REVIEW of this
            # loan's corrections; it never suppresses them. The delinquency
            # figure was calculated by the same normal ladder as every other
            # loan, and the servicer's reported DQ played no part in it.
            _terminal = _is_collection_status(c.get("srf_status"))
            # Accumulated rather than an if/elif chain: a loan can carry more
            # than one of these at once (a reversal on a loan the servicer also
            # reports as FORECLOSURE), and a reviewer must see every reason it
            # needs judging, not just the first. With exactly one reason present
            # the joined string is that reason verbatim, so the single-reason
            # cases keep the exact text they have always had.
            _review_reasons = []
            if _cross_domain:
                _review_reasons.append(REVIEW_REASON_CROSS_DOMAIN)
            if _terminal:
                _review_reasons.append(REVIEW_REASON_TERMINAL_STATUS)
            # Mutually exclusive by construction: a loan is either a
            # net-negative reversal (the deficiency formula ran) or an
            # offsetting zero-net one (it did not). Never both, so an existing
            # net-negative reversal keeps its exact original wording.
            if _reversal_net_negative:
                _review_reasons.append(REVIEW_REASON_REVERSAL)
            elif _reversal_offset:
                _review_reasons.append(REVIEW_REASON_REVERSAL_OFFSET)
            _emit(
                "Delinquency Status",
                str(c["calculated_dq"]).strip().upper(),
                "STAT-01 deterministic calculation",
                review_required=bool(_review_reasons),
                review_reason=" ".join(_review_reasons) or None,
            )

        # A reversal blocks the LOAN STATUS branches below exactly as the old
        # eligibility-layer `continue` did — that half of the 2026-08-03 rule is
        # deliberately unchanged. Only the Delinquency Status recommendation
        # above is newly routed to review; nothing here is widened. An
        # offsetting zero-net reversal (2026-08-12) reaches this block on the
        # same footing as a net-negative one: the whole point of that change is
        # that the two are the same kind of case, so treating them differently
        # HERE would reintroduce the hole one layer down. Placed after
        # the DQ branch and before the three Loan Status branches so the scope of
        # the block is visible in one line rather than repeated three times.
        if _reversal:
            continue

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
            # Strict equality, deliberately NOT checks._FULL_PNI_CALC_TYPES.
            # An "unclassified" deferral rides the full P&I path for the
            # delinquency ARITHMETIC (so no number moves on unknown evidence),
            # but it must never justify writing a Loan Status. Its DEFERRAL
            # DATA QUALITY WARNING already excludes the loan above; this keeps
            # it excluded even if that warning were ever dropped.
            and c.get("deferral_type") == "full"
            and c.get("deferral_phase") in ("start", "mid")
            and not _is_terminal_loan_status(c.get("srf_status"))
            and not _is_terminal_loan_status(c.get("soup_status"))
        ):
            _emit("Loan Status", "DEFERRED", f"STAT-01 confirmed Full P&I deferral ({c.get('deferral_phase')})")

        # --- Loan Status: servicer-asserted terminal status contradicted by
        # --- independent evidence (2026-08-08) --------------------------------
        # SEPARATE from the deferral branch above, which is untouched. Scoped
        # to one situation: the servicer reports a terminal collection status
        # this cycle and the audit's own independently derived Expected Loan
        # Status says something else.
        #
        # The proposed value is expected_loan_status_effective -- STAT-01's
        # single source of truth, computed once in checks.py. Deliberately NOT
        # soup_status: prior SOUP is independent evidence but not always the
        # answer, and copying it would ignore every deferral / resume /
        # bankruptcy branch that legitimately changes it (e.g. SOUP still says
        # DEFERRED while the memo confirms payments resumed -- expected is the
        # resumed status, and proposing DEFERRED would be stale). This module
        # never re-derives that value.
        #
        # _canonical_proposable_loan_status is both the validity gate and the
        # canonicaliser: it rejects blank / UNVERIFIED / malformed /
        # cross-domain values by construction, restricts what the APPLICATION
        # may generate to ACTIVE/DEFERRED/PAID-IN-FULL (so a terminal
        # collection status can never be machine-generated), and returns the
        # workbook's own spelling rather than whatever case SOUP used.
        #
        # ALWAYS review_required -- recommendation only, never written without
        # explicit approval. A collection status can never be the PROPOSED
        # value (allow-list, plus the guard here and again at write time), so
        # this can only move a loan OUT of a servicer-asserted terminal
        # status, never into one.
        _proposed_status = _canonical_proposable_loan_status(c.get("expected_loan_status_effective"))
        if (
            "LOAN STATUS MISMATCH" in reasons
            and _is_collection_status(c.get("srf_status"))
            and _proposed_status is not None
            and _norm(_proposed_status) != _norm(c.get("srf_status"))
        ):
            _emit(
                "Loan Status",
                _proposed_status,
                "Servicer-reported terminal status contradicted by the audit's independently "
                "derived Expected Loan Status",
                review_required=True,
                review_reason=REVIEW_REASON_TERMINAL_STATUS,
            )

        # --- Loan Status: confirmed deferral resume contradicted by the
        # --- servicer still reporting DEFERRED (2026-08-10) -------------------
        # SEPARATE from both branches above, which are untouched. Scoped to one
        # situation: the Servicer Memo confirms this cycle is the deferral
        # resume cycle, its evidence is documented, and the servicer still
        # reports DEFERRED.
        #
        # SUPERSEDES the v1 scope decision recorded on the "-> DEFERRED" branch
        # above: "a correction is only ever generated INTO Deferred, never out
        # of it -- v1 never auto-exits a loan from Deferred during a resume
        # cycle, by construction." That was a deliberate v1 narrowing, not a
        # safety property, and the client superseded it on 2026-08-10. Nothing
        # is auto-applied: review_required is always True here and the
        # reviewer_approved gate in _apply_one_correction still governs every
        # write, so "never automatically" survives intact.
        #
        # The proposed value is expected_loan_status_effective, STAT-01's single
        # source of truth, run through _canonical_proposable_loan_status --
        # which restricts what the APPLICATION may generate to
        # ACTIVE/DEFERRED/PAID-IN-FULL, so a terminal collection status can
        # never be machine-generated from this path either.
        #
        # evidence_status must be "documented". A resume phase can also be
        # reached on a date-anomaly record, and this branch proposes EXITING a
        # deferral -- the one direction v1 refused -- so it is deliberately
        # gated on the strongest evidence the contract defines rather than
        # merely on the phase. (The unclassified-type case is already excluded
        # upstream by WARNING_REASONS; this gate is what covers the rest.)
        #
        # Cannot collide with the two branches above: the "-> DEFERRED" branch
        # requires expected_loan_status_raw == "DEFERRED" (here it is ACTIVE)
        # and phase in (start, mid) (here it is resume); the terminal branch
        # requires _is_collection_status(srf_status), and DEFERRED is not a
        # collection status.
        _proposed_resume = _canonical_proposable_loan_status(c.get("expected_loan_status_effective"))
        if (
            "LOAN STATUS MISMATCH" in reasons
            and c.get("deferral_phase") == "resume"
            and c.get("evidence_status") == _EVIDENCE_DOCUMENTED
            and _is_deferred_status(c.get("srf_status"))
            and _proposed_resume is not None
            and _norm(_proposed_resume) != _norm(c.get("srf_status"))
        ):
            _emit(
                "Loan Status",
                _proposed_resume,
                "STAT-01 confirmed deferral resume contradicted by the servicer still "
                "reporting DEFERRED",
                review_required=True,
                review_reason=REVIEW_REASON_RESUME_CONTRADICTION,
            )

        # --- Loan Status: BANKRUPTCY supported by a current-period Bankruptcy
        # --- Date (2026-08-12) ------------------------------------------------
        # SEPARATE from the three branches above, all of which are untouched.
        # Scoped to one situation: the servicer reports something other than
        # BANKRUPTCY, and THIS reporting period's own Bankruptcy Date column is
        # populated on a loan that still has money outstanding.
        #
        # Reached only after the `if _reversal: continue` above, so reversal
        # precedence is preserved for free rather than restated here -- a
        # reversal-driven loan gets its Delinquency Status routed to review and
        # no Loan Status recommendation at all, exactly as before.
        #
        # Cannot collide with the three branches above. The "-> DEFERRED" branch
        # requires expected_loan_status_raw == "DEFERRED" (null on these loans);
        # the terminal branch requires _is_collection_status(srf_status), which is
        # excluded here by construction (rule 3 below rejects LIQUIDATION, and a
        # FORECLOSURE/REO loan is left to that branch and its own SOUP-derived
        # proposal); the resume branch requires deferral_phase == "resume" and a
        # DEFERRED reported status.
        #
        # The four gates, in the client's own terms:
        #   1. This period's Bankruptcy Date CONTAINS a date. A presence test on
        #      the raw cell, decided once in checks.py -- never a date parse, since
        #      the column is genuinely mixed-typed in real files.
        #   2. Ending Balance > $0, at cent precision (the precision every figure
        #      in this application is compared at).
        #   3. THE ZERO-BALANCE EXCEPTION. Preserved as the client stated it: when
        #      the balance is $0 because the loan is PAID IN FULL or LIQUIDATION,
        #      the Bankruptcy Date is ignored for this rule. Gate 2 already
        #      excludes every $0 balance, so this is the belt-and-braces half --
        #      it also refuses a PAID-IN-FULL / LIQUIDATION loan that reports a
        #      POSITIVE balance, which gate 2 would otherwise let through and
        #      which must never be re-proposed as BANKRUPTCY. Checked against the
        #      reported status AND the Paid-In-Full Date column, since either one
        #      alone is enough to say the loan is finished.
        #   4. The servicer does not already say BANKRUPTCY. Without this the
        #      module would propose a value identical to the reported one for
        #      every consistent bankruptcy loan in the file.
        #
        # soup_status is deliberately absent from every gate. The client was
        # explicit that prior SOUP BANKRUPTCY must not on its own support this
        # proposal, so it is not merely insufficient here -- it is not consulted.
        # A loan whose SOUP history says BANKRUPTCY but whose current Bankruptcy
        # Date is blank produces no recommendation from this branch.
        #
        # ALWAYS review_required -- a recommendation, never an automatic write.
        _bankruptcy_dated = bool(c.get("bankruptcy_date_populated"))
        _balance_outstanding = _is_positive_balance(c.get("ending_balance"))
        _finished_loan = (
            bool(c.get("paid_in_full_date_populated"))
            or _is_terminal_loan_status(c.get("srf_status"))
        )
        if (
            "LOAN STATUS MISMATCH" in reasons
            # LAST of the four Loan Status branches, so any of the three above
            # takes precedence. A confirmed full-P&I deferral with a stray
            # Bankruptcy Date is the real overlap: both branches' gates can hold
            # at once, and without this the loan would carry two contradictory
            # Loan Status recommendations sharing one id.
            and "Loan Status" not in _emitted_fields
            and _bankruptcy_dated
            and _balance_outstanding
            and not _finished_loan
            and not _is_bankruptcy_status(c.get("srf_status"))
        ):
            # Flag-and-propose, never suppress (see
            # REVIEW_REASON_BANKRUPTCY_LIQUIDATION_CONTRADICTION). Accumulated the
            # same way the Delinquency Status reasons are, so a loan carrying both
            # facts shows both; with only the first present the joined string is
            # that reason verbatim.
            _bk_reasons = [REVIEW_REASON_BANKRUPTCY_CURRENT_PERIOD]
            if bool(c.get("liquidation_date_populated")):
                _bk_reasons.append(REVIEW_REASON_BANKRUPTCY_LIQUIDATION_CONTRADICTION)
            _emit(
                "Loan Status",
                "BANKRUPTCY",
                "Current-period Bankruptcy Date populated with a positive Ending Balance",
                review_required=True,
                review_reason=" ".join(_bk_reasons),
                evidence_marker=CURRENT_PERIOD_EVIDENCE_MARKER,
            )

    return eligible


def _recommendation_id(servicer, gp, field):
    """Stable, deterministic id for one recommendation within a session.
    Deterministic (not a uuid) so a resubmitted review maps to the same
    candidate. servicer/gp/field never contain '||'."""
    return f"{servicer}||{gp}||{field}"


def build_recommendations(stat01_result, servicer):
    """Turns one servicer's STAT-01 result into the review-page contract,
    WITHOUT writing anything. This is what /api/audit returns now instead of
    a generated workbook.

    Returns (recommendations, candidates, skipped):
      - recommendations: client-facing dicts (one per eligible loan/field
        that has at least one valid posting row to write to), carrying the
        display evidence the reviewer needs. A recommendation is a PROPOSAL,
        never an applied change — a `reviewFlag` names any defect (e.g. a
        cross-domain reported value) the reviewer must weigh.
      - candidates: the matching full server-side records (same order/ids),
        including the sheet/row/header write metadata. main.py stashes these
        in the session so generate_reviewed_workbook can act on a reviewer's
        decision without trusting client-supplied write targets.
      - skipped: eligible loans/fields with NO valid posting row to write to
        — not reviewable (there is nowhere to apply a decision), surfaced via
        corrections_skipped exactly as before.

    Eligibility itself is unchanged — build_eligible_corrections is the sole
    allow-list; this function only shapes and routes its output."""
    correction_data = (stat01_result or {}).get("correction_data", [])
    eligible = build_eligible_corrections(correction_data)

    recommendations = []
    candidates = []
    skipped = []
    for e in eligible:
        rec_id = _recommendation_id(servicer, e["gp"], e["field"])

        if not e["excel_rows"]:
            # Eligible by every business rule, but no valid row to write to —
            # a decision could never be applied, so it is not reviewable.
            detail = (
                "; ".join(f"{s['row_ref']} ({s['reason']})" for s in e.get("skipped_row_refs") or [])
                if e.get("skipped_row_refs") else "no row location metadata available"
            )
            skipped.append({
                "SBA GP Number": e["gp"],
                "field": e["field"],
                "reason": (
                    f"Eligible for correction, but no valid posting row could be located "
                    f"({detail}) — cannot be reviewed."
                ),
            })
            continue

        candidate = dict(e)
        candidate["id"] = rec_id
        candidate["servicer"] = servicer
        candidates.append(candidate)

        recommendations.append({
            "id": rec_id,
            "servicer": servicer,
            "gp": e["gp"],
            "field": e["field"],
            "rows": list(e["excel_rows"]),
            "postings": e.get("posting_count"),
            "sheet": e.get("sheet_name"),
            "reported": e.get("reported_value"),
            "prior": e.get("prior_value"),
            "calculated": e["corrected_value"],
            "paymentReceived": e.get("payment_received"),
            "scheduledPayment": e.get("scheduled_payment"),
            "servicingFee": e.get("servicing_fee"),
            # Non-null ONLY on a reversal row (the count of scheduled-payment
            # equivalents the negative amount represents). The review page reads
            # this to suppress its BR-002 payment-units ratio, which is not the
            # arithmetic that produced `calculated` for these rows.
            "reversalEquivalents": e.get("reversal_equivalents"),
            # True ONLY on an offsetting zero-net reversal. The review page reads
            # this to say so beside the $0.00, which would otherwise be
            # indistinguishable from a genuine no-payment cycle.
            "offsettingZeroNetReversal": bool(e.get("offsetting_zero_net_reversal")),
            "balance": e.get("balance"),
            "reviewFlag": e.get("review_reason") if e.get("review_required") else None,
            "rule": e.get("eligibility_rule"),
            # Deferral evidence, carried to the review page so a reviewer can
            # judge a deferral correction without opening the memo workbook
            # alongside the audit. Display only, computed once in STAT-01.
            "evidenceStatus": e.get("evidence_status"),
            "deferralStart": e.get("deferral_start"),
            "deferralEnd": e.get("deferral_end"),
            "deferralResume": e.get("deferral_resume"),
            "memoSecuritizationAction": e.get("securitization_action"),
            "servicerLoanStatus": e.get("servicer_loan_status"),
            "priorLoanStatus": e.get("prior_loan_status"),
            # CLIENT-CONFIRMED (2026-08-12): current-period terminal-event
            # evidence, carried so the review page can show the reviewer the
            # Bankruptcy Date this recommendation rests on -- and, for a
            # contradicted row, the Liquidation Date and positive balance beside
            # it. Display only; the eligibility decision was made server-side and
            # the review page never re-derives it. The evidence_marker itself is
            # deliberately NOT exposed: it is a write gate read off the session
            # candidate, and echoing it to the client would invite it back as
            # client input.
            "bankruptcyDate": e.get("bankruptcy_date"),
            "paidInFullDate": e.get("paid_in_full_date"),
            "liquidationDate": e.get("liquidation_date"),
        })

    return recommendations, candidates, skipped


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
    # 2026-08-08: EVERY Loan Status write requires explicit reviewer approval,
    # not merely the terminal ones. This is now the whole of the "never
    # automatically applied" rule -- there is no blanket ban on any particular
    # value, so this gate carries the guarantee alone and is deliberately
    # broader than the terminal case it replaced.
    #
    # Fails closed: reviewer_approved is set in exactly one place, after the
    # decision is confirmed an approval AND the value has passed validation. A
    # caller that omits it is refused.
    if corr["field"] == "Loan Status" and not corr.get("reviewer_approved"):
        return False, (
            f"Refused to write Loan Status '{corr['corrected_value']}' for loan {corr['gp']} — "
            f"a Loan Status may only be written on an explicit reviewer approval."
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


def _add_correction_log_sheet(wb, log_rows, institution, common):
    """Writes the Correction Log. `common` carries the values shared by
    every row of one review submission: {timestamp, reviewed_by,
    reviewed_on, source_sha256}. Every decided correction — applied,
    rejected, and returned — gets exactly one row (rejected/returned rows
    have Rows Updated = 0 and a blank Corrected Value, since no cell
    changed)."""
    ws = wb.create_sheet(_unique_log_sheet_name(wb))

    rows = [
        [
            institution,
            r.get("sheet_name") or "",
            ", ".join(str(x) for x in (r.get("excel_rows") or [])),
            r.get("rows_updated", 0),
            r["gp"],
            r["field"],
            "; ".join("" if v is None else str(v) for v in (r.get("original_values") or [])),
            "" if r.get("corrected_value") in (None, "") else r["corrected_value"],
            AUDIT_CHECK_ID,
            r.get("eligibility_rule") or "",
            common["timestamp"],
            "; ".join(f"{s['row_ref']} ({s['reason']})" for s in r.get("skipped_rows") or []),
            r.get("review_status") or "",
            r.get("review_reason") or "",
            r.get("calculated_value") or "",
            r.get("approved_value") or "",
            r.get("decision") or "",
            r.get("reviewer_note") or "",
            common["reviewed_by"],
            common["reviewed_on"],
            common["source_sha256"],
            r.get("notify_servicer") or "",
        ]
        for r in log_rows
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


def format_self_declared_reviewer(name):
    """The Correction Log's reviewer identity. The app has no authentication,
    so a typed name is UNVERIFIED and must be recorded as such — a reader of
    the log must never mistake it for an authenticated identity. If real auth
    is added upstream later, replace this with the authenticated source
    rather than relabelling a free-text field."""
    name = (name or "").strip()
    if not name:
        return ""
    return f"{name} (unverified)"


# Human-readable Decision / Review Status text, kept in one place so the log
# and the API summary can't drift apart.
_DECISION_APPROVE = "approve"
_DECISION_REJECT = "reject"
_DECISION_RETURN = "return"
_DECISION_LABELS = {
    _DECISION_APPROVE: "Approved and applied",
    _DECISION_REJECT: "Rejected",
    _DECISION_RETURN: "Returned to servicer",
}


def generate_reviewed_workbook(
    file_bytes, original_filename, institution, candidates, decisions,
    reviewer, source_sha256, cycle_month, reviewed_on=None,
):
    """Generates ONE servicer's corrected workbook from a FRESH copy of its
    ORIGINAL uploaded bytes (never mutated), writing ONLY the reviewer-
    approved values and recording every decision in the Correction Log.

    Inputs:
      candidates  — that servicer's candidate list from build_recommendations
                    (each carries a stable `id` and its write metadata).
      decisions   — a list of {id, decision, approved_value?, reviewer_note?,
                    notify_servicer?}. Only ids present here are acted on and
                    logged; an undecided candidate is neither written nor
                    logged. `decision` is approve / reject / return.
      reviewer    — the raw reviewer name (formatted self-declared here).
      source_sha256 — SHA-256 of `file_bytes`, written to the log so it is
                    provable which source artifact was reviewed.

    Returns a dict:
      {file_bytes, filename, applied_count, rejected_count, returned_count,
       corrections_not_applied: [{gp, field, reason}],
       unknown_ids: [...], skip_reason,
       verified_overrides: {(gp, field): approved_value}}

    verified_overrides is the VERIFIED layer -- the third of the three the
    client confirmed on 2026-08-10:

        Original   what the lender submitted
        Audited    what the audit calculated / recommended
        Verified   the final value after Manual Review

    It carries an entry ONLY where a value was approved, validated and
    successfully written, so it describes the same cells the corrected
    workbook contains and cannot drift from it. Absence means the lender's
    original value stands -- which is equally true of the workbook, since
    nothing was written to those cells.

    Never raises for expected conditions (unknown id, unsupported format,
    template drift, an invalid substituted value) — each is reported so one
    servicer's problem never breaks a multi-servicer review. Applying only
    approved values, the domain restriction on substitutions, and the
    terminal-Loan-Status write block are all enforced here, never bypassable
    by client input."""
    empty = {
        "file_bytes": None, "filename": None,
        "applied_count": 0, "rejected_count": 0, "returned_count": 0,
        "corrections_not_applied": [], "unknown_ids": [], "skip_reason": None,
        # Always present, even on an early return, so every caller sees one
        # contract and never has to guess whether the key exists.
        "verified_overrides": {},
    }

    by_id = {c["id"]: c for c in (candidates or [])}
    decisions = decisions or []
    unknown_ids = [d.get("id") for d in decisions if d.get("id") not in by_id]
    actionable = [d for d in decisions if d.get("id") in by_id]
    if not actionable:
        return {**empty, "unknown_ids": unknown_ids,
                "skip_reason": "No reviewable decisions were submitted for this servicer."}

    if original_filename and original_filename.lower().strip().endswith(".xls"):
        return {**empty, "unknown_ids": unknown_ids,
                "skip_reason": ("Decisions were submitted, but a corrected copy cannot be "
                                "generated safely from the legacy .xls format.")}

    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
    except Exception as exc:
        return {**empty, "unknown_ids": unknown_ids,
                "skip_reason": f"Could not reopen the original workbook: {exc}"}

    log_rows = []
    corrections_not_applied = []
    applied_count = rejected_count = returned_count = 0
    # {(gp, field): approved_value} -- the VERIFIED layer.
    #
    # Populated in exactly one place: the `written_ok` branch below, the same
    # condition that increments applied_count. So an entry exists if and only
    # if a reviewer approved the value, it passed validate_approved_value, AND
    # it was actually written to the workbook. Reject, return, undecided, a
    # domain-refused substitution and a failed write all correctly produce no
    # entry -- and "no entry" means "the lender's original value stands",
    # which is exactly what the corrected workbook contains for those cells.
    #
    # This is what lets the corrected workbook and the Final Audited Portfolio
    # Summary be generated from ONE source: there is no second decision point
    # where they could disagree.
    verified_overrides = {}

    for d in actionable:
        cand = by_id[d["id"]]
        gp, field = cand["gp"], cand["field"]
        calculated = cand.get("corrected_value")
        reported = cand.get("reported_value")
        note = d.get("reviewer_note")
        notify = "Yes" if d.get("notify_servicer") else ""
        review_reason = cand.get("review_reason") or ""
        decision_kind = (d.get("decision") or "").strip().lower()

        base_row = {
            "sheet_name": cand.get("sheet_name"),
            "excel_rows": [],
            "rows_updated": 0,
            "gp": gp,
            "field": field,
            "original_values": [reported] if reported is not None else [],
            "corrected_value": "",
            "eligibility_rule": cand.get("eligibility_rule"),
            "skipped_rows": [],
            "review_reason": review_reason,
            "calculated_value": calculated,
            "approved_value": "",
            "reviewer_note": note,
            "notify_servicer": notify,
        }

        if decision_kind in (_DECISION_REJECT, _DECISION_RETURN):
            # No write. The reported value is kept; the decision is logged
            # with Rows Updated = 0 so the review record is complete.
            log_rows.append({
                **base_row,
                "review_status": _DECISION_LABELS[decision_kind],
                "decision": _DECISION_LABELS[decision_kind],
            })
            if decision_kind == _DECISION_REJECT:
                rejected_count += 1
            else:
                returned_count += 1
            continue

        if decision_kind != _DECISION_APPROVE:
            corrections_not_applied.append({
                "gp": gp, "field": field,
                "reason": f"Unrecognised decision {d.get('decision')!r} — not applied.",
            })
            continue

        # --- APPROVE ---------------------------------------------------------
        # Default the approved value to STAT-01's calculated value; the
        # reviewer may substitute, but only from the field's own domain.
        substituted = d.get("approved_value") not in (None, "")
        approved_input = d.get("approved_value") if substituted else calculated
        # evidence_marker comes off the SERVER-SIDE candidate (main.py stashes
        # these in the session), never off `d` — a caller cannot mint the marker
        # in a decision payload to unlock a BANKRUPTCY write.
        ok, result = validate_approved_value(
            field, approved_input, substituted=substituted,
            evidence_marker=cand.get("evidence_marker"),
        )
        if not ok:
            # A substitution outside the field's domain is refused here — it
            # can never reintroduce a cross-domain value or a terminal status.
            corrections_not_applied.append({
                "gp": gp, "field": field,
                "reason": f"Approved value not written — {result}",
            })
            log_rows.append({
                **base_row,
                "approved_value": approved_input,
                "review_status": "Not applied",
                "review_reason": result,
                "decision": "Approved (value refused)",
            })
            continue
        approved_value = result  # normalized

        corr = dict(cand)
        corr["corrected_value"] = approved_value  # write the APPROVED value, not necessarily the calculated one
        # Explicit, in-band evidence that a reviewer approved THIS correction.
        # Set here and nowhere else -- after the decision was confirmed to be
        # an approval AND the value passed validation. _apply_one_correction
        # refuses a terminal Loan Status without it, so "never AUTOMATICALLY"
        # is enforced on data the write guard can see rather than on an
        # assumption about which branch called it. Any other caller omits the
        # flag and is refused, which is the correct default.
        corr["reviewer_approved"] = True
        written_ok, detail = _apply_one_correction(wb, corr)
        if written_ok:
            verified_overrides[(gp, field)] = approved_value
            log_rows.append({
                **detail,
                "review_status": "Applied",
                "decision": _DECISION_LABELS[_DECISION_APPROVE],
                "calculated_value": calculated,
                "approved_value": approved_value,
                "reviewer_note": note,
                "review_reason": review_reason,
                "notify_servicer": notify,
            })
            applied_count += 1
        else:
            # Approved, but could not be safely written (template drift / GP
            # mismatch / no verifiable row). Never counted as applied; the
            # reason is surfaced, and it is logged as a decided-but-unwritten
            # row so the attempt isn't lost.
            corrections_not_applied.append({"gp": gp, "field": field, "reason": detail})
            log_rows.append({
                **base_row,
                "approved_value": approved_value,
                "review_status": "Not applied",
                "review_reason": detail,
                "decision": "Approved (could not write)",
            })

    if not log_rows:
        return {**empty, "unknown_ids": unknown_ids,
                "skip_reason": "No reviewable decisions were submitted for this servicer."}

    now_iso = datetime.now(timezone.utc).isoformat(timespec="seconds")
    common = {
        "timestamp": now_iso,
        "reviewed_by": format_self_declared_reviewer(reviewer),
        "reviewed_on": reviewed_on or now_iso,
        "source_sha256": source_sha256 or "",
    }
    _add_correction_log_sheet(wb, log_rows, institution, common)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return {
        "file_bytes": buf.getvalue(),
        "filename": _build_corrected_filename(institution, cycle_month),
        "applied_count": applied_count,
        "rejected_count": rejected_count,
        "returned_count": returned_count,
        "corrections_not_applied": corrections_not_applied,
        "unknown_ids": unknown_ids,
        "skip_reason": None,
        "verified_overrides": verified_overrides,
    }
