"""Submitted Information vs. FBC Mismatch Report — CLIENT-CONFIRMED (2026-08-17).

A simple reconciliation artifact for the third party reviewing the lender
submissions: where does the information submitted to FBC differ from FBC's
final determination after Human Review?

THIS IS NOT ANOTHER AUDIT REPORT. It performs no calculation and reaches no
conclusion of its own. Comments are BLANK BY DEFAULT and carry supporting
context only where the two comparison columns cannot: an unsupported
liquidation, a documented Servicer Memo deferral's own dates, a cascade
clearing, a restored prior-period balance. That context is read from evidence
the application already computed -- read-only reporting context, which
CLIENT-CONFIRMED (2026-08-17) never participates in a calculation, a Human
Review decision, validation, eligibility or an FBC determination.

A PURE REPORTING CONSUMER
-------------------------
Every value it reports is read from two frames the application has already
built:

    submitted = checks.consolidate_cycle_postings(servicer_inputs_df)
    final     = checks._apply_verified_overrides(submitted, verified_overrides)

The second is the same call the Final Audited Portfolio Summary and the Monthly
Servicer Report already make. This module therefore adds no business rule,
recalculates no delinquency, infers no Loan Status, reclassifies no
recommendation, alters no verified override, calls no Human Review logic,
mutates no input frame, touches no session state and imports no frontend code.
There is one authoritative path -- submitted data, existing analysis, Human
Review, existing final values -- and this reads the end of it.

SCOPE: exactly the five fields a reviewer-approved correction can change.
DEFERRED, FORECLOSURE and LIQUIDATION are VALUES of Loan Status, not fields of
their own, so they are reported as Loan Status -- the transition itself is shown
by the Submitted Information and FBC Determination columns, never restated in
Comments. Audit findings that detect a submitted-data problem WITHOUT producing
a corrected value (POP-01, BAL-01/02, PAY-01/02, DUP-01, DQ-01, CONT-01,
MEMO-01) are out of scope for version 1.

POP-01 Missing Loan is a PLANNED requirement, deliberately PARKED
(CLIENT-CONFIRMED 2026-08-17): its findings ranged from 360 across the June
submissions to a single July record that was a reader artifact (a stray "," read
as an SBA GP Number), so it must be established as reliably representing genuine
missing loans before it may enter a client-facing reconciliation.

CASCADE CHANGES ARE NOT COLLAPSED. One reviewer decision may legitimately move
Loan Status, Ending Balance, Liquidation Date and Liquidation Loss. Each is a
real difference between submitted and final, so each is its own row. Deciding
that one difference "caused" another would be exactly the causal rule logic this
module must not contain.
"""

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import pandas as pd
from openpyxl import Workbook

import checks
from report import _autofit, _write_header, _write_rows


# The five fields a reviewer-approved correction can change. Kept in the order
# a liquidation reversal reads naturally, which is also the row sort order.
MISMATCH_FIELDS = (
    "Loan Status",
    "Delinquency Status",
    "Ending Balance",
    "Liquidation Date",
    "Liquidation Loss",
)

# Guard, not documentation: the report cannot drift from what an override is
# actually able to change. checks._VERIFIED_OVERRIDE_COLUMNS is the single
# definition of that mapping and this module never restates it.
assert set(MISMATCH_FIELDS) == set(checks._VERIFIED_OVERRIDE_COLUMNS), (
    "MISMATCH_FIELDS must match checks._VERIFIED_OVERRIDE_COLUMNS exactly"
)

MISMATCH_COLUMNS = [
    "Lender", "SBA GP Number", "Lender Loan Number",
    "Submitted Information", "FBC Determination", "Mismatch Type", "Comments",
]

# CLIENT-CONFIRMED (2026-08-17): exactly two worksheets, always both present.
WORKSHEET_2024 = "SOUP 2024 FBC1"
WORKSHEET_2025 = "SOUP 2025 FBC1"
WORKSHEETS = (WORKSHEET_2024, WORKSHEET_2025)

NO_MISMATCHES_TEXT = "No mismatches were identified for this deal."

_STATUS_FIELDS = ("Loan Status", "Delinquency Status")
_MONEY_FIELDS = ("Ending Balance", "Liquidation Loss")
_DATE_FIELDS = ("Liquidation Date",)

# One sentinel for every way a value can be absent. Deliberately NOT equal to
# Decimal("0.00") -- see _comparable.
_ABSENT = object()


# ---------------------------------------------------------------------------
# Value handling
#
# TWO VALUES PER CELL, and the split is the whole point:
#   * a COMPARISON KEY, used only to decide equal / not equal
#   * a DISPLAY VALUE, the raw value rendered
# Normalisation can therefore suppress a difference that is only formatting,
# and can never change what a reader sees or what a figure means.
# ---------------------------------------------------------------------------

def _is_absent(value):
    """None, empty string, whitespace-only, NaN and NaT all mean "no value".

    The liquidation cascade clears a date to "" and a loss to None in the same
    reviewer decision; treating those as different would report one
    determination as two."""
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return False


def _comparable(field, value):
    """The comparison key. Returns _ABSENT, or a (kind, value) pair.

    The kind tag prevents cross-type collisions -- money 0.00 must never
    compare equal to the text "0.00".

    WHAT IS NORMALISED, and why none of it changes financial meaning:
      status  strip + uppercase. SOUP genuinely carries "Active", "ACTIVE" and
              "ACTiVE" for one status (corrections._canonical_in documents
              this); a case difference is not a determination.
      money   Decimal quantised to 2dp, the reporting granularity. Removes
              float representation noise (145729.870000001), nothing else.
      date    the date only, so "2026-03-27", a Timestamp and a date object
              agree. An unparseable value falls back to text rather than being
              silently coerced.

    WHAT IS DELIBERATELY NOT NORMALISED: sign, magnitude, rounding beyond 2dp,
    and -- most importantly -- ABSENT IS NOT ZERO. "Balance is 0.00" and "no
    balance recorded" are different financial facts and always compare
    unequal."""
    if _is_absent(value):
        return _ABSENT
    if field in _MONEY_FIELDS:
        try:
            return ("money", Decimal(str(float(value))).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP))
        except (TypeError, ValueError, InvalidOperation):
            return ("text", str(value).strip().upper())
    if field in _DATE_FIELDS:
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.isna(parsed):
            return ("text", str(value).strip().upper())
        return ("date", parsed.date())
    return ("text", str(value).strip().upper())


def _display(field, value):
    """The value as the reader sees it. Never the comparison key.

    An absent value renders as an empty cell. It stays distinguishable from 0.00
    because _comparable -- not this function -- decides equality, and there
    absence and zero never compare equal."""
    if _is_absent(value):
        return ""
    if field in _MONEY_FIELDS:
        try:
            return f"{float(value):,.2f}"
        except (TypeError, ValueError):
            return str(value).strip()
    if field in _DATE_FIELDS:
        parsed = pd.to_datetime(value, errors="coerce")
        if not pd.isna(parsed):
            return parsed.date().isoformat()
        return str(value).strip()
    return str(value).strip()


# ---------------------------------------------------------------------------
# Comments -- CLIENT-CONFIRMED (2026-08-17), revised
#
# BLANK BY DEFAULT. Submitted Information and FBC Determination already show the
# mismatch, so restating them in prose is noise in a seven-column table. A
# Comment appears ONLY where it carries context those two columns cannot.
#
# Nothing here is generated, summarised or inferred. Every Comment is a fixed
# sentence selected by a deterministic test over evidence the application has
# ALREADY computed, and every date printed is a value that already exists --
# rendered exactly as held, never parsed, reformatted or derived.
#
# DELIBERATELY REMOVED: the submitted-vs-FBC restatement, and review_reason.
# Those reasons run 299-524 characters of prose written to instruct an FBC
# REVIEWER ("select the Loan Status you believe is correct"), addressed to
# someone who is not the reader of this report.
# ---------------------------------------------------------------------------

COMMENT_UNSUPPORTED_LIQUIDATION = "No supporting data found for reported liquidation."
COMMENT_CLEARED_UNSUPPORTED = "Cleared because the reported liquidation was not supported."
COMMENT_BALANCE_RESTORED = "Restored to the prior period ending balance."

# The five fields whose presence constitutes support for a reported liquidation.
# ALL must be absent before this report will say none was found. Read-only
# reporting context (CLIENT-CONFIRMED 2026-08-17): these never participate in a
# calculation, a Human Review decision, validation, eligibility, or an FBC
# determination -- STAT-01 decided all of that long before this module runs.
LIQUIDATION_EVIDENCE_FIELDS = (
    "liquidation_date", "liquidation_proceeds", "liquidation_loss",
    "foreclosure_date", "reo_date",
)

_LIQUIDATION = "LIQUIDATION"
_DEFERRED = "DEFERRED"


def _text(value):
    """Upper-cased comparison form, or "" when absent."""
    return "" if _is_absent(value) else str(value).strip().upper()


def _deferral_sentence(evidence):
    """The Servicer Memo's own deferral dates, or "".

    Dates are printed EXACTLY as the application holds them. A date that does
    not exist is simply not mentioned -- nothing is calculated, inferred or
    manufactured to fill a gap, and a documented deferral carrying no date at
    all yields no Comment rather than an empty claim."""
    if str(evidence.get("evidence_status") or "").strip() != checks.EVIDENCE_DOCUMENTED:
        return ""
    start = "" if _is_absent(evidence.get("deferral_start")) else str(evidence["deferral_start"]).strip()
    end = "" if _is_absent(evidence.get("deferral_end")) else str(evidence["deferral_end"]).strip()
    resume = "" if _is_absent(evidence.get("deferral_resume")) else str(evidence["deferral_resume"]).strip()

    if start and end:
        base = f"Servicer Memo records a deferral from {start} to {end}"
    elif start:
        base = f"Servicer Memo records a deferral beginning {start}"
    elif end:
        base = f"Servicer Memo records a deferral ending {end}"
    elif resume:
        # The only date on record. Reported on its own rather than dropped.
        return f"Servicer Memo records regular payments resuming {resume}."
    else:
        return ""
    return f"{base}; regular payments resume {resume}." if resume else f"{base}."


def _comment(field, submitted_value, final_value, unsupported_liquidation, evidence):
    """The Comment for one mismatch row, or "" -- which is the normal case.

    unsupported_liquidation is computed ONCE per loan by the caller, so the
    Loan Status row and the Liquidation Date / Liquidation Loss rows it cascades
    into cannot describe the same liquidation differently."""
    if field == "Loan Status":
        if unsupported_liquidation:
            return COMMENT_UNSUPPORTED_LIQUIDATION
        # A deferral FBC established. Blank when the memo documents no dates:
        # "FBC Determination: DEFERRED" already says the rest.
        if _text(final_value) == _DEFERRED:
            return _deferral_sentence(evidence)
        return ""

    if field in ("Liquidation Date", "Liquidation Loss"):
        # Cleared as a CASCADE of the Loan Status decision. Says why the value
        # went away rather than repeating the value that went away.
        if unsupported_liquidation and _is_absent(final_value) and not _is_absent(submitted_value):
            return COMMENT_CLEARED_UNSUPPORTED
        return ""

    if field == "Ending Balance":
        # The reviewer restored the balance the liquidation reversal had zeroed.
        # Recognised by equality with the server's own single restorable amount
        # -- never by re-deriving what the balance "should" be.
        restore = evidence.get("balance_restore_value")
        if not _is_absent(restore) and _comparable(field, final_value) == _comparable(field, restore):
            return COMMENT_BALANCE_RESTORED
        return ""

    # Delinquency Status, and anything added later: blank.
    return ""


_EVIDENCE_KEYS = LIQUIDATION_EVIDENCE_FIELDS + (
    "evidence_status", "deferral_start", "deferral_end", "deferral_resume",
    "balance_restore_value",
)


def _index_evidence(candidates):
    """{gp: {evidence key: value}}, merged per LOAN rather than per field.

    Keyed by SBA GP Number on purpose. Liquidation Date and Liquidation Loss are
    cleared as a CASCADE of the Loan Status decision, so no candidate exists
    carrying `field == "Liquidation Date"` -- those rows can only find their
    context under the loan.

    Merging is safe because these values are per-loan facts STAT-01 passes
    through onto every candidate for that loan, not per-field ones. The first
    non-absent value wins, so a candidate that happens to omit a key cannot
    blank out one that carries it."""
    index = {}
    for cand in candidates or []:
        gp = str(cand.get("gp") or "").strip()
        if not gp:
            continue
        bucket = index.setdefault(gp, {})
        for key in _EVIDENCE_KEYS:
            if key not in bucket or _is_absent(bucket.get(key)):
                value = cand.get(key)
                if not _is_absent(value):
                    bucket[key] = value
    return index


def _index_log_rows(log_rows):
    """{(gp, field): row}. Later rows win, matching write order.

    PROVENANCE DETECTION ONLY. Since 2026-08-17 nothing on these rows reaches
    the workbook: Comments no longer carry review_reason, and never carried
    reviewer_note. The lookup survives solely to answer "does a review record
    exist for this change", which the caller surfaces as a server-side
    diagnostic. A missing record changes no cell."""
    index = {}
    for row in log_rows or []:
        gp = str(row.get("gp") or "").strip()
        field = str(row.get("field") or "").strip()
        if gp and field:
            index[(gp, field)] = row
    return index


# ---------------------------------------------------------------------------
# Row construction
# ---------------------------------------------------------------------------

def build_mismatch_rows(submitted, final, soup_key, log_rows=None,
                        provenance_gaps=None, candidates=None):
    """One row per (loan, field) where the submitted value differs from FBC's
    final determination, for ONE servicer.

    submitted is checks.consolidate_cycle_postings(servicer_inputs_df); final is
    checks._apply_verified_overrides(submitted, verified_overrides). BOTH are
    supplied by the caller -- neither is built here, and neither is mutated:
    every access is a read.

    soup_key is the servicer identity the caller resolved from the upload
    filename convention, the same value the Monthly Servicer Report writes to
    Lender/Servicer Name. Deliberately NOT the SRF's own "Servicer:" cell, which
    for one servicer reads "Community Bank & Trust-West Georgia" and would not
    match.

    provenance_gaps, when given, collects any changed field with no matching
    review record. The row is STILL emitted -- the difference is real whatever
    its provenance, and suppressing it would hide a genuine mismatch. This
    module only reports the gap; how it surfaces is the caller's decision.

    candidates are the session's own review candidates, read ONLY for Comment
    context (deferral dates, liquidation evidence, the restorable balance).
    CLIENT-CONFIRMED (2026-08-17): reporting context only -- nothing here feeds
    a calculation, a Human Review decision, validation, eligibility or an FBC
    determination, all of which were settled upstream before this runs.

    Rows are sorted by SBA GP Number then MISMATCH_FIELDS order, so output is
    fully determined by the inputs."""
    provenance_gaps = provenance_gaps if provenance_gaps is not None else []
    index = _index_log_rows(log_rows)
    evidence_index = _index_evidence(candidates)
    rows = []

    status_column = checks._VERIFIED_OVERRIDE_COLUMNS["Loan Status"]

    for gp in submitted.index:
        key = str(gp).strip()
        if gp not in final.index:
            continue
        lender_loan_number = submitted.at[gp, "lender_loan_number"] \
            if "lender_loan_number" in submitted.columns else None
        evidence = evidence_index.get(key, {})

        # Computed ONCE per loan, so the Loan Status row and every row its
        # decision cascades into cannot describe the same liquidation
        # differently.
        #
        # STRICTLY ASYMMETRIC (CLIENT-CONFIRMED 2026-08-17). Absence of evidence
        # is observable; a reviewer's alternative rationale is not. So this says
        # "no supporting data" only when the servicer reported LIQUIDATION, FBC
        # determined otherwise, and EVERY supporting field is absent. If any
        # evidence exists the Comment stays blank rather than inferring why the
        # reviewer decided as they did. Scoped to LIQUIDATION alone -- no
        # general terminal-status rule is created here.
        unsupported_liquidation = False
        if status_column in submitted.columns and status_column in final.columns:
            unsupported_liquidation = (
                _text(submitted.at[gp, status_column]) == _LIQUIDATION
                and _text(final.at[gp, status_column]) != _LIQUIDATION
                and all(_is_absent(evidence.get(f)) for f in LIQUIDATION_EVIDENCE_FIELDS)
            )

        for field in MISMATCH_FIELDS:
            column = checks._VERIFIED_OVERRIDE_COLUMNS[field]
            if column not in submitted.columns or column not in final.columns:
                continue
            submitted_value = submitted.at[gp, column]
            final_value = final.at[gp, column]
            if _comparable(field, submitted_value) == _comparable(field, final_value):
                continue

            log_row = index.get((key, field))
            if log_row is None:
                provenance_gaps.append({
                    "servicer": soup_key, "SBA GP Number": key, "field": field,
                    "message": (
                        f"{soup_key} GP {key}: the final FBC determination for {field} "
                        f"differs from the submitted value, but no Human Review record "
                        f"was found for it. The difference is reported; its provenance "
                        f"could not be established."
                    ),
                })

            submitted_display = _display(field, submitted_value)
            final_display = _display(field, final_value)
            rows.append({
                "Lender": soup_key,
                "SBA GP Number": key,
                "Lender Loan Number": "" if _is_absent(lender_loan_number)
                                      else str(lender_loan_number).strip(),
                "Submitted Information": submitted_display,
                "FBC Determination": final_display,
                "Mismatch Type": field,
                # Built from the RAW values and the loan's evidence, never from
                # the display strings -- a Comment must not depend on how a cell
                # happens to be rendered.
                "Comments": _comment(field, submitted_value, final_value,
                                     unsupported_liquidation, evidence),
            })

    rows.sort(key=lambda r: (r["SBA GP Number"],
                             MISMATCH_FIELDS.index(r["Mismatch Type"])))
    return rows


def worksheet_label(deal):
    """"SOUP2024_1" -> "SOUP 2024 FBC1", or None when the deal is missing,
    malformed, or outside the two supported deals.

    Parsed by checks._deal_year_and_series -- the same strict parser POP-01 and
    the Portfolio Summary grouping already use. No filename inference, no
    servicer-name lookup, no guessing at a year."""
    parsed = checks._deal_year_and_series(deal)
    if parsed is None:
        return None
    label = f"SOUP {parsed[0]} FBC{parsed[1]}"
    return label if label in WORKSHEETS else None


def group_by_deal(rows_by_servicer, deals, unassigned=None):
    """{worksheet label: rows}, with BOTH worksheets always present.

    rows_by_servicer and deals are parallel lists in caller (upload) order --
    the same call shape checks.group_portfolio_summary_by_deal takes. Servicer
    order is preserved within each worksheet, so ordering is Lender (upload),
    then SBA GP Number, then field.

    CLIENT-CONFIRMED (2026-08-17): a servicer that cannot be assigned
    deterministically to one of the two approved deals is NOT guessed at, NOT
    misfiled into either sheet, and does NOT get a third sheet. Its rows are
    returned in `unassigned` for the caller to handle. This module detects the
    condition; it does not decide how it is surfaced."""
    unassigned = unassigned if unassigned is not None else []
    blocks = {label: [] for label in WORKSHEETS}

    for rows, deal in zip(rows_by_servicer, deals):
        label = worksheet_label(deal)
        if label is None:
            for row in rows or []:
                unassigned.append({"servicer": row.get("Lender"), "deal": deal, "row": row})
            continue
        blocks[label].extend(rows or [])
    return blocks


# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------

def generate_mismatch_report_bytes(blocks, buffer):
    """The two-worksheet workbook. Computes nothing -- every value arrives
    finished, and a deal with no mismatches keeps its worksheet and says so."""
    wb = Workbook()
    first = True
    for label in WORKSHEETS:
        ws = wb.active if first else wb.create_sheet(label)
        if first:
            ws.title = label
            first = False
        rows = blocks.get(label) or []
        _write_header(ws, MISMATCH_COLUMNS)
        if rows:
            _write_rows(
                ws, MISMATCH_COLUMNS,
                [[r.get(c, "") for c in MISMATCH_COLUMNS] for r in rows],
                currency_cols=set(), percentage_cols=set(),
            )
            _autofit(ws, MISMATCH_COLUMNS, [[r.get(c, "") for c in MISMATCH_COLUMNS]
                                            for r in rows])
        else:
            ws.cell(row=2, column=1, value=NO_MISMATCHES_TEXT)
            _autofit(ws, MISMATCH_COLUMNS, [[NO_MISMATCHES_TEXT] + [""] * 6])
        ws.freeze_panes = "A2"
    wb.save(buffer)
    buffer.seek(0)


def mismatch_report_filename(cycle_month=None):
    """e.g. "Submitted_vs_FBC_Mismatch_Report_June_2026.xlsx". Omits the period
    rather than guessing one when the cycle month is unknown."""
    if cycle_month is not None:
        try:
            return f"Submitted_vs_FBC_Mismatch_Report_{cycle_month.strftime('%B_%Y')}.xlsx"
        except AttributeError:
            pass
    return "Submitted_vs_FBC_Mismatch_Report.xlsx"
