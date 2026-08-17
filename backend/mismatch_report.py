"""Submitted Information vs. FBC Mismatch Report — CLIENT-CONFIRMED (2026-08-17).

A simple reconciliation artifact for the third party reviewing the lender
submissions: where does the information submitted to FBC differ from FBC's
final determination after Human Review?

THIS IS NOT ANOTHER AUDIT REPORT. It carries no evidence, no calculation and no
explanation beyond the comparison itself.

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
their own, so they are reported as Loan Status with the transition in Comments.
Audit findings that detect a submitted-data problem WITHOUT producing a
corrected value (POP-01, BAL-01/02, PAY-01/02, DUP-01, DQ-01, CONT-01, MEMO-01)
are deliberately out of scope for version 1 and are reserved for a separate
Submitted Data Defects Report.

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

# Shown in place of a value that is absent, so a blank cell is never mistaken
# for a zero or an empty string in the comparison sentence.
ABSENT_DISPLAY = "(none)"

NO_REVIEW_RECORD = "No review record was found for this field."

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

    An absent value renders as "" in the cell and ABSENT_DISPLAY in the
    sentence, so it stays distinguishable from 0.00 in both places."""
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


def _comment(field, submitted_display, final_display, log_row):
    """The deterministic comparison sentence, plus the existing review reason.

    CLIENT-CONFIRMED (2026-08-17): reviewer_note is NOT included. It is
    internal free-form text and this report goes to a third party. Nothing here
    is generated, summarised or rewritten -- review_reason is copied verbatim
    from the review record the application already wrote."""
    sentence = (f"Submitted {field}: {submitted_display or ABSENT_DISPLAY}. "
                f"FBC Determination: {final_display or ABSENT_DISPLAY}.")
    parts = [sentence]
    if log_row is None:
        parts.append(NO_REVIEW_RECORD)
    else:
        reason = log_row.get("review_reason")
        if reason is not None and str(reason).strip():
            parts.append(str(reason).strip())
    return " ".join(parts)


def _index_log_rows(log_rows):
    """{(gp, field): row}. Later rows win, matching write order."""
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
                        provenance_gaps=None):
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

    Rows are sorted by SBA GP Number then MISMATCH_FIELDS order, so output is
    fully determined by the inputs."""
    provenance_gaps = provenance_gaps if provenance_gaps is not None else []
    index = _index_log_rows(log_rows)
    rows = []

    for gp in submitted.index:
        key = str(gp).strip()
        if gp not in final.index:
            continue
        lender_loan_number = submitted.at[gp, "lender_loan_number"] \
            if "lender_loan_number" in submitted.columns else None

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
                "Comments": _comment(field, submitted_display, final_display, log_row),
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
