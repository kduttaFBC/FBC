"""Delinquency Information Request Form — CLIENT-CONFIRMED (2026-08-18).

A post-Human-Review deliverable sent TO the lenders: the loans FBC needs
additional servicing information about. One workbook, one worksheet per uploaded
lender-and-deal combination, each reproducing the client's supplied form.

A PURE REPORTING CONSUMER
-------------------------
Every value comes from data the application has already produced:

    submitted = checks.consolidate_cycle_postings(servicer_inputs_df)
    final     = checks._apply_verified_overrides(submitted, verified_overrides)

so Loan Status, Delinquency Status and Ending Balance are FBC's determination
after Human Review, never a superseded submitted value. This module defines no
business rule, computes no delinquency, ranks no status of its own, alters no
review decision and mutates no input.

SEVERITY IS NOT REDEFINED HERE. "30 DAYS or worse" is evaluated by calling
checks._status_to_missed_count -- the application's existing ladder, in which
CURRENT is 0, DELINQUENT is 1 and 30 DAYS is 2. The threshold is therefore the
single number 2, expressed once. A second hard-coded status list would be a
duplicate ladder that could drift from the real one, which is exactly what this
avoids: a status added upstream is ranked by the application, not by this report.

NOTHING IS SILENTLY DROPPED. A final Delinquency Status that is off-domain
('60-DAYS'), a typo ('DELIQNUENT') or otherwise unparseable QUALIFIES rather than
failing a numeric test in silence -- an unreadable status is precisely a case
needing lender information. The same applies to Loan Status, which is why the
Loan Status rule is written as "not ordinary ACTIVE" rather than as a list of
nine permitted values: a list would silently drop a tenth, unexpected one.

THE NOTES COLUMN IS ALWAYS BLANK. It is the lender's to complete, and this module
never writes to it.
"""

import re

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import checks
import corrections


# ---------------------------------------------------------------------------
# Eligibility
# ---------------------------------------------------------------------------

# CLIENT-CONFIRMED (2026-08-18): "30 DAYS or worse" on the application's own
# ladder, where CURRENT=0, DELINQUENT=1, 30 DAYS=2. So CURRENT and DELINQUENT do
# NOT qualify on delinquency alone, and every more severe permitted value
# (60/90/120/150/150+/180/210 DAYS) qualifies automatically without being listed.
MIN_MISSED_COUNT = 2

ORDINARY_LOAN_STATUS = "ACTIVE"

_BLANK = ("", "NAN", "NONE", "NAT")


def _norm(value):
    """Stripped, upper-cased comparison form. "" for anything absent."""
    if value is None:
        return ""
    if not isinstance(value, str):
        try:
            if pd.isna(value):
                return ""
        except (TypeError, ValueError):
            pass
    text = str(value).strip().upper()
    return "" if text in _BLANK else text


def delinquency_qualifies(value):
    """True when the final Delinquency Status is 30 DAYS or worse, OR cannot be
    read at all.

    A blank does NOT qualify: no delinquency has been asserted, and treating
    silence as severe would sweep in every loan with an empty cell."""
    status = _norm(value)
    if not status:
        return False
    if status not in corrections._PERMITTED_DQ_STATUS:
        return True                      # off-domain -- never dropped silently
    missed = checks._status_to_missed_count(status)
    if missed is None:
        return True                      # unparseable -- never dropped silently
    return missed >= MIN_MISSED_COUNT


def loan_status_qualifies(value):
    """True for any final Loan Status other than ordinary ACTIVE.

    CLIENT-CONFIRMED (2026-08-18). Written as a NEGATION rather than as the nine
    values of corrections._PERMITTED_LOAN_STATUS on purpose: an unexpected or
    misspelled status must be surfaced to the lender, and an allow-list would
    silently exclude it. A blank does not qualify -- there is nothing to ask
    about."""
    status = _norm(value)
    return bool(status) and status != ORDINARY_LOAN_STATUS


def qualifies(loan_status, dq_status):
    """The union. A loan meeting BOTH rules is still one loan, and this returns
    one answer for it -- the caller emits one row per loan, never per rule."""
    return delinquency_qualifies(dq_status) or loan_status_qualifies(loan_status)


# ---------------------------------------------------------------------------
# Row construction
# ---------------------------------------------------------------------------

COLUMNS = [
    "SBA GP Number", "Lender Loan Number", "Securitization Balance",
    "Loan Status", "Delinquency Status", "Last Payment Receipt Date",
    "Servicing Notes / Additional Information",
]

# The lender's column. Written blank on every row, always.
NOTES_COLUMN = COLUMNS[-1]

NO_QUALIFYING_LOANS_TEXT = "No loans requiring additional information."

_RECEIPT_DATE_COLUMN = "Payment Receipt Date"


def _receipt_date_index(servicer_inputs_df):
    """{(sheet, excel row): Payment Receipt Date} over the RAW postings.

    Keyed on the sheet as well as the row number because a workbook may carry
    postings on more than one sheet, where row numbers repeat."""
    if servicer_inputs_df is None or _RECEIPT_DATE_COLUMN not in servicer_inputs_df.columns:
        return {}
    if "_excel_row" not in servicer_inputs_df.columns:
        return {}
    index = {}
    for _i, row in servicer_inputs_df.iterrows():
        key = (row.get("_sheet_name"), row.get("_excel_row"))
        if key[1] is not None:
            index[key] = row.get(_RECEIPT_DATE_COLUMN)
    return index


def _identifier(value):
    """A loan identifier as text, without a float artifact.

    Discovered on real July data: a numeric Lender Loan Number arrives from
    pandas as 2010300146.0 and rendered as "2010300146.0", which is not the
    lender's own reference. An INTEGRAL float is written as its integer; a
    genuinely fractional value is left exactly as it is rather than being
    rounded, and a string identifier (the client's own form carries "0000775",
    leading zeros and all) is untouched."""
    if _norm(value) == "":
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _last_payment_receipt_date(final_row, receipt_index):
    """The receipt date of the posting consolidation ALREADY chose as
    authoritative.

    _excel_row on the consolidated frame is effective_last["_excel_row"] -- the
    one posting Loan Status, Delinquency Status and Ending Balance are all taken
    from. Reading its receipt date reuses that existing decision instead of
    inventing a "last payment" rule, and guarantees the date describes the same
    posting as the three columns beside it. It inherits the incomplete-final-
    posting fallback for free: when the true last posting is blank across every
    authoritative signal, effective_last is the last posting with real data, and
    this reads that one.

    Validated on all 104 qualifying July rows: 58 populated, 46 blank, and ZERO
    disagreements with the latest date anywhere on the loan. The blanks are
    correct -- these are non-paying loans, so no payment was received. Nothing
    reaches backwards into an earlier cycle to fill the cell."""
    value = receipt_index.get((final_row.get("_sheet_name"), final_row.get("_excel_row")))
    return None if _norm(value) == "" else value


def build_lender_rows(final, servicer_inputs_df=None, excluded_gps=None):
    """The qualifying loans for ONE lender-and-deal, in SBA GP Number order.

    `final` is checks._apply_verified_overrides(consolidate_cycle_postings(...),
    verified_overrides) -- supplied by the caller, never built here and never
    mutated. servicer_inputs_df is the same lender's RAW postings, read only for
    the receipt date.

    excluded_gps is the application's own closed-at-zero suppression set,
    captured at /api/audit from STAT-01. It is APPLIED here and never recomputed:
    a loan the audit suppressed is not one to ask a lender about."""
    if final is None or len(final) == 0:
        return []

    suppressed = {str(gp).strip() for gp in (excluded_gps or [])}
    receipt_index = _receipt_date_index(servicer_inputs_df)
    rows = []

    for gp in final.index:
        key = str(gp).strip()
        if key in suppressed:
            continue
        row = final.loc[gp]
        loan_status = row.get("loan_status")
        dq_status = row.get("dq_status")
        if not qualifies(loan_status, dq_status):
            continue

        lender_loan_number = row.get("lender_loan_number")
        rows.append({
            "SBA GP Number": key,
            "Lender Loan Number": _identifier(lender_loan_number),
            "Securitization Balance": row.get("ending_balance"),
            # Written EXACTLY as determined -- an off-domain value is preserved
            # verbatim so the lender sees what their own system reported.
            "Loan Status": "" if loan_status is None else str(loan_status).strip(),
            "Delinquency Status": "" if dq_status is None else str(dq_status).strip(),
            "Last Payment Receipt Date": _last_payment_receipt_date(row, receipt_index),
            NOTES_COLUMN: "",          # the lender's to complete
        })

    rows.sort(key=lambda r: r["SBA GP Number"])
    return rows


# ---------------------------------------------------------------------------
# Workbook
#
# Structure and formatting are REPRODUCED from the client's supplied file, not
# designed here. Every constant below was measured off
# "Lendistry_SOUP 2025-FBC1_Delinquency Additional Information Request Form_1.1
# .xlsx". report.py's helpers are deliberately not reused -- they emit the FBC
# house style, which is a different form.
# ---------------------------------------------------------------------------

TITLE = "Delinquency Information Request Form"

# Column A is a narrow spacer in the supplied file; data begins at column B.
FIRST_COLUMN = 2
TITLE_ROW = 1
HEADER_ROW = 4
FIRST_DATA_ROW = 5

COLUMN_WIDTHS = {"A": 1.86, "B": 12.57, "C": 15.14, "D": 12.71,
                 "E": 10.29, "F": 12.00, "G": 13.29, "H": 85.86}
ROW_HEIGHTS = {1: 30.0, 2: 4.5, 3: 12.0, 4: 45.0}

FONT_NAME = "Aptos Narrow"
TITLE_FONT = Font(name=FONT_NAME, size=16, bold=True)
HEADER_FONT = Font(name=FONT_NAME, size=11, bold=False)
DATA_FONT = Font(name=FONT_NAME, size=10, bold=False)

# Theme references, not hex: the supplied file uses theme colours, so writing
# the same theme index and tint reproduces it under any workbook theme.
HEADER_FILL = PatternFill("solid", fgColor=Color(theme=4, tint=0.8))
BAND_FILL = PatternFill("solid", fgColor=Color(theme=0, tint=-0.05))
_THIN = Side(style="thin", color=Color(theme=3))
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

CENTER = Alignment(horizontal="center", vertical="center")
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)

BALANCE_FORMAT = "#,##0"
DATE_FORMAT = "mm-dd-yy"

BALANCE_HEADING = "Securitization Balance"

# Verbatim from the supplied form, including its missing space after "Notes".
NOTES_HEADING = (
    "Servicing Notes /Additional Information - Provide notes from the servicing "
    "system such as last date of Borrower contact, reason(s) affecting the "
    "Borrower's ability to make payments, etc."
)

_INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")


def balance_heading(cycle_month=None):
    """"Securitization Balance as of 7/31/26" for July 2026.

    PRESENTATION ONLY (CLIENT-CONFIRMED 2026-08-18). The month end is derived
    here, at write time, purely for this heading; no figure is computed from it
    and nothing else in the application reads it. cycle_month is the same
    Timestamp the filenames already use, normalised to the 1st, so the month end
    is pandas' own month-end offset rather than a hand-rolled calendar.

    With no cycle month the heading degrades to the bare label -- never a guessed
    date."""
    if cycle_month is None:
        return BALANCE_HEADING
    try:
        end = pd.Timestamp(cycle_month) + pd.offsets.MonthEnd(0)
    except (ValueError, TypeError):
        return BALANCE_HEADING
    return (f"{BALANCE_HEADING} as of "
            f"{end.month}/{end.day}/{end.strftime('%y')}")


def worksheet_title(display_label, taken=None):
    """An Excel-legal sheet name from the application's own lender identity.

    display_label is what the Monthly Servicer Report, Portfolio Summary and
    chat already use ("Lendistry 2024", "Lendistry 2025"), so a lender in two
    deals gets two sheets and the two securitizations are never merged."""
    name = _INVALID_SHEET_CHARS.sub("", str(display_label or "").strip()) or "Lender"
    name = name[:31]
    if taken is None:
        return name
    candidate, n = name, 2
    while candidate in taken:
        suffix = f" ({n})"
        candidate = name[:31 - len(suffix)] + suffix
        n += 1
    return candidate


def _write_sheet(ws, rows, cycle_month=None):
    ws.cell(row=TITLE_ROW, column=FIRST_COLUMN, value=TITLE).font = TITLE_FONT

    headings = list(COLUMNS)
    headings[2] = balance_heading(cycle_month)
    headings[6] = NOTES_HEADING
    for offset, heading in enumerate(headings):
        cell = ws.cell(row=HEADER_ROW, column=FIRST_COLUMN + offset, value=heading)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER_WRAP
        cell.border = BORDER

    for letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[letter].width = width
    for row, height in ROW_HEIGHTS.items():
        ws.row_dimensions[row].height = height

    if not rows:
        # Named, not omitted: a lender with nothing outstanding is told so.
        cell = ws.cell(row=FIRST_DATA_ROW, column=FIRST_COLUMN,
                       value=NO_QUALIFYING_LOANS_TEXT)
        cell.font = DATA_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center")
        return

    for i, row in enumerate(rows):
        excel_row = FIRST_DATA_ROW + i
        banded = (excel_row % 2) == 0          # even rows filled, as supplied
        for offset, column in enumerate(COLUMNS):
            value = row.get(column)
            if column == NOTES_COLUMN:
                value = None               # ALWAYS blank -- the lender's column
            elif column == "Last Payment Receipt Date" and value is not None:
                parsed = pd.to_datetime(value, errors="coerce")
                value = None if pd.isna(parsed) else parsed.to_pydatetime()
            elif column == "Securitization Balance" and value is not None:
                try:
                    value = float(value)
                except (TypeError, ValueError):
                    pass
            cell = ws.cell(row=excel_row, column=FIRST_COLUMN + offset, value=value)
            cell.font = DATA_FONT
            cell.alignment = CENTER
            cell.border = BORDER
            if banded:
                cell.fill = BAND_FILL
            if column == "Securitization Balance":
                cell.number_format = BALANCE_FORMAT
            elif column == "Last Payment Receipt Date":
                cell.number_format = DATE_FORMAT


def generate_delinquency_request_form_bytes(blocks, buffer, cycle_month=None):
    """The workbook. `blocks` is [(display_label, rows), ...] in upload order.

    One worksheet per lender-and-deal, every sheet present -- a lender with no
    qualifying loans keeps its sheet and says so. Computes nothing: every value
    arrives finished."""
    wb = Workbook()
    wb.remove(wb.active)
    taken = set()
    for display_label, rows in blocks:
        title = worksheet_title(display_label, taken)
        taken.add(title)
        _write_sheet(wb.create_sheet(title), rows or [], cycle_month)
    if not wb.sheetnames:
        # openpyxl cannot save a workbook with no sheets.
        _write_sheet(wb.create_sheet("Lenders"), [], cycle_month)
    wb.save(buffer)
    buffer.seek(0)


def delinquency_request_form_filename(cycle_month=None):
    """e.g. "Delinquency_Information_Request_Form_July_2026.xlsx". Omits the
    period rather than guessing one when the cycle month is unknown."""
    if cycle_month is not None:
        try:
            return ("Delinquency_Information_Request_Form_"
                    f"{cycle_month.strftime('%B_%Y')}.xlsx")
        except AttributeError:
            pass
    return "Delinquency_Information_Request_Form.xlsx"
