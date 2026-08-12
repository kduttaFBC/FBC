"""Monthly Servicer Report — CLIENT-CONFIRMED (2026-08-11).

A POST-REVIEW consolidation of reviewer-approved lender data into the client's
own report format. ONE workbook per SOUP deal:

    8 lender submissions -> Audit -> Human Review -> Corrected Lender Workbooks
      -> Consolidated Monthly Servicer Reports -> Final Audited Portfolio Summary

Structure is taken verbatim from the client's supplied workbook
"2026.06_SOUP 7a 2025-FBC1_Monthly Servicer Report_June_2026_v2.xlsx": two
worksheets, a metadata block, a SUBTOTAL row above the headers, 21 loan-level
columns, and a Remittance Amount sheet whose amounts are SUMIFS formulas
against the loan sheet rather than pasted values.

THIS IS CONSOLIDATION, NOT A CALCULATION.
-----------------------------------------
Nothing here computes a financial value. Every figure is the reviewer-approved
lender value, passed through unrounded with the client's display format applied
(CLIENT-CONFIRMED 2026-08-11 -- the supplied reference workbook stores values
rounded to 2dp, and in at least one column re-derived from OTHER rounded
columns; reproducing that would be a second audit calculation, which is exactly
what this module must not be). Consolidating a loan's multiple postings reuses
checks.consolidate_cycle_postings -- the single shared implementation -- rather
than re-deriving the rule. Verified against the client's own workbook on six
Lendistry loans where SUM differs from LAST: all six agree with SUM.

DATA SOURCE
-----------
The reviewer-approved values, i.e. the lender's submitted frame plus the
verified overrides that were actually WRITTEN into the corrected workbook.
Deliberately NOT a re-read of the generated corrected workbook bytes: this
app's own corrected output is precisely what preflight.py blocks from
re-ingestion (openpyxl leaves no value cache), so parsing it back would be
fragile in exactly the way that control exists to prevent. Both artifacts are
therefore built from ONE verified layer -- the same
`verified_overrides_by_servicer` map the Final Audited Portfolio Summary
consumes -- which is what makes the two reconcile by construction.

Only Loan Status and Delinquency Status can carry an override. That is the
existing correction allow-list and this module does not widen it.

POPULATION
----------
Reviewed loans MINUS closed-at-zero suppressions, so the consolidated
population matches the Final Audited population exactly (CLIENT-CONFIRMED
2026-08-11). Measured against the client's supplied workbook: the 2025 deal has
1,437 distinct reviewed loans and the client's report has 1,436 rows; the one
omitted loan is CB&T GP 4338609104, which is precisely the loan the audit's own
closed-at-zero rule suppresses.

BOUNDARIES
----------
Reads from checks/file_readers; writes nothing back. CONT-01 and every other
audit check, the Original and Audited and Final Audited Portfolio Summaries,
corrections.py and the corrected workbooks are all untouched by this module.
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

import checks


MONTHLY_SERVICER_REPORT_SHEET = "Monthly Servicer Report"
REMITTANCE_AMOUNT_SHEET = "Remittance Amount"

REPORT_TITLE = "SOUP Monthly Servicer Report"

# The client's exact header row, in the client's exact order and spelling.
# "Modification  Date" carries TWO spaces in the supplied workbook; reproduced
# verbatim rather than tidied, because the header text is part of the format
# the client asked to be preserved.
MSR_COLUMNS = [
    "SBA GP Number", "Lender Loan Number", "Loan Status", "Delinquency Status",
    "Net Interest Rate", "Net Interest", "Principal", "Total Payment Received",
    "Beginning Actual Balance", "Ending Actual Balance", "Paid-In-Full Date",
    "Liquidation Date", "Foreclosure Date", "REO Date", "Liquidation Proceeds",
    "Subsequent Recoveries/Losses", "Cumulative Recoveries/Losses",
    "Liquidation Loss", "Modification  Date", "Modification Comments",
    "Lender/Servicer Name",
]

REMITTANCE_COLUMNS = ["Servicer", "Amount ($)", "Note Account", "Net Amount ($)"]

# Worksheet geometry, matching the supplied workbook exactly.
_TITLE_ROW = 1
_NOTE_SERIES_ROW = 2
_REMITTANCE_DATE_ROW = 3
_COLLECTION_PERIOD_ROW = 4
_SUBTOTAL_ROW = 7
_HEADER_ROW = 8
_FIRST_DATA_ROW = 9

# SUBTOTAL(9, ...) sums, on exactly the columns the client's workbook sums.
# SUBTOTAL(2, A...) counts. Ranges are built from the ACTUAL row count, never
# the template's hardcoded 1513 (which overshot its own 1444 rows).
_SUBTOTAL_SUM_COLUMNS = ("Net Interest", "Principal", "Total Payment Received",
                         "Beginning Actual Balance", "Ending Actual Balance",
                         "Liquidation Proceeds", "Subsequent Recoveries/Losses",
                         "Cumulative Recoveries/Losses", "Liquidation Loss")

_CURRENCY = '"$"#,##0.00'
_ACCOUNTING = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
_DATE = "mm-dd-yy"
_PERCENT = "0.00%"
_COUNT = "0"

# Column -> number format. The supplied workbook applies a currency format to
# several TEXT columns (Loan Status, Modification Comments, Lender/Servicer
# Name) -- a copy-paste artifact that would render text as $0.00 if a value
# were ever numeric. Those use General here; every genuinely numeric column
# keeps the client's format exactly.
_NUMBER_FORMATS = {
    "SBA GP Number": _COUNT,
    "Net Interest Rate": _PERCENT,
    "Net Interest": _CURRENCY,
    "Principal": _CURRENCY,
    "Total Payment Received": _CURRENCY,
    "Beginning Actual Balance": _CURRENCY,
    "Ending Actual Balance": _CURRENCY,
    "Paid-In-Full Date": _DATE,
    "Liquidation Date": _DATE,
    "Foreclosure Date": _DATE,
    "REO Date": _DATE,
    "Liquidation Proceeds": _ACCOUNTING,
    "Subsequent Recoveries/Losses": _ACCOUNTING,
    "Cumulative Recoveries/Losses": _ACCOUNTING,
    "Liquidation Loss": _ACCOUNTING,
    "Modification  Date": _DATE,
}

_COLUMN_WIDTHS = {
    "SBA GP Number": 20.7, "Lender Loan Number": 16.0, "Loan Status": 20.1,
    "Delinquency Status": 13.4, "Net Interest Rate": 14.6, "Net Interest": 14.6,
    "Principal": 14.6, "Total Payment Received": 20.0,
    "Beginning Actual Balance": 20.9, "Ending Actual Balance": 14.6,
    "Paid-In-Full Date": 14.6, "Liquidation Date": 14.6, "Foreclosure Date": 14.6,
    "REO Date": 14.6, "Liquidation Proceeds": 14.6,
    "Subsequent Recoveries/Losses": 20.6, "Cumulative Recoveries/Losses": 20.7,
    "Liquidation Loss": 14.6, "Modification  Date": 14.6,
    "Modification Comments": 39.6, "Lender/Servicer Name": 26.6,
}

# Servicer Inputs source column for each output column. None means the field
# has no source in the uploaded files -- see UNSOURCED_COLUMNS.
_SOURCE_COLUMNS = {
    "Net Interest Rate": "Net Interest Rate",
    "Paid-In-Full Date": "Paid-In-Full Date",
    "Liquidation Date": "Liquidation Date",
    "Foreclosure Date": "Foreclosure Date",
    "REO Date": "REO Date",
    "Liquidation Proceeds": "Liquidation Proceeds",
    "Subsequent Recoveries/Losses": "Subsequent Recoveries/(Losses)",
    "Liquidation Loss": "Liquidation Loss",
    "Modification  Date": "Modification  Date",
    "Modification Comments": "Modification Comments (Free-form)",
}

# CLIENT-CONFIRMED (2026-08-11): fields with NO source in the uploaded lender
# files. Written blank, never inferred, manufactured or substituted.
#   Cumulative Recoveries/Losses -- absent from Servicer Inputs entirely. The
#     Data Validation sheet offers "Cumulative Principal Recovered", a
#     different quantity. Blank for all 1,436 rows of the client's own report.
#   Note Account -- no source anywhere; blank in the client's report. The Net
#     Amount formula =B-C still evaluates against a blank C.
UNSOURCED_COLUMNS = ("Cumulative Recoveries/Losses",)
UNSOURCED_REMITTANCE_COLUMNS = ("Note Account",)

_DATE_COLUMNS = ("Paid-In-Full Date", "Liquidation Date", "Foreclosure Date",
                 "REO Date", "Modification  Date")

# Validation issue codes surfaced to the reviewer.
ISSUE_COLLECTION_PERIOD_UNAVAILABLE = "collection_period_unavailable"
ISSUE_NOTE_SERIES_UNAVAILABLE = "note_series_unavailable"
ISSUE_NOTE_SERIES_CONFLICT = "note_series_conflict"
ISSUE_REMITTANCE_DATE_UNAVAILABLE = "remittance_date_unavailable"
ISSUE_REMITTANCE_DATE_CONFLICT = "remittance_date_conflict"
ISSUE_POSTINGS_DISAGREE = "postings_disagree"


def _blank(value):
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() in ("", "nan", "NaT", "None")


def _posting_period_label(row):
    """How one posting identifies itself, for a conflict message: its accrual
    window, its payment instance, and its worksheet row."""
    parts = []
    instance = row.get("Loan Payment Instance")
    if not _blank(instance):
        try:
            parts.append(f"posting {int(float(instance))}")
        except (TypeError, ValueError):
            parts.append(f"posting {instance}")
    interest_from, interest_to = row.get("Interest-From"), row.get("Interest-To")
    if not _blank(interest_from) and not _blank(interest_to):
        parts.append(f"interest {str(interest_from)[:10]} to {str(interest_to)[:10]}")
    elif not _blank(row.get("Payment Receipt Date")):
        parts.append(f"received {str(row.get('Payment Receipt Date'))[:10]}")
    excel_row = row.get("_excel_row")
    if not _blank(excel_row):
        try:
            parts.append(f"row {int(float(excel_row))}")
        except (TypeError, ValueError):
            pass
    return ", ".join(parts) if parts else "unidentified posting"


def _single_posting_value(group, column, gp, servicer_name, issues):
    """The one distinct non-blank value a loan's postings carry for a column
    that consolidate_cycle_postings does not itself consolidate.

    CLIENT-CONFIRMED (2026-08-11): when a field requiring a SINGLE value holds
    two or more distinct valid source values and NO documented business rule
    establishes FIRST, LAST, SUM or any other resolution, this does NOT pick
    one. It returns None -- leaving only that one output cell blank -- and
    raises a validation issue naming the loan, the field, every conflicting
    value and the source period each came from, so the reviewer resolves it.

    Silently taking the first value would be a choice dressed up as a warning.
    Generation is never blocked: every other field for the same loan populates
    normally, and both reports still generate.

    The rule is generic, not loan-specific. Measured across all 1,774 eligible
    June loans (88 of them multi-posting) against all ten of these fields, the
    only conflict in the entire population is Millennium 2024 GP 5446319102's
    Net Interest Rate: a VAR-rate loan with a quarterly reset, remitting two
    contiguous accrual periods in one cycle at 7.75% (2025-11-30 to 2026-01-01)
    and 7.25% (2026-01-01 to 2026-02-12). Both rates are correct for their own
    window; neither is "the" rate for the cycle, and consolidate_cycle_postings
    documents no authority for this field."""
    if column not in group.columns:
        return None
    present = [(index, row) for index, row in group.iterrows()
               if not _blank(row.get(column))]
    if not present:
        return None

    distinct = []
    for _index, row in present:
        text = str(row.get(column)).strip()
        if text not in distinct:
            distinct.append(text)
    if len(distinct) == 1:
        return present[0][1].get(column)

    candidates = []
    for value in distinct:
        periods = [_posting_period_label(row) for _index, row in present
                   if str(row.get(column)).strip() == value]
        candidates.append({"value": value, "source_periods": periods})
    detail = "; ".join(f"{c['value']} ({'; '.join(c['source_periods'])})" for c in candidates)
    issues.append({
        "code": ISSUE_POSTINGS_DISAGREE,
        "servicer": servicer_name,
        "SBA GP Number": gp,
        "field": column,
        "values": distinct,
        "candidates": candidates,
        "resolution": "unresolved",
        "message": (
            f"{servicer_name} GP {gp}: the postings for this loan report different "
            f"values for {column} -- {detail}. No business rule establishes which is "
            f"authoritative for a single-value field, so this cell was left BLANK "
            f"rather than a value being chosen. Please enter the correct {column} for "
            f"this loan before distributing the report. Every other field for this "
            f"loan is populated normally."
        ),
    })
    return None


def _coerce_date(value):
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime()


def build_servicer_rows(servicer_inputs_df, servicer_name, verified_overrides=None,
                        excluded_gps=None, issues=None):
    """One report row per reviewed, non-suppressed loan for ONE servicer.

    servicer_name is the soup_key -- the value that lands in Lender/Servicer
    Name and that the Remittance sheet's SUMIFS matches on. Verified against
    the client's workbook: its column U reads "Lendistry", "Millennium Bank",
    "Community Bank & Trust", which are exactly the soup_keys. Deliberately NOT
    the SRF's own "► Servicer:" cell, which for CB&T reads "Community Bank &
    Trust-West Georgia" and would not match.

    verified_overrides is {(gp, field): approved_value} exactly as
    corrections.generate_reviewed_workbook returns it -- only Loan Status and
    Delinquency Status are mapped, by checks._VERIFIED_OVERRIDE_COLUMNS. A
    rejected recommendation contributes no entry, so the lender's submitted
    value stands untouched.
    """
    issues = issues if issues is not None else []
    excluded = {str(g).strip() for g in (excluded_gps or [])}

    consolidated = checks.consolidate_cycle_postings(servicer_inputs_df)
    consolidated = checks._apply_verified_overrides(consolidated, verified_overrides or {})

    raw = servicer_inputs_df.copy()
    raw["_gp"] = checks._norm_id(raw["SBA GP Number"])
    by_gp = dict(list(raw.groupby("_gp", sort=False)))

    rows = []
    for gp, consolidated_row in consolidated.iterrows():
        key = str(gp).strip()
        if key in excluded:
            continue
        group = by_gp.get(key)
        row = {c: None for c in MSR_COLUMNS}

        row["SBA GP Number"] = key
        row["Lender Loan Number"] = consolidated_row.get("lender_loan_number")
        row["Loan Status"] = consolidated_row.get("loan_status")
        row["Delinquency Status"] = consolidated_row.get("dq_status")
        row["Net Interest"] = consolidated_row.get("net_interest")
        row["Principal"] = consolidated_row.get("principal")
        row["Total Payment Received"] = consolidated_row.get("received")
        row["Beginning Actual Balance"] = consolidated_row.get("beginning_balance")
        row["Ending Actual Balance"] = consolidated_row.get("ending_balance")

        if group is not None:
            for output_column, source_column in _SOURCE_COLUMNS.items():
                value = _single_posting_value(group, source_column, key, servicer_name, issues)
                if output_column in _DATE_COLUMNS:
                    value = _coerce_date(value)
                row[output_column] = value

        for column in UNSOURCED_COLUMNS:
            row[column] = None

        row["Lender/Servicer Name"] = servicer_name
        rows.append(row)
    return rows


def _resolve_shared_field(values, servicers, issues, unavailable_code, conflict_code, label):
    """One value shared by every servicer in a deal (Note Series, Remittance
    Date). Absent or conflicting is a VALIDATION ISSUE, never a guess."""
    present = [(s, v) for s, v in zip(servicers, values) if not _blank(v)]
    if not present:
        issues.append({
            "code": unavailable_code,
            "field": label,
            "message": (
                f"{label} could not be read from any servicer file in this deal. "
                f"The Monthly Servicer Report was generated with {label} blank."
            ),
        })
        return None
    distinct = {str(v) for _s, v in present}
    if len(distinct) > 1:
        issues.append({
            "code": conflict_code,
            "field": label,
            "message": (
                f"The servicer files in this deal state different values for {label} "
                f"({', '.join(sorted(distinct))}). The value from "
                f"{present[0][0]} was used; please confirm which is correct."
            ),
        })
    return present[0][1]


def build_deal_report(deal, servicers):
    """One deal's complete report payload.

    servicers is a list of dicts, in upload order:
        {"servicer_name": soup_key, "display_label": ..., "rows": [...],
         "note_series": "2025 FBC-1" | None,
         "remittance_date": Timestamp | None}

    Collection Period is ALWAYS None. CLIENT-CONFIRMED (2026-08-11): the
    client's reference workbook states 11, while the SRF's Data Validation
    sheet states 26 for every servicer in BOTH deals -- so it is not a
    deal-level period and cannot be that field -- and Lendistry 2025 ships no
    Data Validation sheet at all. No reliable source exists, so the report is
    generated with the field blank and a validation issue raised, rather than a
    fabricated number. Recorded as a future enhancement: collect the Collection
    Period explicitly from the reviewer before generation."""
    issues = []
    names = [s["display_label"] for s in servicers]

    note_series = _resolve_shared_field(
        [s.get("note_series") for s in servicers], names, issues,
        ISSUE_NOTE_SERIES_UNAVAILABLE, ISSUE_NOTE_SERIES_CONFLICT, "Note Series")
    remittance_date = _resolve_shared_field(
        [s.get("remittance_date") for s in servicers], names, issues,
        ISSUE_REMITTANCE_DATE_UNAVAILABLE, ISSUE_REMITTANCE_DATE_CONFLICT,
        "Remittance Date")

    issues.append({
        "code": ISSUE_COLLECTION_PERIOD_UNAVAILABLE,
        "field": "Collection Period",
        "message": (
            "Collection Period could not be derived from the uploaded files and "
            "was left blank in the Monthly Servicer Report. The servicer files "
            "state a Collection Period that does not vary by SOUP deal, so it "
            "cannot be used here. Please enter the correct Collection Period in "
            "cell B4 of the Monthly Servicer Report worksheet before "
            "distributing the report."
        ),
    })

    rows = []
    for servicer in servicers:
        rows.extend(servicer["rows"])
        issues.extend(servicer.get("issues") or [])

    # Remittance rows: one per DISTINCT servicer name, in upload order. A
    # servicer split across note series contributes to only one deal, so within
    # one report each name appears once.
    remittance_servicers = []
    for servicer in servicers:
        name = servicer["servicer_name"]
        if name not in remittance_servicers:
            remittance_servicers.append(name)

    return {
        "deal": deal,
        "note_series": note_series,
        "remittance_date": remittance_date,
        "collection_period": None,
        "rows": rows,
        "remittance_servicers": remittance_servicers,
        "validation_issues": issues,
    }


def _write_metadata(ws, report):
    ws.cell(row=_TITLE_ROW, column=1, value=REPORT_TITLE).font = Font(bold=True)

    ws.cell(row=_NOTE_SERIES_ROW, column=1, value="► Note Series:")
    note_series_cell = ws.cell(row=_NOTE_SERIES_ROW, column=2, value=report.get("note_series"))
    note_series_cell.number_format = "@"

    ws.cell(row=_REMITTANCE_DATE_ROW, column=1, value="► Remittance Date:")
    remittance_date = report.get("remittance_date")
    if remittance_date is not None and not _blank(remittance_date):
        parsed = pd.to_datetime(remittance_date, errors="coerce")
        if not pd.isna(parsed):
            remittance_date = parsed.to_pydatetime()
    else:
        remittance_date = None
    date_cell = ws.cell(row=_REMITTANCE_DATE_ROW, column=2, value=remittance_date)
    date_cell.number_format = _DATE

    ws.cell(row=_COLLECTION_PERIOD_ROW, column=1, value="► Collection Period:")
    # Deliberately left as None -- see build_deal_report.
    ws.cell(row=_COLLECTION_PERIOD_ROW, column=2,
            value=report.get("collection_period")).number_format = _COUNT


def _write_subtotal_row(ws, row_count):
    """SUBTOTAL over the ACTUAL generated population.

    The supplied workbook hardcodes A9:A1513 while holding 1,436 rows, so its
    own totals span 69 empty rows. Built from row_count here so filtering
    recalculates against exactly the rows that exist. With no rows at all the
    range would be inverted, so the cells are left empty rather than written
    with a broken formula."""
    if row_count <= 0:
        return
    last = _FIRST_DATA_ROW + row_count - 1
    count_cell = ws.cell(row=_SUBTOTAL_ROW, column=1,
                         value=f"=SUBTOTAL(2,A{_FIRST_DATA_ROW}:A{last})")
    count_cell.number_format = _COUNT
    count_cell.font = Font(bold=True)
    for column_name in _SUBTOTAL_SUM_COLUMNS:
        index = MSR_COLUMNS.index(column_name) + 1
        letter = get_column_letter(index)
        cell = ws.cell(row=_SUBTOTAL_ROW, column=index,
                       value=f"=SUBTOTAL(9,{letter}{_FIRST_DATA_ROW}:{letter}{last})")
        cell.number_format = _NUMBER_FORMATS.get(column_name, _CURRENCY)
        cell.font = Font(bold=True)


def _write_loan_sheet(ws, report):
    _write_metadata(ws, report)
    rows = report["rows"]
    _write_subtotal_row(ws, len(rows))

    for index, name in enumerate(MSR_COLUMNS, start=1):
        cell = ws.cell(row=_HEADER_ROW, column=index, value=name)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    for offset, row in enumerate(rows):
        r = _FIRST_DATA_ROW + offset
        for index, name in enumerate(MSR_COLUMNS, start=1):
            value = row.get(name)
            if isinstance(value, float) and pd.isna(value):
                value = None
            cell = ws.cell(row=r, column=index, value=value)
            number_format = _NUMBER_FORMATS.get(name)
            if number_format:
                cell.number_format = number_format

    last = _FIRST_DATA_ROW + len(rows) - 1 if rows else _HEADER_ROW
    ws.auto_filter.ref = f"A{_HEADER_ROW}:{get_column_letter(len(MSR_COLUMNS))}{last}"
    # Freeze BELOW the header so the loan table scrolls under it. The supplied
    # workbook froze at A134, which is where someone had scrolled to rather
    # than a deliberate layout choice.
    ws.freeze_panes = f"A{_FIRST_DATA_ROW}"
    for index, name in enumerate(MSR_COLUMNS, start=1):
        width = _COLUMN_WIDTHS.get(name)
        if width:
            ws.column_dimensions[get_column_letter(index)].width = width


def _write_remittance_sheet(ws, report):
    """Amounts are SUMIFS FORMULAS against the loan sheet, never pasted values.

    That is what keeps the two worksheets reconciled: filtering or editing the
    loan sheet moves the remittance figures with it, exactly as the client's
    own workbook behaves."""
    ws.cell(row=1, column=1, value="► Note Series:")
    ws.cell(row=1, column=2, value=report.get("note_series")).number_format = "@"
    ws.cell(row=2, column=1, value="► Remittance Date:")
    ws.cell(row=2, column=2,
            value=f"='{MONTHLY_SERVICER_REPORT_SHEET}'!B{_REMITTANCE_DATE_ROW}"
            ).number_format = _DATE
    ws.cell(row=3, column=1, value="► Collection Period:")
    ws.cell(row=3, column=2,
            value=f"='{MONTHLY_SERVICER_REPORT_SHEET}'!B{_COLLECTION_PERIOD_ROW}"
            ).number_format = _COUNT

    header_row = 5
    for index, name in enumerate(REMITTANCE_COLUMNS, start=1):
        cell = ws.cell(row=header_row, column=index, value=name)
        cell.font = Font(bold=True)

    payment_column = get_column_letter(MSR_COLUMNS.index("Total Payment Received") + 1)
    servicer_column = get_column_letter(MSR_COLUMNS.index("Lender/Servicer Name") + 1)

    first = header_row + 1
    for offset, servicer_name in enumerate(report["remittance_servicers"]):
        r = first + offset
        ws.cell(row=r, column=1, value=servicer_name)
        ws.cell(row=r, column=2, value=(
            f"=SUMIFS('{MONTHLY_SERVICER_REPORT_SHEET}'!{payment_column}:{payment_column},"
            f"'{MONTHLY_SERVICER_REPORT_SHEET}'!{servicer_column}:{servicer_column},"
            f"'{REMITTANCE_AMOUNT_SHEET}'!A{r})"
        )).number_format = _CURRENCY
        # Note Account has no source -- left blank. B-C still evaluates.
        ws.cell(row=r, column=3, value=None).number_format = _CURRENCY
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = _CURRENCY

    last = first + len(report["remittance_servicers"]) - 1
    total_row = last + 1
    if report["remittance_servicers"]:
        # The supplied workbook sums B4:B8 while its data starts at B6, so its
        # total spans the header row and a blank. Bounded to the real rows here.
        for index, letter in ((2, "B"), (3, "C"), (4, "D")):
            cell = ws.cell(row=total_row, column=index,
                           value=f"=SUM({letter}{first}:{letter}{last})")
            cell.number_format = _CURRENCY
            cell.font = Font(bold=True)

    ws.column_dimensions["A"].width = 30
    for letter in ("B", "C", "D"):
        ws.column_dimensions[letter].width = 18


def generate_monthly_servicer_report_bytes(report, buffer):
    """The two-worksheet workbook, in the client's sheet order."""
    wb = Workbook()
    loan_sheet = wb.active
    loan_sheet.title = MONTHLY_SERVICER_REPORT_SHEET
    _write_loan_sheet(loan_sheet, report)

    remittance_sheet = wb.create_sheet(REMITTANCE_AMOUNT_SHEET)
    _write_remittance_sheet(remittance_sheet, report)

    wb.save(buffer)
    buffer.seek(0)


def monthly_servicer_report_filename(deal, note_series, cycle_month):
    """e.g. "Monthly_Servicer_Report_2025_FBC-1_June_2026.xlsx".

    Falls back to the deal identifier when the note series is unavailable, and
    omits the period when the cycle month is unknown -- never a guessed date."""
    label = (note_series or deal or "Report").replace(" ", "_").replace("/", "-")
    if cycle_month is not None:
        try:
            return f"Monthly_Servicer_Report_{label}_{cycle_month.strftime('%B_%Y')}.xlsx"
        except AttributeError:
            pass
    return f"Monthly_Servicer_Report_{label}.xlsx"
