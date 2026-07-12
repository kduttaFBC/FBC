"""
FalconBridge Audit Intelligence — Excel report generation with openpyxl.

Builds a multi-tab workbook from the SBA reconciliation results produced by
checks.py: one Summary tab, then one detail tab per check per servicer.
"""

from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


NAVY = "0C447C"
WHITE = "FFFFFF"
ROW_ALT = "EBF4FF"
GREEN = "0F6E56"
GREEN_LIGHT = "E1F5EE"
AMBER = "BA7517"
AMBER_LIGHT = "FAEEDA"
RED = "A32D2D"
RED_LIGHT = "FCEBEB"
BLUE = "185FA5"
BLUE_LIGHT = "E6F1FB"

HEADER_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FONT = Font(color=WHITE, bold=True, size=11)
ALT_FILL = PatternFill("solid", fgColor=ROW_ALT)
WHITE_FILL = PatternFill("solid", fgColor=WHITE)

PASS_FILL = PatternFill("solid", fgColor=GREEN_LIGHT)
FAIL_FILL = PatternFill("solid", fgColor=RED_LIGHT)
WARN_FILL = PatternFill("solid", fgColor=AMBER_LIGHT)
INFO_FILL = PatternFill("solid", fgColor=BLUE_LIGHT)

PASS_FONT = Font(color=GREEN, bold=True)
FAIL_FONT = Font(color=RED, bold=True)
WARN_FONT = Font(color=AMBER, bold=True)
INFO_FONT = Font(color=BLUE, bold=True)

THIN = Side(style="thin", color="E2E8F0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _status_fill_font(status):
    if status == "PASS":
        return PASS_FILL, PASS_FONT
    if status == "FAIL":
        return FAIL_FILL, FAIL_FONT
    if status == "INFO":
        return INFO_FILL, INFO_FONT
    return WARN_FILL, WARN_FONT


def _tab_color(status):
    return {"PASS": GREEN, "FAIL": RED, "WARN": AMBER, "INFO": BLUE}.get(status, NAVY)


def _autofit(ws, headers, rows):
    for col_idx, header in enumerate(headers, start=1):
        max_len = len(str(header))
        for row in rows:
            value = row[col_idx - 1] if col_idx - 1 < len(row) else ""
            max_len = max(max_len, len(str(value)))
        width = max(12, min(45, max_len + 2))
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def _write_header(ws, headers):
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", horizontal="left")
        cell.border = BORDER
    ws.freeze_panes = "A2"


_CURRENCY_KEYWORDS = ("balance", "payment", "amount", "diff", "received", "interest", "loss")


def _write_rows(ws, headers, rows, status_col=None):
    currency_cols = {
        i for i, h in enumerate(headers, start=1)
        if any(k in str(h).lower() for k in _CURRENCY_KEYWORDS)
    }
    for r_offset, row in enumerate(rows):
        excel_row = r_offset + 2
        fill = ALT_FILL if r_offset % 2 else WHITE_FILL
        for col_idx in range(1, len(headers) + 1):
            value = row[col_idx - 1] if col_idx - 1 < len(row) else ""
            cell = ws.cell(row=excel_row, column=col_idx, value=value)
            cell.fill = fill
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", horizontal="left")
            if col_idx in currency_cols and isinstance(value, (int, float)) and not isinstance(value, bool):
                cell.number_format = '"$"#,##0.00'
                cell.alignment = Alignment(vertical="center", horizontal="right")
            if status_col is not None and col_idx == status_col:
                s_fill, s_font = _status_fill_font(str(value))
                cell.fill = s_fill
                cell.font = s_font


_INVALID_SHEET_CHARS = set('[]:*?/\\')


def _sheet_name_for(result):
    """Excel sheet names cap at 31 chars, can't repeat, and can't contain
    []:*?/\\ — combine check name + servicer, sanitized and truncated."""
    base = f"{result['name']} - {result['servicer']}"
    base = "".join(c for c in base if c not in _INVALID_SHEET_CHARS)
    return base[:31]


CHECK_LEGEND = [
    ("Population Check", "FAIL",
     "Confirms every loan in the servicer's active population (Loan Terms & Prior Period) "
     "also appears in SOUP's history for that servicer, and vice versa. A loan in SOUP's "
     "history but missing from the active list is often just a paid-off or liquidated loan — "
     "but a large or unexplained gap is worth a closer look."),
    ("Balance Continuity", "FAIL / INFO",
     "Confirms this cycle's reported Beginning Balance ties to SOUP's most recently recorded "
     "Ending Balance for that loan. Loans with more than one payment posted this cycle are "
     "called out separately (informational) rather than treated as a real break."),
    ("Status / Delinquency Match", "WARN",
     "Confirms the servicer's reported Loan Status and Delinquency Status this cycle match "
     "SOUP's last recorded status for that loan. Status changes are often legitimate "
     "month-over-month movement, but every change is surfaced for review."),
    ("Missing Payment Periods", "FAIL",
     "Reviews each active loan's full monthly reporting history in SOUP and flags any skipped "
     "month, or a loan that has stopped appearing in SOUP altogether even though it's still "
     "listed as active. Often the most actionable finding — it surfaces loans that have gone "
     "dark in the servicing data."),
    ("Payment Shortfall Check", "FAIL",
     "Compares the amount actually remitted (payment received plus the servicer's fee, which "
     "is normally netted out before remittance) against the scheduled payment for that loan. "
     "Flags a genuine shortfall or overage beyond a small tolerance. This check depends on "
     "Scheduled Payment and Servicing Fee columns being present in the servicer's file — it can "
     "be removed if a particular servicer's template doesn't report those fields."),
    ("Duplicate Postings", "INFO",
     "Informational only. Flags loans with more than one payment posted in the same cycle "
     "(e.g. a catch-up payment) so they aren't mistaken for a data error elsewhere in the report."),
]


def _build_legend(ws, start_row):
    ws.cell(row=start_row, column=1, value="What These Checks Mean").font = Font(bold=True, size=12)
    r = start_row + 1
    headers = ["Check", "Severity", "What it checks"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=r, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
    r += 1
    for name, severity, desc in CHECK_LEGEND:
        fill = ALT_FILL if (r - start_row) % 2 else WHITE_FILL
        c1 = ws.cell(row=r, column=1, value=name)
        c2 = ws.cell(row=r, column=2, value=severity)
        c3 = ws.cell(row=r, column=3, value=desc)
        for c in (c1, c2, c3):
            c.fill = fill
            c.border = BORDER
            c.alignment = Alignment(vertical="top", horizontal="left", wrap_text=True)
        ws.row_dimensions[r].height = 48
        r += 1
    return r + 1  # next free row


def _build_summary(ws, results, summary):
    ws.cell(row=1, column=1, value="FalconBridge SBA Reconciliation Summary").font = Font(bold=True, size=13)
    ws.cell(row=2, column=1, value=f"Servicers processed: {', '.join(summary.get('servicers_processed', []))}")
    ws.cell(row=3, column=1, value=f"SOUP records: {summary.get('soup_records', 0)}  |  "
                                    f"Servicer records: {summary.get('servicer_records', 0)}  |  "
                                    f"Total issues: {summary.get('total_issues', 0)}")

    header_row = _build_legend(ws, start_row=5)
    header_row += 1  # blank row before the results table

    headers = ["Servicer", "Check ID", "Check Name", "Category", "Status", "Issues Found"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
    ws.freeze_panes = f"A{header_row + 1}"

    rows = []
    for r in results:
        base_id = r["id"].split("::")[0]
        rows.append([r["servicer"], base_id, r["name"], r["category"], r["status"], r["issues_found"]])

    for r_offset, row in enumerate(rows):
        excel_row = header_row + 1 + r_offset
        fill = ALT_FILL if r_offset % 2 else WHITE_FILL
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=excel_row, column=col_idx, value=value)
            cell.fill = fill
            cell.border = BORDER
            if col_idx == 5:
                s_fill, s_font = _status_fill_font(str(value))
                cell.fill = s_fill
                cell.font = s_font

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(str(header))
        for row in rows:
            v = row[col_idx - 1]
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(col_idx)].width = max(14, min(40, max_len + 2))

    ws.column_dimensions["C"].width = max(ws.column_dimensions["C"].width or 0, 70)  # room for legend descriptions

    ws.sheet_properties.tabColor = NAVY


def _build_detail_sheet(ws, result):
    records = result["records"]
    if records:
        headers = list(records[0].keys())
        rows = [[rec.get(h, "") for h in headers] for rec in records]
    else:
        headers = ["Result"]
        rows = [["No issues found."]]
    _write_header(ws, headers)
    _write_rows(ws, headers, rows)
    _autofit(ws, headers, rows)
    ws.sheet_properties.tabColor = _tab_color(result["status"])


def generate_report_bytes(results, summary, buffer):
    wb = Workbook()

    summary_ws = wb.active
    summary_ws.title = "Summary"
    _build_summary(summary_ws, results, summary)

    used_names = {"Summary"}
    for result in results:
        name = _sheet_name_for(result)
        # guard against duplicate truncated names
        suffix = 2
        original = name
        while name in used_names:
            name = f"{original[:28]}-{suffix}"
            suffix += 1
        used_names.add(name)

        ws = wb.create_sheet(name)
        _build_detail_sheet(ws, result)

    wb.save(buffer)
    buffer.seek(0)


def report_filename():
    return f"FalconBridge_SBA_Reconciliation_{date.today().isoformat()}.xlsx"
