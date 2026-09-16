"""Formula / cache preflight — runs BEFORE Calamine ingestion.

WHY THIS EXISTS

    file_readers._load_workbook uses CalamineWorkbook, which exposes cached
    VALUES only. It has no formula API, so nothing downstream can tell a
    computed cell from a hand-typed literal:

        =COUNTIFS($A$10:A12,A12)  ->  calamine reads  1
        1  (pasted over the formula)  ->  calamine reads  1

    Two real failure modes follow, both observed on live files:

    1. OVERWRITTEN FORMULAS. The June 2026 Peoples Bank submission had 102
       cells across 5 rows replaced with pasted literals, freezing Loan
       Payment Instance at 1 and putting "Y" on three postings of one loan.
       Every downstream check saw plausible numbers.

    2. MISSING CACHED VALUES. A workbook written by openpyxl carries its
       formulas but NO cached values — openpyxl does not evaluate. Calamine
       then reads blanks. Measured on a corrections.py-style output: 57 of
       62 rows lost Loan Payment Instance and Last Payment Instance?
       entirely, while the file still opened cleanly.

    Neither is detectable once the DataFrame exists. Both are cheap to catch
    here, on the raw bytes, before Calamine sees them.

SCOPE

    This module only reports. It raises nothing, changes nothing, and knows
    nothing about fail-vs-warn policy — the caller decides. It does not
    validate balances, statuses, populations or any business rule; those
    belong in checks.py and are untouched.

COST
    ~0.15s on a 2.6MB workbook, read_only on both passes.
"""

import io

import openpyxl

# Columns that must be formula-derived, resolved BY HEADER NAME rather than
# by letter — the vF template shifts every column left, and name resolution
# handles that the same way the rest of file_readers already does.
#
# Names duplicated within a template are deliberately excluded: "Net Interest"
# and "Reporting Errors" each appear twice in the submitted template (once as
# a servicer input, once as a derived column), so a name lookup cannot tell
# them apart and would produce false findings.
PROTECTED_BY_NAME = [
    "Loan Payment Instance",
    "Last Payment Instance?",
    "Total Payments Received",
    "Monthly Payments Made?",
    "Last Month DQ Status",
    "New DQ Calc",
    "New DQ Status",
    "Status Match",
    "Error #",
    "Gross Interest Rate",
    "Gross Interest",
    "Servicing Fee",
    "Interest Difference",
]

# Present in every template variant seen. If none of these resolves, the
# preflight cannot verify anything and says so rather than passing silently.
CORE_PROTECTED = ["Loan Payment Instance", "Last Payment Instance?"]

SHEET_FILTER = "servicer input"
MAX_HEADER_SCAN_ROWS = 15
MAX_REPORTED_CELLS = 25


def _norm(s):
    return " ".join(str(s).replace("\xa0", " ").split()).lower() if s is not None else ""


def _is_formula(v):
    # read_only mode returns formulas as strings; ArrayFormula objects carry
    # .text. Anything else is a literal.
    if hasattr(v, "text"):
        return True
    return isinstance(v, str) and v.startswith("=")


def _find_sheet_and_header(ws):
    """Returns (header_row_idx, {normalised name: col_idx}) or (None, {})."""
    best = (-1, None, {})
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=MAX_HEADER_SCAN_ROWS,
                                         values_only=True), start=1):
        names = {}
        dupes = set()
        for c, v in enumerate(row, start=1):
            n = _norm(v)
            if not n:
                continue
            if n in names:
                dupes.add(n)
            names[n] = c
        for d in dupes:            # ambiguous -- cannot resolve by name
            names.pop(d, None)
        hits = sum(1 for p in PROTECTED_BY_NAME if _norm(p) in names)
        if hits > best[0]:
            best = (hits, i, names)
    return (best[1], best[2]) if best[0] > 0 else (None, {})


def preflight_workbook(file_bytes, filename=None):
    """Inspect the raw bytes for formula integrity and cached values.

    Returns a dict:
        ok                bool   -- no blocking findings
        sheet             str    -- sheet inspected
        data_rows         int
        columns_checked   list
        findings          list of {code, severity, column, rows, detail}

    Codes:
        FORM-01  a protected column contains a pasted literal
        FORM-02  a protected column is empty where a formula belongs
        FORM-03  no formula cell carries a cached value (never recalculated)
        FORM-04  the workbook could not be inspected
        FORM-05  no protected column resolved -- nothing could be verified
    """
    out = {"ok": True, "sheet": None, "data_rows": 0,
           "columns_checked": [], "findings": [], "rows_affected": []}

    def add(code, severity, column, rows, detail):
        out["findings"].append({"code": code, "severity": severity,
                                "column": column, "rows": rows,
                                "detail": detail})
        if severity == "FAIL":
            out["ok"] = False

    try:
        wbf = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True,
                                     data_only=False)
    except Exception as e:
        add("FORM-04", "FAIL", None, [],
            f"could not open the workbook for formula inspection: "
            f"{type(e).__name__}: {e}")
        return out

    try:
        names = [s for s in wbf.sheetnames if SHEET_FILTER in s.lower()] \
            or list(wbf.sheetnames)
        sheet_name = names[0]
        out["sheet"] = sheet_name
        ws = wbf[sheet_name]

        hdr_row, cols = _find_sheet_and_header(ws)
        if hdr_row is None:
            add("FORM-05", "WARN", None, [],
                f"no protected column could be resolved by name on sheet "
                f"'{sheet_name}' — formula integrity was NOT verified")
            return out

        resolved = {p: cols[_norm(p)] for p in PROTECTED_BY_NAME
                    if _norm(p) in cols}
        out["columns_checked"] = sorted(resolved)
        if not any(p in resolved for p in CORE_PROTECTED):
            add("FORM-05", "WARN", None, [],
                f"none of {CORE_PROTECTED} resolved on sheet '{sheet_name}' — "
                f"formula integrity was NOT verified")
            return out

        key_col = cols.get(_norm("SBA GP Number"), 1)
        literals, empties = {}, {}
        data_rows = 0
        for r, row in enumerate(ws.iter_rows(min_row=hdr_row + 1,
                                             values_only=True),
                                start=hdr_row + 1):
            if key_col - 1 >= len(row):
                continue
            key = row[key_col - 1]
            if key is None or (isinstance(key, str) and not key.strip()):
                continue
            data_rows += 1
            for name, c in resolved.items():
                v = row[c - 1] if c - 1 < len(row) else None
                if _is_formula(v):
                    continue
                if v is None or (isinstance(v, str) and not v.strip()):
                    empties.setdefault(name, []).append(r)
                else:
                    literals.setdefault(name, []).append((r, v))
        out["data_rows"] = data_rows
    finally:
        wbf.close()

    for name, hits in sorted(literals.items()):
        rows = [r for r, _ in hits]
        sample = "; ".join(f"row {r}={v!r}" for r, v in hits[:3])
        add("FORM-01", "FAIL", name, rows,
            f"{len(rows)} of {out['data_rows']} rows contain a pasted literal "
            f"where a formula belongs ({sample}"
            f"{', …' if len(hits) > 3 else ''}). Calamine reads these as "
            f"ordinary values, so every downstream check would treat them as "
            f"computed.")

    for name, rows in sorted(empties.items()):
        add("FORM-02", "FAIL", name, rows,
            f"{len(rows)} of {out['data_rows']} rows are empty where a "
            f"formula belongs (rows {rows[:MAX_REPORTED_CELLS]}"
            f"{', …' if len(rows) > MAX_REPORTED_CELLS else ''}).")

    # -- cached values: a workbook written programmatically has none --------
    try:
        wbv = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True,
                                     data_only=True)
    except Exception:
        return out
    try:
        ws = wbv[out["sheet"]]
        cached = 0
        checked = 0
        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            if key_col - 1 >= len(row):
                continue
            key = row[key_col - 1]
            if key is None or (isinstance(key, str) and not key.strip()):
                continue
            for c in resolved.values():
                if c - 1 < len(row):
                    checked += 1
                    if row[c - 1] is not None:
                        cached += 1
            if cached:
                break
    finally:
        wbv.close()

    affected = set()
    for f in out["findings"]:
        affected.update(f.get("rows") or [])
    out["rows_affected"] = sorted(affected)

    if checked and not cached:
        add("FORM-03", "FAIL", None, [],
            "no formula cell in any protected column carries a cached value. "
            "The workbook has been written programmatically (openpyxl does "
            "not evaluate formulas) and never recalculated in Excel. Calamine "
            "reads cached values only, so these columns would arrive blank. "
            "Open and recalculate in Excel, save, and resubmit.")

    return out


def format_preflight(result):
    """One-line-per-finding rendering, for logs or a report cell."""
    if result["ok"] and not result["findings"]:
        return (f"Preflight PASSED — {len(result['columns_checked'])} protected "
                f"column(s), {result['data_rows']} data rows on "
                f"'{result['sheet']}'.")
    lines = [f"Preflight {'PASSED with warnings' if result['ok'] else 'FAILED'} "
             f"— sheet '{result['sheet']}', {result['data_rows']} data rows."]
    if result["rows_affected"]:
        # The actionable unit is the row, not the column: one pasted row
        # trips every protected column at once.
        by_row = {}
        for f in result["findings"]:
            for r in (f.get("rows") or []):
                by_row.setdefault(r, set()).add(f["column"])
        lines.append(f"  {len(result['rows_affected'])} row(s) affected: "
                     + ", ".join(f"{r} ({len(by_row[r])} col)"
                                 for r in result["rows_affected"]))
    codes = {}
    for f in result["findings"]:
        codes.setdefault((f["code"], f["severity"]), []).append(f)
    for (code, sev), fs in sorted(codes.items()):
        cols = sorted({f["column"] for f in fs if f["column"]})
        if cols:
            lines.append(f"  {code} {sev}: {len(cols)} column(s) — "
                         + ", ".join(cols))
        for f in fs:
            if not f["column"]:
                lines.append(f"  {code} {sev}: {f['detail']}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# DISPOSITION: whole-file rejection vs row-scoped quarantine
#
# preflight_workbook itself is unchanged and still REPORTS ONLY — it raises
# nothing and knows nothing about fail-vs-warn policy. This helper answers a
# narrower question for the caller: "are this workbook's blocking findings
# confined to identifiable ROWS, or do they condemn the whole file?"
#
# Only FORM-01 is row-isolatable, and that is a property of what each code
# means, not a policy preference:
#
#   FORM-01  a specific cell on a specific row carries a pasted literal where
#            a formula belongs. The defect IS the row. Every other row in the
#            workbook is untouched and still formula-derived, so rejecting
#            them too discards good lender data to punish a bad row.
#
#   FORM-02  a protected column is empty where a formula belongs. Observed
#            live as 113-of-113 rows across four columns — a template-level
#            fault, not a row-level one.
#   FORM-03  NO formula cell anywhere carries a cached value. The workbook was
#            written programmatically and never recalculated. There is no
#            "unaffected" subset to salvage; every protected column arrives
#            blank. This is the corrected-workbook re-ingestion control.
#   FORM-04  the workbook could not be opened at all.
#   FORM-05  nothing could be verified (WARN, never blocking).
#
# So this returns rows ONLY when EVERY blocking finding is a FORM-01 that
# names at least one row. A single FORM-02/03/04 alongside them, or a FORM-01
# with an empty row list, collapses the answer back to None and the caller
# rejects the whole file exactly as before. Fail-closed by construction: the
# default answer is None.
ROW_ISOLATABLE_CODES = ("FORM-01",)


def isolatable_rows(result):
    """The 1-based worksheet rows a FORM-01-only failure is confined to.

    Returns a sorted list of row numbers when every blocking (FAIL) finding is
    a row-isolatable code carrying at least one row, otherwise None.

    None means "this failure is not row-scopable — reject the workbook", and
    it is also what a clean workbook returns (nothing to quarantine). The
    caller must treat None as the pre-existing whole-file disposition; this
    helper never decides to accept anything on its own."""
    if result.get("ok"):
        return None
    blocking = [f for f in result.get("findings", []) if f.get("severity") == "FAIL"]
    if not blocking:
        return None
    rows = set()
    for finding in blocking:
        if finding.get("code") not in ROW_ISOLATABLE_CODES:
            return None
        finding_rows = finding.get("rows") or []
        if not finding_rows:
            # A blocking finding that names no row cannot be quarantined to
            # one -- it condemns the file.
            return None
        rows.update(finding_rows)
    return sorted(rows) or None
