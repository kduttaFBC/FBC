"""
FalconBridge Audit Intelligence — smart file readers for SOUP and servicer
SRF files.

Real files bury the header row a few rows down (title rows, summary rows
above it) and the data can be on any sheet, so rather than assuming row 1 /
sheet 1, these readers scan for the row that best matches a set of expected
field names — the same approach used in the frontend's UploadScreen.jsx, so
what validates client-side matches what the backend actually reads.

Uses python-calamine (a Rust-based Excel parser) instead of openpyxl —
roughly 9x faster on large files like the ~25,000-row SOUP dataset, which
matters a lot on a CPU-limited hosting tier.
"""

import io
import pandas as pd
from python_calamine import CalamineWorkbook


MAX_HEADER_SCAN_ROWS = 15

SOUP_FIELDS = [
    "SBA GP Number", "Servicer", "Payment Month", "Loan Status",
    "Delinquency Status", "Beginning Actual Balance", "Ending Actual Balance",
]
SERVICER_INPUTS_FIELDS = [
    "SBA GP Number", "Lender Loan Number", "Loan Status",
    "Delinquency Status", "Beginning Balance", "Ending Balance", "Payment Received",
]
LOAN_TERMS_FIELDS = [
    "SBA GP Number", "Lender Name", "Loan Status", "Delinquency Status", "Ending Balance",
]


def _norm(s):
    return str(s).strip().lower().replace(" ", "").replace("_", "").replace("-", "")


def _count_matches(required_fields, header_row):
    normed = {_norm(h) for h in header_row if h is not None}
    return sum(1 for f in required_fields if _norm(f) in normed)


def _load_workbook(file_bytes):
    return CalamineWorkbook.from_object(io.BytesIO(file_bytes))


def _sheet_rows(wb, sheet_name):
    """Returns all rows of a sheet as a list of lists."""
    sheet = wb.get_sheet_by_name(sheet_name)
    return sheet.to_python()


def _find_best_header(wb, required_fields, sheet_filter=None, _rows_cache=None):
    """Scans every sheet (or only sheets matching sheet_filter, a substring
    match on sheet name) for the row that best matches required_fields.
    _rows_cache, if provided, is a dict this function will populate with
    {sheet_name: rows} so callers can reuse the already-parsed rows instead
    of re-reading the sheet."""
    best = {"matches": -1, "sheet_name": None, "row_idx": None, "header_row": None}

    sheet_names = wb.sheet_names
    if sheet_filter:
        filtered = [s for s in sheet_names if sheet_filter.lower() in s.lower()]
        if filtered:
            sheet_names = filtered

    for sheet_name in sheet_names:
        rows = _sheet_rows(wb, sheet_name)
        if _rows_cache is not None:
            _rows_cache[sheet_name] = rows
        for i, row in enumerate(rows[:MAX_HEADER_SCAN_ROWS]):
            matches = _count_matches(required_fields, row)
            if matches > best["matches"]:
                best = {"matches": matches, "sheet_name": sheet_name, "row_idx": i, "header_row": row}

    return best


def _rows_to_df(rows, header_row_idx):
    """Builds a DataFrame from a list-of-lists sheet, starting at a given
    header row index (0-based)."""
    data_section = rows[header_row_idx:]
    if not data_section:
        return pd.DataFrame()
    header_row = data_section[0]
    headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(header_row)]
    data_rows = [r for r in data_section[1:] if any(c is not None and str(c).strip() != "" for c in r)]
    df = pd.DataFrame(data_rows, columns=headers)
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def normalize_id_series(series):
    """Normalizes an ID column to clean strings — SBA GP Numbers load as
    floats when any row in the column is blank, which turns '1261949109'
    into '1261949109.0' under a naive str cast. Route numeric-looking values
    through Int64 first so they compare cleanly against the same IDs in
    other files. Also treats empty strings as missing — calamine returns ''
    for blank cells rather than None, which would otherwise slip past
    dropna() and leave a phantom blank-ID row in the data."""
    as_str = series.astype("string").str.strip()
    as_str = as_str.replace("", pd.NA)
    numeric = pd.to_numeric(as_str, errors="coerce")
    numeric_mask = numeric.notna()
    as_str = as_str.mask(numeric_mask, numeric.astype("Int64").astype("string"))
    return as_str


def read_soup_file(file_bytes):
    """Returns the full SOUP performance dataset as a DataFrame."""
    wb = _load_workbook(file_bytes)
    rows_cache = {}
    best = _find_best_header(wb, SOUP_FIELDS, _rows_cache=rows_cache)
    if best["matches"] < len(SOUP_FIELDS) // 2:
        raise ValueError("Could not find the expected SOUP columns in this file.")
    df = _rows_to_df(rows_cache[best["sheet_name"]], best["row_idx"])
    if "SBA GP Number" in df.columns:
        df["SBA GP Number"] = normalize_id_series(df["SBA GP Number"])
    return df


def read_servicer_file(file_bytes):
    """Returns (servicer_name, loan_terms_df, servicer_inputs_df) for one SRF file."""
    wb = _load_workbook(file_bytes)
    rows_cache = {}

    si_best = _find_best_header(wb, SERVICER_INPUTS_FIELDS, sheet_filter="servicer input", _rows_cache=rows_cache)
    if si_best["matches"] < len(SERVICER_INPUTS_FIELDS) // 2:
        si_best = _find_best_header(wb, SERVICER_INPUTS_FIELDS, _rows_cache=rows_cache)
    if si_best["matches"] < len(SERVICER_INPUTS_FIELDS) // 2:
        raise ValueError("Could not find the Servicer Inputs columns in this file.")
    servicer_inputs_df = _rows_to_df(rows_cache[si_best["sheet_name"]], si_best["row_idx"])

    lt_best = _find_best_header(wb, LOAN_TERMS_FIELDS, sheet_filter="loan terms", _rows_cache=rows_cache)
    if lt_best["matches"] < len(LOAN_TERMS_FIELDS) // 2:
        lt_best = _find_best_header(wb, LOAN_TERMS_FIELDS, _rows_cache=rows_cache)
    loan_terms_df = _rows_to_df(rows_cache[lt_best["sheet_name"]], lt_best["row_idx"]) if lt_best["matches"] >= 2 else pd.DataFrame()

    # servicer name — look for the "► Servicer:" label near the top of
    # Servicer Inputs and take the cell to its right
    servicer_name = _find_servicer_name(rows_cache[si_best["sheet_name"]])

    for df in (servicer_inputs_df, loan_terms_df):
        if "SBA GP Number" in df.columns:
            df["SBA GP Number"] = normalize_id_series(df["SBA GP Number"])
            # drop fully-blank trailing rows that slip past the blank-row filter
            # when only the ID column has a stray value
            df.dropna(subset=["SBA GP Number"], inplace=True)

    return servicer_name, loan_terms_df, servicer_inputs_df


def _find_servicer_name(rows):
    for row in rows[:8]:
        for i, cell in enumerate(row):
            if cell and "servicer" in str(cell).lower() and ":" in str(cell):
                if i + 1 < len(row) and row[i + 1]:
                    return str(row[i + 1]).strip()
    return "Unknown Servicer"


def filter_soup_by_servicer(soup_df, servicer_name):
    if "Servicer" not in soup_df.columns:
        return soup_df.iloc[0:0]
    mask = soup_df["Servicer"].astype("string").str.strip().str.lower() == servicer_name.strip().lower()
    return soup_df[mask].copy()
