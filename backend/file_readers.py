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
import re
import pandas as pd
from python_calamine import CalamineWorkbook


MAX_HEADER_SCAN_ROWS = 15

SERVICE_MEMO_FIELDS = [
    "Report Date", "Lender", "SBA Loan ID", "Lender Loan ID", "Borrower Name",
    "Servicer Action", "Deferral Start Date", "Deferral End Date",
    "Regular Payment Resume Date",
]

# Deferral date column -> the column read_service_memo stores that cell's
# ORIGINAL text in, before pd.to_datetime(errors="coerce") flattens
# unparseable values to NaT. Diagnostics only (memo_audit.py).
MEMO_RAW_DATE_COLUMNS = {
    "Deferral Start Date": "_raw_deferral_start",
    "Deferral End Date": "_raw_deferral_end",
    "Regular Payment Resume Date": "_raw_deferral_resume",
}

# Centralized, explicit map from a servicer's canonical identity (the same
# soup_key value SERVICER_SOUP_KEYS/filter_soup_by_servicer already use — a
# note-series year like "2025"/"2024" is stripped out BEFORE this lookup, it
# never appears as a key or value here) to every known exact spelling that
# servicer's memo entries use in the "Lender" column.
#
# This replaces fuzzy substring matching on display labels, which silently
# failed for two known cases:
#   - "Millennium 2025"/"Millennium 2024" never matched "Millennium Bank" or
#     "Millennium Bank, IL" (the display label doesn't contain the word
#     "Bank" at all, so no substring relationship exists in either
#     direction) -- ALL Millennium deferrals were invisible to STAT-01.
#   - "Lendistry 2025"/"Lendistry 2024" matched plain "Lendistry" rows fine,
#     but missed the ~28% of Lendistry memo rows filed under "Lendistry
#     SBLC, LLC, CA" -- any deferral recorded under that entity name was
#     invisible to STAT-01.
#
# Update this table (not the matching logic in checks._lookup_deferral) when
# a new memo Lender spelling variant is discovered for a servicer already
# listed here. A servicer with NO entry here at all falls back to the old
# fuzzy heuristic rather than failing outright, so a brand-new servicer
# isn't silently broken until someone adds it -- see
# checks._lookup_deferral's docstring for that fallback behavior.
MEMO_LENDER_ALIASES = {
    # NB: the "&" spelling is the canonical soup_key -- the frontend's
    # SERVICER_SOUP_KEYS maps its "Community Bank and Trust" display label TO
    # this exact string, so this key must stay spelled with "&" to be found
    # by MEMO_LENDER_ALIASES.get(servicer_key) in checks._lookup_deferral.
    "Community Bank & Trust": ["Community Bank & Trust"],
    "Millennium Bank": ["Millennium Bank", "Millennium Bank, IL"],
    "Lendistry": ["Lendistry", "Lendistry SBLC, LLC, CA"],
    "Mission Valley Bank": ["Mission Valley Bank", "Mission Valley Bank, CA"],
    "Peoples Bank": ["Peoples Bank", "Peoples Bank, MS"],
    "Poppy Bank": ["Poppy Bank"],
    # Added 2026-08-04, PLACEHOLDER single-alias entries only -- no real
    # Servicer Memo Summary data has been seen yet for these 4 banks. Given
    # this exact gap silently broke ALL Millennium Bank deferral matching
    # until a real memo file surfaced the actual spelling variant, treat
    # these as provisional: the first real memo file received for each of
    # these banks should be checked for its actual "Lender" column
    # spelling, and this list expanded to match if it differs from the
    # plain name below.
    "Shoreham": ["Shoreham"],
    "Cogent": ["Cogent"],
    "FinWise": ["FinWise"],
    "X-Caliber": ["X-Caliber"],
}


def _build_alias_index(alias_table):
    """Reverse index of MEMO_LENDER_ALIASES: {normalized memo Lender spelling
    -> canonical soup_key}. Lets a caller answer "is this Lender string known
    to ANY servicer?" in one lookup, which is what makes an unrecognized memo
    Lender a reportable finding instead of a silent zero-match.

    A spelling claimed by two different soup_keys is a table defect (it would
    make memo attribution ambiguous). This does NOT raise on one -- an import-
    time exception would take the whole app down over a table typo -- it keeps
    the first binding and leaves detection to the alias-table drift test in
    test_memo_audit.py, which fails loudly at test time instead."""
    index = {}
    for soup_key, spellings in alias_table.items():
        for spelling in spellings:
            index.setdefault(str(spelling).strip().lower(), soup_key)
    return index


_ALIAS_TO_SOUP_KEY = _build_alias_index(MEMO_LENDER_ALIASES)


def memo_lender_soup_key(lender_value):
    """Returns the canonical soup_key for a memo "Lender" cell, or None if
    that exact spelling isn't catalogued in MEMO_LENDER_ALIASES. None means
    "no servicer in the alias table claims this spelling" -- the condition
    that silently yields zero SOUP matches today."""
    if lender_value is None:
        return None
    s = str(lender_value).strip().lower()
    if not s:
        return None
    return _ALIAS_TO_SOUP_KEY.get(s)


# Servicer Action free text that identifies a PRINCIPAL-ONLY (or interest-only)
# deferral, which follows normal payment-vs-scheduled delinquency math (just
# against a reduced, deferral-adjusted amount) rather than the full P&I
# freeze/auto-advance rule. Anything not matching this is treated as a full
# P&I deferral. Order matters: "principal and interest" / "P&I" / "PNI" must
# NOT be misclassified as partial just because they also contain "principal".
# Checked FIRST, and it has to be: text naming BOTH principal and interest is
# a full P&I deferral even though it also contains the word "principal", which
# several of the partial patterns below would otherwise match.
#
# Long phrases match as plain substrings. The short forms are word-bounded
# instead, because "pni" / "p and i" as bare substrings can appear inside
# unrelated words.
_FULL_PNI_OVERRIDE_KEYWORDS = ("principal and interest", "principal & interest")
_FULL_PNI_TOKEN_RE = re.compile(r"\b(?:p\s*&\s*i|p\s+and\s+i|pni)\b")

# Servicer Action free text identifying a PRINCIPAL-ONLY (or interest-only)
# deferral. Widened 2026-08-07 after measuring the delivered Servicer Memo
# Summary: the previous list held only the literal strings "principal only"
# and "interest only", so every OTHER way of writing a principal-only
# deferral fell through to the "full" default below. That misclassified 11 of
# 68 Deferred rows -- among them
#
#   "3-month deferment of principal paying interest"   (interest IS paid)
#   "2-month deferment with $7200 interest payments"   (interest IS paid)
#   "3-month principal deferment" / "6M principal deferral"
#
# all of which are partial deferrals being read as a full P&I freeze. That is
# the dangerous direction: "full" is what makes STAT-01 advance the
# delinquency rank +1/cycle regardless of payment, and what gates
# corrections.py's automatic Loan Status -> DEFERRED write.
#
# "paying interest" / "interest payment" are included on the same logic: if
# the servicer says interest is still being paid, principal alone is deferred.
# They cannot swallow a genuine full deferral, because any text naming both
# principal and interest is caught by the override above first.
_PARTIAL_DEFERRAL_KEYWORDS = (
    "principal only", "interest only",
    "principal deferment", "principal deferral", "principal payment deferral",
    "deferment of principal", "deferral of principal", "deferring principal",
    "paying interest", "interest payment",
)


def classify_deferral_type(servicer_action):
    """Classifies the memo's free-text Servicer Action as a full P&I or a
    principal-only/interest-only deferral.

    Evaluation order is load-bearing and must stay exactly:

        1. explicit full P&I  -> "full"
        2. partial keywords   -> "partial"
        3. neither            -> the default below

    Step 1 has to come first: text like "deferment of principal and interest"
    contains a principal-only phrase from step 2, and reading it as partial
    would be wrong. Anything naming both principal and interest is full.

    The step-3 default is "full", which is the direction that makes STAT-01
    advance the delinquency rank +1/cycle regardless of payment and that gates
    corrections.py's automatic Loan Status -> DEFERRED write. Measured against
    the delivered memo, 9 of 68 Deferred rows reach step 3 on text carrying no
    payment-type signal at all ("90-Day Deferment", "3 month payment
    deferral") — see docs/deferralevidencebrief.md."""
    s = str(servicer_action).strip().lower()
    if any(k in s for k in _FULL_PNI_OVERRIDE_KEYWORDS) or _FULL_PNI_TOKEN_RE.search(s):
        return "full"
    if any(k in s for k in _PARTIAL_DEFERRAL_KEYWORDS):
        return "partial"
    return "full"

# Fixed set of servicers this client works with. Used to identify the
# servicer from the uploaded filename, which is far more reliable than
# scanning for an in-sheet "Servicer:" label — that label pattern only
# exists in some servicers' templates (e.g. Poppy Bank, Millennium Bank).
# Lendistry's "Last Month Final" template has no such label anywhere in the
# file, so relying on the label alone silently falls through to
# "Unknown Servicer" for any servicer using this template.
KNOWN_SERVICERS = [
    "Poppy Bank", "Millennium Bank", "Mission Valley Bank",
    "Peoples Bank", "Lendistry", "Community Bank and Trust",
    # Added 2026-08-04. This list is only consulted as a fallback for
    # filenames that don't follow the frontend's current
    # "{soupKey}||{displayLabel}__{filename}" convention (see
    # _parse_servicer_filename) -- effectively dead code for uploads through
    # the current UI, kept updated anyway as defense-in-depth against any
    # future upload path that bypasses that convention.
    "Shoreham", "Cogent", "FinWise", "X-Caliber",
]

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


def _rows_to_df(rows, header_row_idx, sheet_name=None):
    """Builds a DataFrame from a list-of-lists sheet, starting at a given
    header row index (0-based).

    When sheet_name is provided, also attaches three extra columns used only
    by the correction engine (backend/corrections.py) to later re-locate the
    exact source cell for a value that gets corrected:
      - "_sheet_name":        the sheet this row came from (constant)
      - "_header_excel_row":  the 1-based Excel row of the header row itself
      - "_excel_row":         this row's own 1-based Excel row number

    Row numbers are computed BEFORE the blank-row filter below runs, and
    carried forward as an explicit column rather than relying on the
    DataFrame's positional index — the index gets reset to a fresh
    0..N-1 range by the filter (and by any later reset_index/groupby/
    sort_values a check might do), so it can't be trusted to mean "the
    original sheet row" by the time a check actually uses it. An explicit
    column rides along through those operations the same way "SBA GP
    Number" already does.

    Callers that don't pass sheet_name (SOUP, Loan Terms, Servicer Memo)
    get back exactly the same DataFrame as before this change — this is
    purely additive for the one caller that opts in."""
    data_section = rows[header_row_idx:]
    if not data_section:
        return pd.DataFrame()
    header_row = data_section[0]
    headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(header_row)]

    if sheet_name is not None:
        header_excel_row = header_row_idx + 1  # 1-based Excel row of the header itself
        # Enumerate BEFORE filtering blanks, so a blank row earlier in the
        # sheet can never shift a later real row's recorded position.
        numbered = list(enumerate(data_section[1:]))
        kept = [(i, r) for i, r in numbered if any(c is not None and str(c).strip() != "" for c in r)]
        data_rows = [r for _, r in kept]
        df = pd.DataFrame(data_rows, columns=headers)
        df = df.loc[:, ~df.columns.duplicated()]
        df["_sheet_name"] = sheet_name
        df["_header_excel_row"] = header_excel_row
        # i is 0-based offset into data_section[1:]; the header itself is at
        # Excel row `header_excel_row`, so the i-th data row is at
        # header_excel_row + 1 + i.
        df["_excel_row"] = [header_excel_row + 1 + i for i, _ in kept]
        return df

    data_rows = [r for r in data_section[1:] if any(c is not None and str(c).strip() != "" for c in r)]
    df = pd.DataFrame(data_rows, columns=headers)
    df = df.loc[:, ~df.columns.duplicated()]
    return df


_NOTE_SERIES_DEAL_PATTERN = re.compile(r"(\d{4}).*?-(\d+)\s*$")


def extract_deal_from_note_series(wb):
    """Reads the "► Note Series:" cell from the Servicer Inputs sheet header
    block (e.g. "2025 FBC-1") and converts it to SOUP's Deal-column
    convention (e.g. "SOUP2025_1"), which is what lets POP-01 scope its
    comparison to Servicer + Deal rather than Servicer alone -- necessary
    for any servicer split across multiple note-series SRF files, since
    SOUP pools all of a servicer's deals together under one Servicer name.
    Validated against every servicer's actual SOUP Deal values before this
    was wired in: Community Bank & Trust, Poppy Bank, Mission Valley Bank,
    Peoples Bank, Millennium Bank (both note series), and Lendistry (both
    note series) all parse to a Deal value that's actually present in SOUP
    for that servicer.

    Returns None if the cell can't be found or doesn't match the expected
    "{year} FBC-{n}" style pattern — callers should treat that as "Deal
    unavailable" and fall back to servicer-only comparison rather than
    erroring, since this format isn't guaranteed to hold for every future
    servicer template."""
    try:
        sheet = wb.get_sheet_by_name("Servicer Inputs")
    except Exception:
        return None
    rows = sheet.to_python()
    for row in rows[:8]:
        for i, cell in enumerate(row):
            if cell and "note series" in str(cell).lower() and i + 1 < len(row):
                raw = row[i + 1]
                if not raw:
                    return None
                m = _NOTE_SERIES_DEAL_PATTERN.search(str(raw).strip())
                if not m:
                    return None
                year, num = m.groups()
                return f"SOUP{year}_{num}"
    return None


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
    rows_cache.clear()  # free the raw parsed rows now that the DataFrame exists
    del wb
    if "SBA GP Number" in df.columns:
        df["SBA GP Number"] = normalize_id_series(df["SBA GP Number"])
    return df


def read_servicer_file(file_bytes, filename=None):
    """Returns (soup_key, display_label, loan_terms_df, servicer_inputs_df, deal)
    for one SRF file. soup_key is what filter_soup_by_servicer should match
    against SOUP's Servicer column; display_label is what should be shown
    in report tabs, the summary, and the chatbot — these differ for files
    like Lendistry's two yearly uploads, which both match SOUP's plain
    "Lendistry" but need distinct display labels. deal is this file's own
    note-series identity in SOUP's Deal-column convention (e.g.
    "SOUP2025_1"), from extract_deal_from_note_series() — None if it
    couldn't be determined, in which case population_check falls back to
    comparing by servicer alone rather than erroring.
    filename is used to resolve both, in order of reliability:
      1. The "{soupKey}||{displayLabel}__{originalFilename}" convention the
         frontend sends (exact, no guessing)
      2. Filename match against KNOWN_SERVICERS (older-style uploads)
      3. In-file "Servicer:" label scan (older servicer templates)
      4. "Unknown Servicer" for both, if nothing else matches
    """
    wb = _load_workbook(file_bytes)
    rows_cache = {}

    si_best = _find_best_header(wb, SERVICER_INPUTS_FIELDS, sheet_filter="servicer input", _rows_cache=rows_cache)
    if si_best["matches"] < len(SERVICER_INPUTS_FIELDS) // 2:
        si_best = _find_best_header(wb, SERVICER_INPUTS_FIELDS, _rows_cache=rows_cache)
    if si_best["matches"] < len(SERVICER_INPUTS_FIELDS) // 2:
        raise ValueError("Could not find the Servicer Inputs columns in this file.")
    servicer_inputs_df = _rows_to_df(rows_cache[si_best["sheet_name"]], si_best["row_idx"], sheet_name=si_best["sheet_name"])

    lt_best = _find_best_header(wb, LOAN_TERMS_FIELDS, sheet_filter="loan terms", _rows_cache=rows_cache)
    if lt_best["matches"] < len(LOAN_TERMS_FIELDS) // 2:
        lt_best = _find_best_header(wb, LOAN_TERMS_FIELDS, _rows_cache=rows_cache)
    loan_terms_df = _rows_to_df(rows_cache[lt_best["sheet_name"]], lt_best["row_idx"]) if lt_best["matches"] >= 2 else pd.DataFrame()

    soup_key, display_label = _parse_servicer_filename(filename)
    if not soup_key:
        fallback = _find_servicer_name_from_filename(filename)
        if not fallback:
            fallback = _find_servicer_name(rows_cache[si_best["sheet_name"]])
        soup_key, display_label = fallback, fallback

    deal = extract_deal_from_note_series(wb)

    rows_cache.clear()  # free the raw parsed rows now that the DataFrames exist
    del wb

    for df in (servicer_inputs_df, loan_terms_df):
        if "SBA GP Number" in df.columns:
            df["SBA GP Number"] = normalize_id_series(df["SBA GP Number"])
            # drop fully-blank trailing rows that slip past the blank-row filter
            # when only the ID column has a stray value
            df.dropna(subset=["SBA GP Number"], inplace=True)

    return soup_key, display_label, loan_terms_df, servicer_inputs_df, deal


def _parse_servicer_filename(filename):
    """Parses the "{soupKey}||{displayLabel}__{originalFilename}" convention
    the frontend uses to unambiguously tell the backend both (a) which SOUP
    servicer name to filter against, and (b) what display label to use for
    this file's report tabs/summary — these differ for cases like Lendistry's
    two yearly files, which both must match SOUP's "Lendistry" but need
    distinct tab labels ("Lendistry 2024" / "Lendistry 2025").

    Returns (soup_key, display_label) or (None, None) if the filename doesn't
    follow this convention (e.g. an older-style upload) — callers should fall
    back to KNOWN_SERVICERS matching or the in-file label scan in that case."""
    if not filename or "__" not in filename:
        return None, None
    prefix = filename.split("__", 1)[0].strip()
    if not prefix:
        return None, None
    if "||" in prefix:
        soup_key, display_label = prefix.split("||", 1)
        soup_key, display_label = soup_key.strip(), display_label.strip()
        return (soup_key or None), (display_label or soup_key or None)
    # no "||" separator — same value serves as both soup key and display label
    return prefix, prefix


def _find_servicer_name_from_filename(filename):
    """Matches the uploaded filename against the known servicer list
    (case-insensitive substring match, ignoring spaces/underscores/hyphens),
    for filenames that don't follow the "{soupKey}||{displayLabel}__..."
    convention. Returns the matched servicer name, or None if nothing matches."""
    if not filename:
        return None
    normed_filename = _norm(filename)
    for servicer in KNOWN_SERVICERS:
        if _norm(servicer) in normed_filename:
            return servicer
    return None


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


def normalize_memo_loan_id(value):
    """Normalizes a memo loan-ID field (SBA Loan ID or Lender Loan ID) to a
    plain digit string so it can be matched against SBA GP Number.

    The memo file is hand-maintained and mixes formats for what is otherwise
    the same 10-digit SBA GP Number — e.g. '65360291-07', '6536029107.0',
    and '6536029107' all represent the same loan. Since the SBA GP Number is
    defined (per the SRF Instructions tab) as containing only digits, any
    non-digit character in these fields is pure formatting and safe to
    strip — this can't accidentally merge two distinct loan numbers.

    Numeric-looking values are routed through numeric conversion first so a
    trailing '.0' from a float doesn't get read as a literal digit (naively
    stripping non-digits from '6536029107.0' would wrongly produce
    '65360291070', an extra trailing zero).

    Not every memo row's Lender Loan ID is an SBA GP Number in disguise —
    some servicers (e.g. Lendistry) use short internal reference numbers
    there instead (e.g. '1497240'). Callers should try matching a loan
    against BOTH the normalized SBA Loan ID and Lender Loan ID columns,
    rather than assuming either one alone is authoritative.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s:
        return None
    numeric = pd.to_numeric(s, errors="coerce")
    if pd.notna(numeric):
        return str(int(numeric))
    digits = re.sub(r"\D", "", s)
    return digits or None


def read_service_memo(file_bytes):
    """Returns the Servicer Memo Summary as a DataFrame: one row per
    deferral/modification memo entry, across all servicers and all report
    dates. A single loan can have multiple rows here (repeat deferrals, or
    the same deferral re-reported on a later Report Date) — callers matching
    a loan + reporting cycle should pick the row whose deferral window
    actually covers that cycle, not just the loan's most recent entry.

    Adds two derived columns:
      - "gp_sba":    normalized digit-only SBA Loan ID
      - "gp_lender": normalized digit-only Lender Loan ID
      - "deferral_type": "full" or "partial", from classify_deferral_type()

    Also carries the source row's own position ("_sheet_name",
    "_header_excel_row", "_excel_row", exactly as read_servicer_file already
    does) and, for each of the three deferral date columns, the cell's
    ORIGINAL text under the matching MEMO_RAW_DATE_COLUMNS name. Both are for
    diagnostics: pd.to_datetime(errors="coerce") turns a literal "n/a" and a
    genuinely empty cell into the same NaT, so without the raw text there is
    no way to tell "the servicer asserted a deferral but wrote no usable
    date" from "no date was claimed" -- and no way to point a reviewer at the
    row to fix. Nothing in the reconciliation path reads either; see
    memo_audit.py.
    """
    wb = _load_workbook(file_bytes)
    rows_cache = {}
    best = _find_best_header(wb, SERVICE_MEMO_FIELDS, _rows_cache=rows_cache)
    if best["matches"] < len(SERVICE_MEMO_FIELDS) // 2:
        raise ValueError("Could not find the expected Servicer Memo Summary columns in this file.")
    df = _rows_to_df(rows_cache[best["sheet_name"]], best["row_idx"], sheet_name=best["sheet_name"])
    rows_cache.clear()
    del wb

    if "SBA Loan ID" in df.columns:
        df["gp_sba"] = df["SBA Loan ID"].apply(normalize_memo_loan_id)
    if "Lender Loan ID" in df.columns:
        df["gp_lender"] = df["Lender Loan ID"].apply(normalize_memo_loan_id)
    if "Servicer Action" in df.columns:
        df["deferral_type"] = df["Servicer Action"].apply(classify_deferral_type)

    # Snapshot the raw text BEFORE the coercion below overwrites it.
    for col, raw_col in MEMO_RAW_DATE_COLUMNS.items():
        if col in df.columns:
            df[raw_col] = df[col].map(lambda v: "" if v is None else str(v).strip())

    for col in ("Deferral Start Date", "Deferral End Date", "Regular Payment Resume Date", "Report Date"):
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    return df
