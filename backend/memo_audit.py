"""
Servicer Memo Summary coverage diagnostic.

WHAT THIS IS FOR
----------------
The memo lookup in checks._lookup_deferral runs PER LOAN, over the servicer's
own population. That direction can only ever see memo rows belonging to a loan
it is already looking at, so one whole class of defect is structurally
invisible to it: a memo row whose loan ID exists in NO population at all is
never visited by any loan iteration. It cannot warn about a row it never
reaches. (The reverse case -- ID found, Lender spelling unrecognized -- IS
already surfaced there, at checks.py:1031-1044.)

This module sweeps the memo file itself, which is the only direction that can
see those orphans. It answers "does the memo file's content actually reach the
audit, and is what it says internally consistent" and produces a report line
per problem row.

SCOPE -- deliberately a STANDALONE DIAGNOSTIC
---------------------------------------------
Nothing here is wired into STAT-01, corrections.py, or the delinquency
calculation, and nothing here decides deferral status. It classifies and
counts; a human reads the result. Feeding any of this back into the
reconciliation path (an evidence_status contract, warnings raised from the
memo sweep) is a separate, later piece of work and must not be smuggled in
here. It imports from file_readers only -- never from checks -- so the
dependency can't quietly reverse.

THE PARTITION
-------------
Every memo row gets exactly ONE row class, and the three classes are
exhaustive over the sheet:

    phantom        blank Borrower Name + both loan IDs + Servicer Action.
                   They are not defects, but any ratio computed without
                   dropping them first is wrong.

                   On the delivered file this class measures ZERO, and that
                   is correct rather than an oversight. python-calamine
                   returns Excel ERROR cells -- both a literal "#N/A" and a
                   formula evaluating to one -- as the empty string, so a row
                   whose only remaining content is a stale "#N/A" Deal lookup
                   is already all-blank when _rows_to_df applies its
                   blank-row filter, and never reaches this module. (Reading
                   the same sheet with openpyxl instead shows 431 rows and
                   233 phantom. The pipeline never uses openpyxl to read; any
                   baseline measured that way is measuring a file the audit
                   does not see.) The class is still live: a stale row that
                   carries a real Report Date or Lender DOES survive the
                   reader, and must be recognized as phantom rather than
                   counted as a content row missing its ID.
    lookup-failed  a content row whose loan ID is in no population -- the
                   orphan case above.
    joined         a content row whose loan ID is in the population.

Independently, a content row can carry any number of QUALITY FLAGS. These
OVERLAP the classes and each other, and are counted separately -- they are not
partition members (a "date-anomaly" row is also a "joined" row):

    deferral-dates-missing  Securitization Action == "Deferred" but Deferral
                            Start Date doesn't parse. Deferral asserted,
                            dates absent -- a third state, neither "has
                            evidence" nor "no evidence".
    date-anomaly            an impossible date ORDER (see below).
    label-conflict          a non-Deferred Securitization Action carrying a
                            real Deferral Start Date. The label and the dates
                            disagree; which one is authoritative is a
                            business-rule question, NOT something this module
                            decides.
    lender-unmapped         the "Lender" spelling is claimed by no servicer in
                            MEMO_LENDER_ALIASES, so no servicer_key can ever
                            match this row.

DATE-ANOMALY COUNTING (matters -- it's the regression baseline)
--------------------------------------------------------------
Counted as VIOLATIONS, not rows, because one row can break more than one
ordering rule. Exactly two rules count toward the baseline:

    end    < start
    resume < start

so a row with both contributes 2. `date_anomaly_violations` sums those;
`date_anomaly_rows` counts distinct rows with at least one. The measured
baseline on the delivered file is 5 violations across 4 rows.

`resume < end` is reported SEPARATELY as an advisory and deliberately kept OUT
of that 5. It is genuinely invalid -- checks.py:948 requires
start <= end <= resume -- but the baseline was measured from the two rules
above, and folding a third in would silently move the number the regression
guard exists to hold still.
"""

import pandas as pd

from file_readers import (
    MEMO_RAW_DATE_COLUMNS,
    memo_lender_soup_key,
    normalize_memo_loan_id,
)


ROW_CLASS_PHANTOM = "phantom"
ROW_CLASS_JOINED = "joined"
ROW_CLASS_LOOKUP_FAILED = "lookup-failed"

# Exhaustive over the sheet: every row gets exactly one of these.
MEMO_ROW_CLASSES = (ROW_CLASS_PHANTOM, ROW_CLASS_JOINED, ROW_CLASS_LOOKUP_FAILED)

FLAG_DATES_MISSING = "deferral-dates-missing"
FLAG_DATE_ANOMALY = "date-anomaly"
FLAG_LABEL_CONFLICT = "label-conflict"
FLAG_LENDER_UNMAPPED = "lender-unmapped"
# Advisory, outside the baseline 5 -- see the module docstring.
FLAG_RESUME_BEFORE_END = "resume-before-end"

MEMO_QUALITY_FLAGS = (
    FLAG_DATES_MISSING, FLAG_DATE_ANOMALY, FLAG_LABEL_CONFLICT,
    FLAG_LENDER_UNMAPPED, FLAG_RESUME_BEFORE_END,
)

_FLAG_COLUMNS = {flag: f"flag_{flag.replace('-', '_')}" for flag in MEMO_QUALITY_FLAGS}

# Scope label for a memo row whose Lender resolves to no servicer at all. Such
# a row belongs to the FILE, not to any servicer in the audit, and is counted
# separately so it is never presented as an uploaded servicer's issue.
MEMO_SCOPE_UNATTRIBUTABLE = "UNATTRIBUTABLE — file level"

_CLASSIFIER_COLUMNS = (
    ["memo_row_class", "memo_quality_flags", "memo_content_row",
     "date_order_violations", "date_order_detail"]
    + list(_FLAG_COLUMNS.values())
)


def _is_blank(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() == ""


def _column(df, name):
    """The named column, or an all-None column of the right length when the
    memo file simply doesn't have it. A missing column must degrade to "blank
    everywhere", never raise -- the memo template is hand-maintained and
    header detection is a best match, not a guarantee."""
    if name in df.columns:
        return df[name]
    return pd.Series([None] * len(df), index=df.index, dtype="object")


def _blank_mask(series):
    return series.map(_is_blank)


def _as_datetime(series):
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    return pd.to_datetime(series, errors="coerce")


def _normalized_action(series):
    return series.map(lambda v: "" if _is_blank(v) else str(v).strip().lower())


def _normalized_population(population_ids):
    """The population as normalized digit strings, matched the same way the
    memo's own gp_sba/gp_lender are normalized -- otherwise '66736791-10' in
    one file and 6673679110 in the other would read as different loans and
    every row would look orphaned."""
    normalized = set()
    for value in population_ids or ():
        gp = normalize_memo_loan_id(value)
        if gp:
            normalized.add(gp)
    return normalized


def classify_memo_rows(memo_df, population_ids):
    """Returns a copy of memo_df with the classification columns appended:

      memo_row_class         one of MEMO_ROW_CLASSES
      memo_quality_flags     "; "-joined flag names ("" when clean)
      flag_<name>            one bool column per MEMO_QUALITY_FLAGS entry
      memo_content_row       1-based ordinal among CONTENT rows (0 for
                             phantom rows) -- the row number a reviewer
                             counts to when phantom rows are ignored
      date_order_violations  0, 1 or 2 -- baseline rules only
      date_order_detail      human-readable description of each violation

    population_ids is required, not optional. Every row class depends on it,
    so defaulting it to None would silently relabel every orphan row as
    "joined" -- exactly the false-clean result this module exists to prevent.
    """
    if memo_df is None or len(memo_df) == 0:
        empty = pd.DataFrame(columns=list(getattr(memo_df, "columns", [])) + _CLASSIFIER_COLUMNS)
        return empty

    df = memo_df.copy()
    population = _normalized_population(population_ids)

    # --- row class -------------------------------------------------------
    phantom = (
        _blank_mask(_column(df, "Borrower Name"))
        & _blank_mask(_column(df, "SBA Loan ID"))
        & _blank_mask(_column(df, "Lender Loan ID"))
        & _blank_mask(_column(df, "Servicer Action"))
    )
    content = ~phantom

    gp_sba = _column(df, "gp_sba") if "gp_sba" in df.columns \
        else _column(df, "SBA Loan ID").map(normalize_memo_loan_id)
    gp_lender = _column(df, "gp_lender") if "gp_lender" in df.columns \
        else _column(df, "Lender Loan ID").map(normalize_memo_loan_id)

    in_population = gp_sba.isin(population) | gp_lender.isin(population)
    # CLIENT-CONFIRMED (2026-08-10): a row the SBA-Loan-ID -> Lender-Loan-Number
    # fallback attributed to a real loan is JOINED, not lookup-failed. Purely
    # additive -- rows the existing probes already matched are unaffected, and
    # a row left unresolved (unmatched, ambiguous or conflicting) keeps its
    # existing classification. resolve_memo_attribution is what populates this
    # column; when it has not been run the classification is exactly as before.
    if "resolved_gp" in df.columns:
        in_population = in_population | df["resolved_gp"].notna()

    df["memo_row_class"] = ROW_CLASS_JOINED
    df.loc[phantom, "memo_row_class"] = ROW_CLASS_PHANTOM
    df.loc[content & ~in_population, "memo_row_class"] = ROW_CLASS_LOOKUP_FAILED

    df["memo_content_row"] = 0
    df.loc[content, "memo_content_row"] = range(1, int(content.sum()) + 1)

    # --- quality flags (content rows only) -------------------------------
    start = _as_datetime(_column(df, "Deferral Start Date"))
    end = _as_datetime(_column(df, "Deferral End Date"))
    resume = _as_datetime(_column(df, "Regular Payment Resume Date"))

    action = _normalized_action(_column(df, "Securitization Action"))
    is_deferred = action == "deferred"
    action_present = action != ""

    dates_missing = content & is_deferred & start.isna()
    label_conflict = content & action_present & ~is_deferred & start.notna()

    end_before_start = content & start.notna() & end.notna() & (end < start)
    resume_before_start = content & start.notna() & resume.notna() & (resume < start)
    resume_before_end = content & end.notna() & resume.notna() & (resume < end)

    violations = end_before_start.astype(int) + resume_before_start.astype(int)
    df["date_order_violations"] = violations
    df["date_order_detail"] = [
        "; ".join(
            part for part, hit in (
                ("Deferral End Date is before Deferral Start Date", e),
                ("Regular Payment Resume Date is before Deferral Start Date", r),
                ("ADVISORY: Regular Payment Resume Date is before Deferral End Date", a),
            ) if hit
        )
        for e, r, a in zip(end_before_start, resume_before_start, resume_before_end)
    ]

    lender_unmapped = content & _column(df, "Lender").map(
        lambda v: memo_lender_soup_key(v) is None
    )

    flag_masks = {
        FLAG_DATES_MISSING: dates_missing,
        FLAG_DATE_ANOMALY: violations > 0,
        FLAG_LABEL_CONFLICT: label_conflict,
        FLAG_LENDER_UNMAPPED: lender_unmapped,
        FLAG_RESUME_BEFORE_END: resume_before_end,
    }
    for flag, mask in flag_masks.items():
        # A phantom row can never carry a flag: it asserts nothing to be
        # inconsistent about, and letting one through would put phantom rows
        # in the report and break the partition's meaning.
        df[_FLAG_COLUMNS[flag]] = (mask & content).fillna(False).astype(bool)

    df["memo_quality_flags"] = [
        "; ".join(flag for flag in MEMO_QUALITY_FLAGS if row[_FLAG_COLUMNS[flag]])
        for _, row in df.iterrows()
    ]

    return df


def memo_partition_counts(classified_df):
    """The partition/baseline counts. Row classes sum to total_rows; the flag
    counts overlap them and each other and are NOT part of that sum."""
    if classified_df is None or len(classified_df) == 0:
        zeros = {f"{flag.replace('-', '_')}": 0 for flag in MEMO_QUALITY_FLAGS}
        return {
            "total_rows": 0, "phantom": 0, "content": 0, "joined": 0,
            "lookup_failed": 0, "date_anomaly_violations": 0,
            "date_anomaly_rows": 0, **zeros,
        }

    cls = classified_df["memo_row_class"]
    counts = {
        "total_rows": int(len(classified_df)),
        "phantom": int((cls == ROW_CLASS_PHANTOM).sum()),
        "joined": int((cls == ROW_CLASS_JOINED).sum()),
        "lookup_failed": int((cls == ROW_CLASS_LOOKUP_FAILED).sum()),
    }
    counts["content"] = counts["joined"] + counts["lookup_failed"]
    for flag in MEMO_QUALITY_FLAGS:
        counts[flag.replace("-", "_")] = int(classified_df[_FLAG_COLUMNS[flag]].sum())
    # date_anomaly counts rows (it's a flag); the baseline is VIOLATIONS, so
    # both are reported under unambiguous names and the flag's own row count
    # is restated as date_anomaly_rows.
    counts["date_anomaly_violations"] = int(classified_df["date_order_violations"].sum())
    counts["date_anomaly_rows"] = counts["date_anomaly"]
    return counts


def unresolved_lender_spellings(classified_df):
    """Every distinct memo "Lender" spelling that MEMO_LENDER_ALIASES doesn't
    claim, with the evidence needed to add it: how many rows, which borrowers,
    which Securitization Actions. This is the output that turns "some deferrals
    are invisible" into a specific alias-table patch."""
    if classified_df is None or len(classified_df) == 0:
        return []
    unmapped = classified_df[classified_df[_FLAG_COLUMNS[FLAG_LENDER_UNMAPPED]]]
    if unmapped.empty:
        return []

    out = []
    lender = _column(unmapped, "Lender")
    for spelling in sorted({("" if _is_blank(v) else str(v).strip()) for v in lender}):
        rows = unmapped[lender.map(lambda v: ("" if _is_blank(v) else str(v).strip())) == spelling]
        out.append({
            "lender_as_written": spelling or "(blank)",
            "rows": int(len(rows)),
            "borrowers": sorted({str(v).strip() for v in _column(rows, "Borrower Name") if not _is_blank(v)}),
            "securitization_actions": sorted({str(v).strip() for v in _column(rows, "Securitization Action") if not _is_blank(v)}),
            "content_rows": [int(v) for v in rows["memo_content_row"]],
        })
    return out


def _format_date(value):
    if value is None or pd.isna(value):
        return ""
    try:
        return pd.Timestamp(value).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return str(value).strip()


# Blank Lender is its own case, not a spelling variant. Measured on the
# delivered file, ALL 8 unmapped rows have an empty Lender cell -- there is no
# alias that could resolve them, so pointing the reviewer at the alias table
# would send them to fix the wrong thing.
_HINT_LENDER_BLANK = (
    "The Lender cell on this memo row is EMPTY, so the row cannot be attributed "
    "to any servicer and no alias table entry can fix it. This needs a Lender "
    "filled in at source, in the Servicer Memo Summary itself."
)

_HINTS = {
    FLAG_LENDER_UNMAPPED: (
        "This exact Lender spelling is in no MEMO_LENDER_ALIASES entry, so no "
        "servicer key can match this row and any deferral on it is invisible to "
        "the audit. Add the spelling to the owning servicer's alias list in "
        "file_readers.py."
    ),
    ROW_CLASS_LOOKUP_FAILED: (
        "This memo row's loan ID appears in no uploaded population, so no loan "
        "iteration ever reaches it. Confirm the SBA Loan ID / Lender Loan ID "
        "against the SOUP dataset -- a mistyped ID here means a real deferral "
        "is never applied."
    ),
    FLAG_DATES_MISSING: (
        "Securitization Action says Deferred but no usable Deferral Start Date "
        "was written, so the deferral window cannot be placed on any cycle. "
        "Needs the real dates from the source memo."
    ),
    FLAG_DATE_ANOMALY: (
        "Impossible date order (see Date Order Detail) -- most often a "
        "year-entry typo. These dates drive deferral phase and the resume "
        "boundary, so they must be corrected at source, not computed through."
    ),
    FLAG_LABEL_CONFLICT: (
        "A non-Deferred Securitization Action carrying real deferral dates: the "
        "label and the dates disagree. Which field is authoritative is a "
        "business-rule question for the client -- surfaced here, not decided."
    ),
    FLAG_RESUME_BEFORE_END: (
        "ADVISORY (not counted in the date-anomaly baseline): Regular Payment "
        "Resume Date precedes Deferral End Date, which checks.py's own "
        "start <= end <= resume validity rule rejects."
    ),
}


def memo_coverage_records(classified_df, uploaded_soup_keys=None):
    """One record per PROBLEM row -- lookup-failed, or carrying any quality
    flag. Clean joined rows and all phantom rows produce nothing: this is an
    exception report, and the counts (memo_partition_counts) carry the
    denominators.

    uploaded_soup_keys scopes the output to the servicers actually uploaded
    for this audit. The memo workbook covers EVERY servicer, so without it a
    single-servicer audit reports other servicers' rows as if they were its
    own. Scoping is by the row's resolved Lender identity -- the same
    memo_lender_soup_key attribution already shown in the "Resolved Servicer
    Key" column, reused rather than reimplemented.

    Scoping is deliberately NOT done by narrowing population_ids: lookup-failed
    means "this loan ID is in no population", so a narrowed population would
    reclassify every other servicer's rows as defects. Measured on the
    delivered memo, that turns 22 findings into 195. Attribution and
    population are different questions and must stay that way.

    Rows whose Lender resolves to NO servicer are kept regardless of scope,
    tagged MEMO_SCOPE_UNATTRIBUTABLE. They cannot be tied to the uploaded
    servicer -- that inability IS the defect -- but dropping them would
    restore exactly the silent zero this check exists to prevent. The caller
    counts them separately so they are never presented as the uploaded
    servicer's issues.

    Key order here is the column order report.py's _build_detail_sheet will
    render, since it derives headers from the first record's keys."""
    if classified_df is None or len(classified_df) == 0:
        return []

    problem = classified_df[
        (classified_df["memo_row_class"] == ROW_CLASS_LOOKUP_FAILED)
        | (classified_df["memo_quality_flags"] != "")
    ]

    records = []
    for _, row in problem.iterrows():
        resolved = memo_lender_soup_key(row.get("Lender"))
        if resolved is None:
            scope = MEMO_SCOPE_UNATTRIBUTABLE
        elif uploaded_soup_keys is not None and resolved not in uploaded_soup_keys:
            continue  # belongs to a servicer not part of this audit
        else:
            scope = resolved
        flags = [f for f in MEMO_QUALITY_FLAGS if row[_FLAG_COLUMNS[f]]]
        hint_keys = ([ROW_CLASS_LOOKUP_FAILED] if row["memo_row_class"] == ROW_CLASS_LOOKUP_FAILED else []) + flags
        hints = [
            _HINT_LENDER_BLANK
            if key == FLAG_LENDER_UNMAPPED and _is_blank(row.get("Lender"))
            else _HINTS[key]
            for key in hint_keys
        ]
        records.append({
            # First column on purpose: a reviewer must see whose issue a row
            # is before reading anything else about it.
            "Memo Scope": scope,
            "Memo Sheet Row": row.get("_excel_row", ""),
            "Content Row #": int(row["memo_content_row"]),
            "Memo Row Class": row["memo_row_class"],
            "Data Quality Flags": row["memo_quality_flags"],
            "Lender (as written)": "" if _is_blank(row.get("Lender")) else str(row.get("Lender")).strip(),
            "Resolved Servicer Key": memo_lender_soup_key(row.get("Lender")) or "",
            "Borrower Name": "" if _is_blank(row.get("Borrower Name")) else str(row.get("Borrower Name")).strip(),
            "SBA Loan ID": "" if _is_blank(row.get("SBA Loan ID")) else str(row.get("SBA Loan ID")).strip(),
            "Lender Loan ID": "" if _is_blank(row.get("Lender Loan ID")) else str(row.get("Lender Loan ID")).strip(),
            "Report Date": _format_date(row.get("Report Date")),
            "Securitization Action": "" if _is_blank(row.get("Securitization Action")) else str(row.get("Securitization Action")).strip(),
            "Servicer Action": "" if _is_blank(row.get("Servicer Action")) else str(row.get("Servicer Action")).strip(),
            "Deferral Start Date": _format_date(row.get("Deferral Start Date")),
            "Deferral End Date": _format_date(row.get("Deferral End Date")),
            "Regular Payment Resume Date": _format_date(row.get("Regular Payment Resume Date")),
            "Start Date (as written)": row.get(MEMO_RAW_DATE_COLUMNS["Deferral Start Date"], ""),
            "End Date (as written)": row.get(MEMO_RAW_DATE_COLUMNS["Deferral End Date"], ""),
            "Resume Date (as written)": row.get(MEMO_RAW_DATE_COLUMNS["Regular Payment Resume Date"], ""),
            "Date Order Violations": int(row["date_order_violations"]),
            "Date Order Detail": row["date_order_detail"],
            "Resolution Hint": " ".join(hints),
        })
    return records


def build_memo_coverage_result(memo_df, population_ids, uploaded_soup_keys=None,
                               servicer_label="Memo Summary"):
    """The MEMO-01 result dict, in the same shape checks._result produces so
    it flows through main.py's response, report.py's tab builder, and
    checks.summarize untouched.

    Built here rather than in checks.py on purpose: this is a file-level
    diagnostic over the memo workbook, not a per-servicer reconciliation
    check, and run_servicer_reconciliation would run it once per uploaded
    servicer file against the same global memo_df.

    SCOPING. The memo workbook covers every servicer, but an audit covers only
    the servicer files uploaded for it. uploaded_soup_keys (the canonical
    soup_keys of those files) restricts the AUDIT-FACING output to rows that
    belong to them. Two counts are reported and never added together:

        issues_found        rows attributable to the uploaded servicers.
                            THIS is the audit's memo-coverage issue count.
        unattributable_count  rows whose Lender resolves to no servicer at
                            all. They are shown, tagged
                            MEMO_SCOPE_UNATTRIBUTABLE, but they are a
                            property of the FILE and are never counted as an
                            uploaded servicer's issues -- doing so would
                            assert they belong to a servicer we know they do
                            not belong to.

    Passing uploaded_soup_keys=None disables scoping and reports every row,
    which is what the standalone CLI wants.

    The full-file partition counts are still returned unchanged as diagnostic
    metadata (and on result["memo_file_counts"]), so the file-level view and
    the regression baselines built on it are unaffected by scoping.

    Returns (result_dict, counts). result_dict is None when there is no memo
    file -- absent memo stays the existing graceful degradation, never a
    finding.
    """
    if memo_df is None or len(memo_df) == 0:
        return None, memo_partition_counts(None)

    classified = classify_memo_rows(memo_df, population_ids)
    counts = memo_partition_counts(classified)
    records = memo_coverage_records(classified, uploaded_soup_keys)

    # Scoped rows first, then the file-level ones, so the servicer's own
    # issues are what a reviewer reads before anything that isn't theirs.
    scoped = [r for r in records if r["Memo Scope"] != MEMO_SCOPE_UNATTRIBUTABLE]
    unattributable = [r for r in records if r["Memo Scope"] == MEMO_SCOPE_UNATTRIBUTABLE]
    records = scoped + unattributable

    issues = len(scoped)
    unmapped = unresolved_lender_spellings(classified)
    scope_note = (
        f" Scoped to the servicer file(s) in this audit ({', '.join(sorted(uploaded_soup_keys))})."
        if uploaded_soup_keys else ""
    )
    unattributable_note = (
        f" Separately, {len(unattributable)} memo row(s) could not be attributed to ANY "
        f"servicer because the Lender cell is empty or unrecognized. They are listed below "
        f"under '{MEMO_SCOPE_UNATTRIBUTABLE}' and are NOT counted as issues for the "
        f"servicer(s) in this audit — they are a defect in the memo file itself and need "
        f"a Lender filled in at source."
        if unattributable else ""
    )

    if issues == 0:
        details = ("No issues found." + scope_note + unattributable_note).strip()
    else:
        details = (
            f"{issues} memo row(s) for the servicer(s) in this audit could not be relied on."
            + scope_note
            + f" Across the whole memo file: {counts['lookup_failed']} row(s) whose loan ID is "
            f"in no SOUP population, {counts['lender_unmapped']} under an unrecognized or "
            f"blank Lender"
            + (f" ({', '.join(u['lender_as_written'] for u in unmapped)})" if unmapped else "")
            + f", {counts['deferral_dates_missing']} asserting a deferral with no usable start "
            f"date, {counts['date_anomaly_violations']} impossible date ordering(s) across "
            f"{counts['date_anomaly_rows']} row(s), and {counts['label_conflict']} where a "
            f"non-Deferred action carries real deferral dates."
            + unattributable_note
            + " Each row below names its own defect; none of them is evidence that the loan "
            "has no deferral."
        )

    result = {
        "id": f"MEMO-01::{servicer_label}",
        "servicer": servicer_label,
        "name": "Servicer Memo Coverage",
        "category": "Data Quality",
        "severity": "WARN",
        "status": "PASS" if (issues == 0 and not unattributable) else "WARN",
        "issues_found": issues,
        "details": details,
        "records": records,
        # Rows shown but deliberately NOT counted in issues_found. main.py
        # reads this so the workbook tab still appears when the only findings
        # are file-level -- otherwise a servicer with no memo issues of its
        # own would hide them, which is the silent zero all over again.
        "unattributable_count": len(unattributable),
        # Full-file diagnostic metadata, unaffected by scoping.
        "memo_file_counts": counts,
    }
    return result, counts
