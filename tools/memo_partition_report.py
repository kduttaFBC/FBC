#!/usr/bin/env python3
"""
Servicer Memo Summary partition report — standalone diagnostic.

Prints, for a real memo drop:

  1. The partition/baseline table (phantom / joined / lookup-failed, plus the
     overlapping quality-flag counts) against the numbers measured from the
     delivered file, so a drift shows up as a diff rather than a surprise.
  2. Every memo "Lender" spelling MEMO_LENDER_ALIASES does not claim, with the
     row counts, borrowers and Securitization Actions behind it. This block is
     the input to an alias-table patch: a spelling no servicer claims can
     never match a servicer_key, so every deferral filed under it is invisible
     to the audit.
  3. Every memo row whose loan ID is in no SOUP population — the orphan class
     the per-loan lookup in checks._lookup_deferral structurally cannot see.

Reports only. Nothing here writes to a workbook, and nothing here is part of
the /api/audit pipeline (see backend/memo_audit.py, which this wraps).

Usage:
    python3 tools/memo_partition_report.py --memo MEMO.xlsx --soup SOUP.xlsx
        [--all-rows]        also list every flagged row, not just the summary
        [--csv OUT.csv]     write the full per-row classification to CSV

Exit codes:  0 = counts match the baseline   1 = a count moved   2 = read error
"""

from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "backend"))

import file_readers  # noqa: E402
import memo_audit  # noqa: E402


# Measured from the delivered Servicer Memo Summary THROUGH read_service_memo
# (docs/deferralevidencebrief.md). Kept in sync with the two baseline dicts in
# backend/test_memo_audit.py.
#
# phantom is 0 because python-calamine returns Excel error cells as '' and
# _rows_to_df drops the resulting all-blank rows. Reading the same sheet with
# openpyxl reports 431 rows / 233 phantom — a file the pipeline never sees.
#
# joined and lookup_failed depend on the population the memo IDs are matched
# against, so a move in EITHER of those two may be legitimate; the rest are
# fixed by the memo sheet alone. These two are measured against the real SOUP
# population — the brief's 189/9 came from the OriginationTapes proxy, which
# lacks AD TRANSPORTATION INC (5850799108).
# UPDATED 2026-08-10, client-approved: 198/190 -> 200/192. Two memo rows were
# appended (Excel rows 204 and 205 -- GP 1534469103 Community Bank & Trust and
# GP 3585327005 Mission Valley Bank, both Deferred, both joining to SOUP, both
# unflagged). No existing row changed and every quality-flag count below is
# unchanged. See backend/test_memo_audit.py for the full provenance note.
BASELINE = {
    "total_rows": 200,
    "phantom": 0,
    "content": 200,
    "joined": 192,
    "lookup_failed": 8,
    "deferral_dates_missing": 4,
    "date_anomaly_violations": 5,
    "date_anomaly_rows": 4,
    "label_conflict": 6,
}

# Counts whose baseline depends on what the memo IDs were matched against.
# Reported with a caveat instead of a flat "MOVED".
POPULATION_DEPENDENT = {"joined", "lookup_failed"}

_LABELS = {
    "total_rows": "Total rows on the sheet",
    "phantom": "  phantom (blank borrower + IDs + action)",
    "content": "  real content rows",
    "joined": "    joined to a SOUP population",
    "lookup_failed": "    LOOKUP-FAILED (ID in no population)",
    "deferral_dates_missing": "flag: deferral asserted, start date unusable",
    "date_anomaly_violations": "flag: date-order violations",
    "date_anomaly_rows": "      ...across distinct rows",
    "label_conflict": "flag: non-Deferred action with deferral dates",
    "lender_unmapped": "flag: unrecognized Lender spelling",
    "resume_before_end": "flag: resume before end (ADVISORY, not in baseline)",
}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--memo", required=True, help="Servicer Memo Summary .xlsx")
    parser.add_argument("--soup", required=True, help="SOUP Performance Dataset .xlsx")
    parser.add_argument("--all-rows", action="store_true",
                        help="list every flagged row, not just the summary")
    parser.add_argument("--csv", help="write the full per-row classification here")
    args = parser.parse_args(argv)

    try:
        with open(args.memo, "rb") as fh:
            memo_df = file_readers.read_service_memo(fh.read())
        with open(args.soup, "rb") as fh:
            soup_df = file_readers.read_soup_file(fh.read())
    except (OSError, ValueError) as exc:
        print(f"Could not read the input files: {exc}", file=sys.stderr)
        return 2

    population = soup_df["SBA GP Number"].dropna().tolist() \
        if "SBA GP Number" in soup_df.columns else []
    classified = memo_audit.classify_memo_rows(memo_df, population)
    counts = memo_audit.memo_partition_counts(classified)

    print("=" * 78)
    print("SERVICER MEMO SUMMARY — PARTITION REPORT")
    print("=" * 78)
    print(f"memo: {args.memo}")
    print(f"soup: {args.soup}  ({len(population)} loan IDs in population)")
    print()
    print(f"{'':52} {'measured':>9} {'baseline':>9}")
    drift = []
    for key, label in _LABELS.items():
        measured = counts.get(key, 0)
        expected = BASELINE.get(key)
        marker = ""
        if expected is not None and measured != expected:
            # A move in a population-dependent count may just mean this SOUP
            # export differs from the one the baseline was taken against, so
            # it is called out but never treated as drift in memo coverage.
            marker = "  <-- differs (population)" if key in POPULATION_DEPENDENT else "  <-- MOVED"
            if key not in POPULATION_DEPENDENT:
                drift.append((key, expected, measured))
        print(f"{label:52} {measured:>9} {('' if expected is None else expected):>9}{marker}")
    print()
    print("Row classes sum to the sheet; the flag counts OVERLAP them and each")
    print("other, so they are not part of that sum.")

    unmapped = memo_audit.unresolved_lender_spellings(classified)
    print()
    print("-" * 78)
    print("UNRECOGNIZED LENDER SPELLINGS — add these to MEMO_LENDER_ALIASES")
    print("-" * 78)
    if not unmapped:
        print("None. Every memo Lender spelling resolves to a servicer.")
    for entry in unmapped:
        print(f"\n  {entry['lender_as_written']!r}  ({entry['rows']} row(s))")
        print(f"    securitization actions : {', '.join(entry['securitization_actions']) or '(none)'}")
        print(f"    borrowers              : {', '.join(entry['borrowers'][:8]) or '(none)'}")
        print(f"    content row(s)         : {entry['content_rows']}")

    orphans = classified[classified["memo_row_class"] == memo_audit.ROW_CLASS_LOOKUP_FAILED]
    print()
    print("-" * 78)
    print("LOOKUP-FAILED ROWS — loan ID in no SOUP population")
    print("-" * 78)
    if orphans.empty:
        print("None. Every real memo row's loan ID is in the population.")
    for _, row in orphans.iterrows():
        print(f"  row {row.get('_excel_row', '?')} (content #{row['memo_content_row']}): "
              f"{row.get('Borrower Name', '')!r} | SBA {row.get('SBA Loan ID', '')!r} | "
              f"Lender ID {row.get('Lender Loan ID', '')!r} | "
              f"{row.get('Securitization Action', '')!r} | Lender {row.get('Lender', '')!r}")

    if args.all_rows:
        print()
        print("-" * 78)
        print("ALL FLAGGED ROWS")
        print("-" * 78)
        for rec in memo_audit.memo_coverage_records(classified):
            print(f"  row {rec['Memo Sheet Row']} [{rec['Memo Row Class']}] "
                  f"{rec['Data Quality Flags'] or '-'}: {rec['Borrower Name']!r} "
                  f"({rec['Securitization Action']!r}) "
                  f"{rec['Deferral Start Date'] or rec['Start Date (as written)']!r} -> "
                  f"{rec['Deferral End Date'] or '?'} / resume "
                  f"{rec['Regular Payment Resume Date'] or '?'}")

    if args.csv:
        classified.to_csv(args.csv, index=False)
        print(f"\nFull per-row classification written to {args.csv}")

    print()
    if drift:
        print("BASELINE MOVED:")
        for key, expected, measured in drift:
            print(f"  {key}: baseline {expected}, measured {measured}")
        print("A moved count is the signal, not the failure. Confirm it against this")
        print("memo drop, then update BASELINE here and in backend/test_memo_audit.py.")
        return 1

    print("All counts match the baseline.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
