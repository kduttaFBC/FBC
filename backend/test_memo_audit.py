"""
Tests for the Servicer Memo Summary coverage diagnostic (memo_audit.py) and
the alias tables it depends on.

Three layers:

  LAYER 1  Synthetic classifier tests. Always run. A hand-built memo sheet
           that hits every row class and every quality flag deliberately,
           including the adversarial rows the real file is known to contain.

  LAYER 2  The real-file baseline: 233 phantom / 189 joined / 9 lookup-failed
           / 4 deferral-dates-missing / 5 date-order violations across 4 rows
           / 6 label-conflict, measured from the delivered Servicer Memo
           Summary. SKIPPED unless the workbooks are present, since neither
           is in the repo. This is the regression guard: when the next memo
           drop moves these numbers, the diff is the signal — a jump in
           lookup-failed means an alias map broke; a jump in label-conflict
           means more deferrals are being filed under the wrong action.

               FALCONBRIDGE_MEMO_XLSX=/path/to/Servicer_Memo_Summary.xlsx \\
               FALCONBRIDGE_SOUP_XLSX=/path/to/SOUP.xlsx \\
               python3 -m pytest test_memo_audit.py

  LAYER 3  Alias-table drift guard. Always runs, needs no data. Asserts the
           backend's MEMO_LENDER_ALIASES actually covers every servicer the
           frontend can produce a soup_key for. This is the one that catches
           the NEXT silent zero-match — a servicer slot added in the UI with
           no memo alias entry — at test time instead of after a bad audit.

Run with:
    python -m pytest test_memo_audit.py
"""

import ast
import io
import os
import re
import unittest

import pandas as pd
from openpyxl import Workbook

import file_readers
import memo_audit
from memo_audit import (
    FLAG_DATES_MISSING, FLAG_DATE_ANOMALY, FLAG_LABEL_CONFLICT,
    FLAG_LENDER_UNMAPPED, FLAG_RESUME_BEFORE_END,
    ROW_CLASS_JOINED, ROW_CLASS_LOOKUP_FAILED, ROW_CLASS_PHANTOM,
)


# A Lender spelling that IS in MEMO_LENDER_ALIASES, so rows meant to test
# something other than lender mapping don't accidentally trip that flag.
KNOWN_LENDER = "Community Bank & Trust"


def _memo_row(**overrides):
    """One memo row with every column the classifier reads, defaulting to a
    clean joined Deferred row. Tests override only the field under test."""
    row = {
        "Report Date": pd.Timestamp("2026-02-12"),
        "Lender": KNOWN_LENDER,
        "SBA Loan ID": "6673679110",
        "Lender Loan ID": "6673679110",
        "Borrower Name": "Test Borrower LLC",
        "Securitization Action": "Deferred",
        "Servicer Action": "3-month full principal and interest deferral",
        "Deferral Start Date": pd.Timestamp("2026-01-01"),
        "Deferral End Date": pd.Timestamp("2026-03-01"),
        "Regular Payment Resume Date": pd.Timestamp("2026-04-01"),
    }
    row.update(overrides)
    return row


def _build_memo_df(rows):
    """Mirrors what read_service_memo produces: the derived gp_sba/gp_lender/
    deferral_type columns, the raw date snapshots, and coerced datetimes — so
    these tests exercise the same frame shape the real reader hands over."""
    df = pd.DataFrame(rows)
    df["gp_sba"] = df["SBA Loan ID"].apply(file_readers.normalize_memo_loan_id)
    df["gp_lender"] = df["Lender Loan ID"].apply(file_readers.normalize_memo_loan_id)
    df["deferral_type"] = df["Servicer Action"].apply(file_readers.classify_deferral_type)
    for col, raw_col in file_readers.MEMO_RAW_DATE_COLUMNS.items():
        df[raw_col] = df[col].map(lambda v: "" if v is None else str(v).strip())
    for col in ("Deferral Start Date", "Deferral End Date",
                "Regular Payment Resume Date", "Report Date"):
        df[col] = pd.to_datetime(df[col], errors="coerce")
    df["_excel_row"] = range(2, len(df) + 2)
    return df


# ---------------------------------------------------------------------------
# LAYER 1 — synthetic classifier behavior
# ---------------------------------------------------------------------------

class RowClassPartitionTests(unittest.TestCase):
    """The three row classes are exhaustive and mutually exclusive."""

    def setUp(self):
        self.rows = [
            _memo_row(),  # joined
            # orphan: a real content row whose loan ID is in no population.
            # No per-loan lookup can ever reach this row.
            _memo_row(**{"SBA Loan ID": "9999999999", "Lender Loan ID": "9999999999"}),
            # phantom: blank borrower + both IDs + servicer action. Survives
            # _rows_to_df's all-blank filter because another column (here
            # Securitization Action, in the real file a stale Deal formula
            # returning #N/A) still holds a value.
            _memo_row(**{
                "Borrower Name": "", "SBA Loan ID": "", "Lender Loan ID": "",
                "Servicer Action": "", "Securitization Action": "#N/A",
                "Deferral Start Date": None, "Deferral End Date": None,
                "Regular Payment Resume Date": None,
            }),
        ]
        self.df = memo_audit.classify_memo_rows(_build_memo_df(self.rows), ["6673679110"])

    def test_classes_assigned(self):
        self.assertEqual(
            list(self.df["memo_row_class"]),
            [ROW_CLASS_JOINED, ROW_CLASS_LOOKUP_FAILED, ROW_CLASS_PHANTOM],
        )

    def test_classes_are_exhaustive_and_sum_to_total(self):
        counts = memo_audit.memo_partition_counts(self.df)
        self.assertEqual(counts["total_rows"], 3)
        self.assertEqual(
            counts["phantom"] + counts["joined"] + counts["lookup_failed"],
            counts["total_rows"],
        )
        self.assertEqual(counts["content"], 2)

    def test_phantom_row_carries_no_flags(self):
        phantom = self.df[self.df["memo_row_class"] == ROW_CLASS_PHANTOM].iloc[0]
        self.assertEqual(phantom["memo_quality_flags"], "")

    def test_content_row_numbering_ignores_phantom_rows(self):
        # Phantom rows get 0 and don't consume an ordinal, so the numbering
        # matches what a reviewer counts when blanks are ignored.
        self.assertEqual(list(self.df["memo_content_row"]), [1, 2, 0])

    def test_id_normalization_joins_across_formats(self):
        # '66736791-10' (dashed str), 6673679110 (int) and '6673679110.0'
        # (float-ish) are the same loan; a population holding any one form
        # must join all three, or every row would look orphaned.
        rows = [
            _memo_row(**{"SBA Loan ID": "66736791-10", "Lender Loan ID": "66736791-10"}),
            _memo_row(**{"SBA Loan ID": 6673679110, "Lender Loan ID": 6673679110}),
            _memo_row(**{"SBA Loan ID": "6673679110.0", "Lender Loan ID": "6673679110.0"}),
        ]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["66736791-10"])
        self.assertEqual(set(df["memo_row_class"]), {ROW_CLASS_JOINED})

    def test_lender_loan_id_alone_is_enough_to_join(self):
        # Some servicers put a short internal reference in Lender Loan ID, so
        # matching must try BOTH columns rather than trusting either alone.
        rows = [_memo_row(**{"SBA Loan ID": "", "Lender Loan ID": "6673679110"})]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        self.assertEqual(df.iloc[0]["memo_row_class"], ROW_CLASS_JOINED)


class DeferralDatesMissingTests(unittest.TestCase):
    """Deferral asserted, dates absent — the third state that is neither
    'has evidence' nor 'no evidence'."""

    def test_na_string_start_date_is_flagged(self):
        rows = [_memo_row(**{"Deferral Start Date": "n/a"})]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        row = df.iloc[0]
        self.assertIn(FLAG_DATES_MISSING, row["memo_quality_flags"])
        # The literal text survives, which is the only thing that separates
        # "the servicer wrote n/a" from "the cell was empty" after
        # to_datetime(errors='coerce') flattens both to NaT.
        self.assertEqual(row["_raw_deferral_start"], "n/a")

    def test_blank_start_date_on_deferred_row_is_flagged(self):
        rows = [_memo_row(**{"Deferral Start Date": None})]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        self.assertIn(FLAG_DATES_MISSING, df.iloc[0]["memo_quality_flags"])

    def test_non_deferred_row_without_dates_is_not_flagged(self):
        # A Modification with no deferral dates asserts no deferral at all —
        # nothing missing, nothing to flag.
        rows = [_memo_row(**{
            "Securitization Action": "Modification", "Deferral Start Date": None,
            "Deferral End Date": None, "Regular Payment Resume Date": None,
        })]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        self.assertEqual(df.iloc[0]["memo_quality_flags"], "")


class DateOrderViolationTests(unittest.TestCase):
    """Violations, not rows — one row can break more than one ordering rule."""

    def test_resume_before_start_is_one_violation(self):
        rows = [_memo_row(**{"Regular Payment Resume Date": pd.Timestamp("2025-04-01")})]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        self.assertEqual(int(df.iloc[0]["date_order_violations"]), 1)
        self.assertIn(FLAG_DATE_ANOMALY, df.iloc[0]["memo_quality_flags"])

    def test_end_and_resume_both_before_start_is_two_violations(self):
        rows = [_memo_row(**{
            "Deferral Start Date": pd.Timestamp("2026-06-01"),
            "Deferral End Date": pd.Timestamp("2026-03-01"),
            "Regular Payment Resume Date": pd.Timestamp("2026-04-01"),
        })]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        self.assertEqual(int(df.iloc[0]["date_order_violations"]), 2)
        counts = memo_audit.memo_partition_counts(df)
        self.assertEqual(counts["date_anomaly_violations"], 2)
        self.assertEqual(counts["date_anomaly_rows"], 1)

    def test_violations_and_rows_are_counted_separately(self):
        # Three rows with one violation each plus one row with two = 5
        # violations across 4 rows, the real file's shape exactly.
        rows = [
            _memo_row(**{"Regular Payment Resume Date": pd.Timestamp("2025-04-01")}),
            _memo_row(**{"Regular Payment Resume Date": pd.Timestamp("2025-05-01")}),
            _memo_row(**{"Regular Payment Resume Date": pd.Timestamp("2025-06-01")}),
            _memo_row(**{
                "Deferral Start Date": pd.Timestamp("2026-06-01"),
                "Deferral End Date": pd.Timestamp("2026-03-01"),
                "Regular Payment Resume Date": pd.Timestamp("2026-04-01"),
            }),
        ]
        counts = memo_audit.memo_partition_counts(
            memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        )
        self.assertEqual(counts["date_anomaly_violations"], 5)
        self.assertEqual(counts["date_anomaly_rows"], 4)

    def test_resume_before_end_is_advisory_and_outside_the_baseline(self):
        # checks.py:948 requires start <= end <= resume, so this IS invalid
        # and worth reporting — but it must not move date_anomaly_violations,
        # which the regression baseline of 5 was measured without it.
        rows = [_memo_row(**{
            "Deferral Start Date": pd.Timestamp("2026-01-01"),
            "Deferral End Date": pd.Timestamp("2026-03-01"),
            "Regular Payment Resume Date": pd.Timestamp("2026-02-01"),
        })]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        counts = memo_audit.memo_partition_counts(df)
        self.assertEqual(counts["date_anomaly_violations"], 0)
        self.assertEqual(counts["date_anomaly_rows"], 0)
        self.assertEqual(counts["resume_before_end"], 1)
        self.assertIn(FLAG_RESUME_BEFORE_END, df.iloc[0]["memo_quality_flags"])
        self.assertNotIn(FLAG_DATE_ANOMALY, df.iloc[0]["memo_quality_flags"])

    def test_clean_window_has_no_violation(self):
        df = memo_audit.classify_memo_rows(_build_memo_df([_memo_row()]), ["6673679110"])
        self.assertEqual(int(df.iloc[0]["date_order_violations"]), 0)
        self.assertEqual(df.iloc[0]["memo_quality_flags"], "")


class LabelConflictTests(unittest.TestCase):
    """A non-Deferred action carrying real deferral dates. Reported, never
    resolved here — which field is authoritative is a business-rule question
    (brief step 2), and this module must not pre-empt it."""

    def test_modification_with_full_pni_window_is_flagged(self):
        # GP 6673679110 row 49 verbatim (docs/trace6673679110.md): a genuine
        # 3-month full-P&I deferral filed as Modification. checks.py:1068
        # drops it silently; here it is named.
        rows = [_memo_row(**{
            "Securitization Action": "Modification",
            "Servicer Action": "3-month full principal and interest deferral",
            "Deferral Start Date": pd.Timestamp("2026-01-01"),
            "Deferral End Date": pd.Timestamp("2026-03-01"),
            "Regular Payment Resume Date": pd.Timestamp("2026-04-01"),
        })]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        row = df.iloc[0]
        self.assertIn(FLAG_LABEL_CONFLICT, row["memo_quality_flags"])
        self.assertEqual(row["memo_row_class"], ROW_CLASS_JOINED)
        self.assertEqual(int(row["date_order_violations"]), 0)

    def test_flag_does_not_assert_a_deferral(self):
        # The whole point: a label conflict is surfaced as a conflict. Nothing
        # in the output may read as "this row is a deferral".
        rows = [_memo_row(**{"Securitization Action": "Modification"})]
        records = memo_audit.memo_coverage_records(
            memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        )
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["Securitization Action"], "Modification")
        self.assertIn("business-rule question", records[0]["Resolution Hint"])
        self.assertNotIn("Deferred", records[0]["Data Quality Flags"])

    def test_deferred_row_is_never_a_label_conflict(self):
        df = memo_audit.classify_memo_rows(_build_memo_df([_memo_row()]), ["6673679110"])
        self.assertNotIn(FLAG_LABEL_CONFLICT, df.iloc[0]["memo_quality_flags"])

    def test_blank_action_with_dates_is_not_a_label_conflict(self):
        # No label at all is a different (and weaker) claim than a label that
        # contradicts the dates; the baseline 6 counts only the latter.
        rows = [_memo_row(**{"Securitization Action": ""})]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        self.assertNotIn(FLAG_LABEL_CONFLICT, df.iloc[0]["memo_quality_flags"])


class LenderUnmappedTests(unittest.TestCase):
    """The silent-zero-match class: a Lender spelling no servicer claims."""

    def test_unknown_spelling_is_flagged(self):
        rows = [_memo_row(**{"Lender": "Oakley Hills Bancorp, TX"})]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        self.assertIn(FLAG_LENDER_UNMAPPED, df.iloc[0]["memo_quality_flags"])

    def test_known_alias_variant_is_not_flagged(self):
        # The catalogued variant spellings must resolve, or the report would
        # cry wolf on the very rows the alias table was built to handle.
        for spelling in ("Lendistry SBLC, LLC, CA", "Millennium Bank, IL",
                         "Mission Valley Bank, CA", "Peoples Bank, MS"):
            with self.subTest(spelling=spelling):
                rows = [_memo_row(**{"Lender": spelling})]
                df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
                self.assertNotIn(FLAG_LENDER_UNMAPPED, df.iloc[0]["memo_quality_flags"])

    def test_spelling_match_is_case_and_whitespace_insensitive(self):
        rows = [_memo_row(**{"Lender": "  lendistry sblc, llc, ca  "})]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        self.assertNotIn(FLAG_LENDER_UNMAPPED, df.iloc[0]["memo_quality_flags"])

    def test_blank_lender_is_flagged(self):
        # A content row with no Lender cannot be attributed to any servicer,
        # so it is unmapped by definition — not clean.
        rows = [_memo_row(**{"Lender": ""})]
        df = memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        self.assertIn(FLAG_LENDER_UNMAPPED, df.iloc[0]["memo_quality_flags"])

    def test_blank_lender_hint_does_not_send_the_reviewer_to_the_alias_table(self):
        # All 8 unmapped rows on the delivered file are blank, not misspelled.
        # No alias entry can resolve an empty cell, so the hint has to name the
        # source fix instead of the code fix.
        blank = memo_audit.memo_coverage_records(
            memo_audit.classify_memo_rows(_build_memo_df([_memo_row(**{"Lender": ""})]), ["6673679110"])
        )[0]
        self.assertIn("EMPTY", blank["Resolution Hint"])
        self.assertNotIn("MEMO_LENDER_ALIASES", blank["Resolution Hint"])

        misspelled = memo_audit.memo_coverage_records(
            memo_audit.classify_memo_rows(
                _build_memo_df([_memo_row(**{"Lender": "Millenium Bank"})]), ["6673679110"])
        )[0]
        self.assertIn("MEMO_LENDER_ALIASES", misspelled["Resolution Hint"])

    def test_unresolved_spellings_report_carries_patch_evidence(self):
        rows = [
            _memo_row(**{"Lender": "Oakley Hills Bancorp, TX", "Borrower Name": "Oakley Hills LLC"}),
            _memo_row(**{"Lender": "Oakley Hills Bancorp, TX", "Borrower Name": "BCRS Inc",
                         "Securitization Action": "Modification"}),
            _memo_row(),  # known lender — must not appear
        ]
        found = memo_audit.unresolved_lender_spellings(
            memo_audit.classify_memo_rows(_build_memo_df(rows), ["6673679110"])
        )
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0]["lender_as_written"], "Oakley Hills Bancorp, TX")
        self.assertEqual(found[0]["rows"], 2)
        self.assertEqual(found[0]["borrowers"], ["BCRS Inc", "Oakley Hills LLC"])
        self.assertEqual(found[0]["securitization_actions"], ["Deferred", "Modification"])


class CoverageRecordTests(unittest.TestCase):
    """The exception report: problem rows only, with the evidence to act."""

    def setUp(self):
        self.rows = [
            _memo_row(),                                                    # clean
            _memo_row(**{"SBA Loan ID": "9999999999", "Lender Loan ID": "9999999999"}),
            _memo_row(**{"Deferral Start Date": "n/a"}),
            _memo_row(**{"Borrower Name": "", "SBA Loan ID": "", "Lender Loan ID": "",
                         "Servicer Action": "", "Securitization Action": "#N/A",
                         "Deferral Start Date": None, "Deferral End Date": None,
                         "Regular Payment Resume Date": None}),
        ]
        self.classified = memo_audit.classify_memo_rows(_build_memo_df(self.rows), ["6673679110"])
        self.records = memo_audit.memo_coverage_records(self.classified)

    def test_only_problem_rows_are_reported(self):
        self.assertEqual(len(self.records), 2)
        self.assertEqual(
            {r["Memo Row Class"] for r in self.records},
            {ROW_CLASS_LOOKUP_FAILED, ROW_CLASS_JOINED},
        )

    def test_lookup_failed_row_says_so_rather_than_no_deferral(self):
        rec = next(r for r in self.records if r["Memo Row Class"] == ROW_CLASS_LOOKUP_FAILED)
        self.assertIn("no uploaded population", rec["Resolution Hint"])

    def test_record_points_at_the_source_row(self):
        rec = self.records[0]
        self.assertTrue(rec["Memo Sheet Row"])
        self.assertTrue(rec["Content Row #"])

    def test_records_share_one_key_order(self):
        # report.py derives the tab's headers from the FIRST record's keys, so
        # a record with different keys would silently lose columns.
        key_orders = {tuple(r.keys()) for r in self.records}
        self.assertEqual(len(key_orders), 1)


class MemoCoverageResultTests(unittest.TestCase):
    """MEMO-01's result dict — the loud report line."""

    def test_absent_memo_is_not_a_finding(self):
        result, counts = memo_audit.build_memo_coverage_result(None, ["6673679110"])
        self.assertIsNone(result)
        self.assertEqual(counts["total_rows"], 0)

    def test_clean_memo_passes(self):
        df = _build_memo_df([_memo_row()])
        result, _ = memo_audit.build_memo_coverage_result(df, ["6673679110"])
        self.assertEqual(result["status"], "PASS")
        self.assertEqual(result["issues_found"], 0)
        self.assertEqual(result["details"], "No issues found.")

    def test_result_shape_matches_the_other_checks(self):
        df = _build_memo_df([_memo_row(**{"Lender": "Unknown Bank, ZZ"})])
        result, _ = memo_audit.build_memo_coverage_result(df, ["6673679110"])
        for key in ("id", "servicer", "name", "category", "severity", "status",
                    "issues_found", "details", "records"):
            self.assertIn(key, result)
        self.assertEqual(result["id"].split("::")[0], "MEMO-01")
        self.assertEqual(result["status"], "WARN")
        self.assertEqual(result["issues_found"], 1)

    def test_details_states_the_denominator_and_disclaims_no_deferral(self):
        rows = [_memo_row(), _memo_row(**{"Lender": "Unknown Bank, ZZ"})]
        result, _ = memo_audit.build_memo_coverage_result(_build_memo_df(rows), ["6673679110"])
        self.assertIn("of 2 real memo row(s)", result["details"])
        self.assertIn("Unknown Bank, ZZ", result["details"])
        self.assertIn("none of them is evidence that the loan has no deferral",
                      result["details"])

    def test_empty_memo_frame_is_handled(self):
        result, counts = memo_audit.build_memo_coverage_result(pd.DataFrame(), ["1"])
        self.assertIsNone(result)
        self.assertEqual(counts["content"], 0)

    def test_missing_columns_degrade_instead_of_raising(self):
        # The memo template is hand-maintained and header detection is a best
        # match, not a guarantee; a missing column must not crash the audit.
        df = pd.DataFrame([{"SBA Loan ID": "6673679110", "Borrower Name": "X"}])
        classified = memo_audit.classify_memo_rows(df, ["6673679110"])
        self.assertEqual(classified.iloc[0]["memo_row_class"], ROW_CLASS_JOINED)


class DeferralTypeClassificationTests(unittest.TestCase):
    """classify_deferral_type — evaluation order and the widened partial set.

    Every string here is real Servicer Action text from the delivered Servicer
    Memo Summary, not invented. A false "full" is the dangerous direction: it
    makes STAT-01 advance the delinquency rank +1/cycle regardless of payment
    and it gates corrections.py's automatic Loan Status -> DEFERRED write.
    """

    def test_evaluation_order_full_beats_partial(self):
        # The ordering guarantee: text naming BOTH principal and interest
        # contains a principal-only phrase too, and must still read as full.
        for text in ("3-month full principal and interest deferral",
                     "deferment of principal and interest",
                     "deferral of principal & interest for 90 days"):
            with self.subTest(text=text):
                self.assertEqual(file_readers.classify_deferral_type(text), "full")

    def test_explicit_full_pni_spellings(self):
        for text in ("full P&I deferral", "P & I deferment", "3 mos P and I deferment",
                     "PNI deferral", "principal and interest", "principal & interest"):
            with self.subTest(text=text):
                self.assertEqual(file_readers.classify_deferral_type(text), "full")

    def test_short_form_tokens_are_word_bounded(self):
        # "pni" as a bare substring appears inside ordinary words. Under plain
        # substring matching this text would force "full" and mask a genuine
        # principal-only deferral; word-bounded, it correctly reads partial.
        text = "Compnie Holdings LLC - 3-month principal deferment"
        self.assertIn("pni", text.lower())  # the premise this test guards
        self.assertEqual(file_readers.classify_deferral_type(text), "partial")

    def test_principal_only_phrasings_now_classify_partial(self):
        # The 11 rows the old two-keyword list misread as full.
        for text in ("3- month principal deferment",
                     "3-month deferment of principal",
                     "3-month principal deferment",
                     "6M principal deferral",
                     "deferring principal for 3 months"):
            with self.subTest(text=text):
                self.assertEqual(file_readers.classify_deferral_type(text), "partial")

    def test_interest_still_being_paid_is_partial(self):
        # If the servicer says interest is still being paid, only principal is
        # deferred — the brief's headline false-full cases.
        for text in ("3-month deferment of principal paying interest",
                     "2-month deferment with $7200 interest payments"):
            with self.subTest(text=text):
                self.assertEqual(file_readers.classify_deferral_type(text), "partial")

    def test_signal_free_text_is_unclassified_not_full(self):
        # The 9 rows carrying no payment-type signal at all. Returning "full"
        # for these asserted the one thing the servicer never wrote down.
        for text in ("three months deferment", "one-time 90-day deferment",
                     "30-day deferment", "90-Day Deferment",
                     "3 month payment deferral", "One-Time 90-Day Defermen"):
            with self.subTest(text=text):
                self.assertEqual(file_readers.classify_deferral_type(text), "unclassified")

    def test_unclassified_is_never_returned_for_typed_text(self):
        self.assertEqual(file_readers.classify_deferral_type("full P&I deferral"), "full")
        self.assertEqual(file_readers.classify_deferral_type("6M principal deferral"), "partial")

    def test_original_two_keywords_still_work(self):
        self.assertEqual(file_readers.classify_deferral_type("principal only deferral"), "partial")
        self.assertEqual(file_readers.classify_deferral_type("interest only period"), "partial")

    def test_real_memo_corpus_full_partial_split(self):
        # Locks the measured split so a future keyword edit can't silently
        # move rows between full and partial.
        corpus_full = [
            "3-month full principal and interest deferral",
            "3 mos P and I deferment ",
            "Second time three-month deferment of principal and interest",
        ]
        corpus_partial = [
            "3-month deferment of principal paying interest",
            "2-month deferment with $7200 interest payments",
            "6M principal deferral",
            "principal only",
            "additional 6M of the interest only term",
        ]
        for t in corpus_full:
            with self.subTest(full=t):
                self.assertEqual(file_readers.classify_deferral_type(t), "full")
        for t in corpus_partial:
            with self.subTest(partial=t):
                self.assertEqual(file_readers.classify_deferral_type(t), "partial")


class ReaderPhantomRowTests(unittest.TestCase):
    """Why the real file measures 0 phantom, and what happens when a phantom
    row DOES survive the reader.

    These two are separate concerns and both matter. The real-file count of 0
    says the delivered sheet's blank rows never get past the reader. It does
    NOT say the phantom class is dead code — a stale row carrying a real
    Report Date survives, and the classifier has to recognize it rather than
    treat it as a content row with a missing ID. Built through real workbook
    bytes, since the behavior under test belongs to the reader, not the frame.
    """

    HEADERS = ["Deal", "Report Date", "Lender", "SBA Loan ID", "Lender Loan ID",
               "Borrower Name", "Servicer Action", "Deferral Start Date",
               "Deferral End Date", "Regular Payment Resume Date"]

    REAL_ROW = ["SOUP2025_1", "2026-05-26", KNOWN_LENDER, 6673679110, 6673679110,
                "Test Borrower LLC", "3-month full principal and interest deferral",
                "2026-04-01", "2026-06-01", "2026-07-01"]

    def _memo_bytes(self, data_rows):
        wb = Workbook()
        ws = wb.active
        ws.title = "Memo Summary"
        ws.cell(row=1, column=1, value="Servicer Memo Summary")  # title row above the header
        for c, h in enumerate(self.HEADERS, start=1):
            ws.cell(row=3, column=c, value=h)
        for i, row in enumerate(data_rows):
            for c, v in enumerate(row, start=1):
                ws.cell(row=4 + i, column=c, value=v)
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def test_na_only_row_never_reaches_the_classifier(self):
        # The measurement behind "phantom == 0" on the real file: calamine
        # returns an Excel error cell as '', so a row whose only remaining
        # content is a stale #N/A lookup is all-blank to _rows_to_df.
        na_only = ["#N/A"] + [None] * (len(self.HEADERS) - 1)
        df = file_readers.read_service_memo(self._memo_bytes([self.REAL_ROW, na_only]))
        self.assertEqual(len(df), 1, "an #N/A-only row should be dropped by the blank-row filter")

    def test_classifier_detects_a_phantom_that_does_survive(self):
        # A stale row still carrying a real Report Date DOES get past the
        # reader. It must classify as phantom — not as a content row missing
        # its ID, which would inflate the content denominator and put a blank
        # row in the exception report.
        dated_blank = ["SOUP2025_1", "2026-05-26"] + [None] * (len(self.HEADERS) - 2)
        df = file_readers.read_service_memo(self._memo_bytes([self.REAL_ROW, dated_blank]))
        self.assertEqual(len(df), 2)
        classified = memo_audit.classify_memo_rows(df, ["6673679110"])
        self.assertEqual(list(classified["memo_row_class"]),
                         [ROW_CLASS_JOINED, ROW_CLASS_PHANTOM])
        counts = memo_audit.memo_partition_counts(classified)
        self.assertEqual((counts["content"], counts["phantom"]), (1, 1))
        self.assertEqual(memo_audit.memo_coverage_records(classified), [])

    def test_reader_attaches_row_positions_and_raw_date_text(self):
        rows = [self.REAL_ROW,
                ["SOUP2025_1", "2026-05-26", KNOWN_LENDER, 5850799108, 5850799108,
                 "AD Transportation", "90-Day Deferment", "n/a", "n/a", "n/a"]]
        df = file_readers.read_service_memo(self._memo_bytes(rows))
        # header on Excel row 3, so the data starts at row 4 — this is what
        # points a reviewer at the cell to fix.
        self.assertEqual(list(df["_excel_row"]), [4, 5])
        self.assertEqual(list(df["_raw_deferral_start"]), ["2026-04-01", "n/a"])
        # ...while the coerced column still flattens "n/a" to NaT, which is
        # exactly why the raw snapshot has to exist.
        self.assertTrue(pd.isna(df["Deferral Start Date"].iloc[1]))


class DiagnosticIsolationTests(unittest.TestCase):
    """This module reports; it must not become part of the audit's decision
    path. Enforced as a test so the boundary can't erode quietly."""

    def test_memo_audit_does_not_import_the_reconciliation_engine(self):
        # Parsed, not grepped: the module's own docstring discusses checks.py
        # at length, and a substring search would flag the prose explaining
        # the boundary as a breach of it.
        with open(os.path.join(os.path.dirname(__file__), "memo_audit.py")) as fh:
            tree = ast.parse(fh.read())
        imported = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                imported.update(a.name.split(".")[0] for a in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                imported.add(node.module.split(".")[0])
        self.assertEqual(imported & {"checks", "corrections", "report", "main"}, set())

    def test_classifier_does_not_mutate_the_caller_frame(self):
        df = _build_memo_df([_memo_row()])
        before = list(df.columns)
        memo_audit.classify_memo_rows(df, ["6673679110"])
        self.assertEqual(list(df.columns), before)


# ---------------------------------------------------------------------------
# LAYER 2 — the real-file baseline (skipped unless the workbooks are present)
# ---------------------------------------------------------------------------

MEMO_PATH = os.environ.get("FALCONBRIDGE_MEMO_XLSX")
SOUP_PATH = os.environ.get("FALCONBRIDGE_SOUP_XLSX")

# Measured from the delivered Servicer Memo Summary THROUGH read_service_memo
# (docs/deferralevidencebrief.md). Measuring the same sheet with openpyxl
# instead reports 431 rows / 233 phantom — openpyxl surfaces the empty and
# #N/A rows that python-calamine returns as '' and _rows_to_df's blank filter
# then drops. The pipeline only ever uses calamine, so that reading is one
# this test can never produce and must not tolerate: accepting it would let a
# genuine regression pass as a reader difference.
#
# Split by what the numbers actually depend on:
#
#   POPULATION_INDEPENDENT  fixed by the memo sheet alone. Nothing about the
#                           SOUP population can move them, so they are hard-
#                           asserted whenever the memo file is available.
#   POPULATION_DEPENDENT    joined vs lookup-failed, which depend on what the
#                           loan IDs are matched against. These were measured
#                           against the OriginationTapes PROXY, not SOUP; the
#                           real population may legitimately shift them, so
#                           they are asserted only with a real SOUP file and a
#                           message that says a move needs interpreting rather
#                           than fixing.
POPULATION_INDEPENDENT_BASELINE = {
    "total_rows": 198,
    "phantom": 0,
    "content": 198,
    "deferral_dates_missing": 4,
    "date_anomaly_violations": 5,
    "date_anomaly_rows": 4,
    "label_conflict": 6,
}

# Measured against the real SOUP population (25,071 loan IDs). The brief's
# 189/9 came from the OriginationTapes proxy; the one row that differs is
# AD TRANSPORTATION INC (5850799108), which is absent from the tape but present
# in SOUP, so it joins here. Its memo row is a Charge-off under Lendistry, not
# a deferral — nothing about deferral coverage changed, only the population it
# was measured against.
#
# The remaining 8 orphans are 5 borrowers across 8 rows (Jkari, BCRS and Chico
# each appear twice): Oakley Hills, World Class Real Estate, Jkari Hospitality,
# BCRS, Chico Fine Art — the #N/A-deal set the brief names.
POPULATION_DEPENDENT_BASELINE = {
    "joined": 190,
    "lookup_failed": 8,
}


@unittest.skipUnless(MEMO_PATH, "set FALCONBRIDGE_MEMO_XLSX to run the real memo baseline")
class RealMemoFileBaselineTests(unittest.TestCase):
    """The counts the memo sheet fixes on its own.

    Classified against an EMPTY population deliberately: none of these numbers
    may depend on what the loan IDs are matched against, and running them with
    no population at all is the strongest available proof of that. (It makes
    every content row lookup-failed, which is precisely why joined/
    lookup-failed are asserted in the SOUP-backed class below, not here.)
    """

    @classmethod
    def setUpClass(cls):
        with open(MEMO_PATH, "rb") as fh:
            cls.memo_df = file_readers.read_service_memo(fh.read())
        cls.classified = memo_audit.classify_memo_rows(cls.memo_df, [])
        cls.counts = memo_audit.memo_partition_counts(cls.classified)

    def test_population_independent_baseline(self):
        for key, expected in POPULATION_INDEPENDENT_BASELINE.items():
            with self.subTest(key=key):
                self.assertEqual(
                    self.counts[key], expected,
                    f"{key}: baseline {expected}, measured {self.counts[key]}. "
                    f"A moved count is the signal, not the failure — a jump in "
                    f"label_conflict means more deferrals are being filed under the "
                    f"wrong Securitization Action; a jump in deferral_dates_missing "
                    f"or date_anomaly means the memo drop got sloppier. Confirm "
                    f"against the new file before updating the baseline.",
                )

    def test_the_reader_leaves_no_phantom_rows(self):
        # Stated separately because it is the assertion most likely to be
        # "fixed" wrongly: 0 is not an oversight. calamine returns #N/A as ''
        # and _rows_to_df drops the all-blank row, so the sheet's blank rows
        # never reach this module. If this ever fails with a nonzero phantom
        # count, the reader changed — not the classifier.
        self.assertEqual(self.counts["phantom"], 0)
        self.assertEqual(self.counts["total_rows"], self.counts["content"])

    def test_classes_still_sum_to_the_sheet(self):
        self.assertEqual(
            self.counts["phantom"] + self.counts["joined"] + self.counts["lookup_failed"],
            self.counts["total_rows"],
        )

    def test_flags_only_attach_to_content_rows(self):
        phantom = self.classified[self.classified["memo_row_class"] == ROW_CLASS_PHANTOM]
        self.assertTrue((phantom["memo_quality_flags"] == "").all())

    def test_every_unresolved_lender_spelling_is_named(self):
        # Printed rather than asserted: the point is to hand over the exact
        # strings needed to patch MEMO_LENDER_ALIASES.
        for entry in memo_audit.unresolved_lender_spellings(self.classified):
            print(f"UNMAPPED LENDER {entry['lender_as_written']!r}: "
                  f"{entry['rows']} row(s), actions={entry['securitization_actions']}, "
                  f"borrowers={entry['borrowers'][:5]}")


@unittest.skipUnless(
    MEMO_PATH and SOUP_PATH,
    "set FALCONBRIDGE_MEMO_XLSX and FALCONBRIDGE_SOUP_XLSX to run the population baseline",
)
class RealPopulationBaselineTests(unittest.TestCase):
    """joined vs lookup-failed against the real SOUP population."""

    @classmethod
    def setUpClass(cls):
        with open(MEMO_PATH, "rb") as fh:
            memo_df = file_readers.read_service_memo(fh.read())
        with open(SOUP_PATH, "rb") as fh:
            soup_df = file_readers.read_soup_file(fh.read())
        cls.classified = memo_audit.classify_memo_rows(
            memo_df, soup_df["SBA GP Number"].dropna().tolist()
        )
        cls.counts = memo_audit.memo_partition_counts(cls.classified)

    def test_population_dependent_baseline(self):
        for key, expected in POPULATION_DEPENDENT_BASELINE.items():
            with self.subTest(key=key):
                self.assertEqual(
                    self.counts[key], expected,
                    f"{key}: baseline {expected}, measured {self.counts[key]}. "
                    f"These two were measured against the OriginationTapes PROXY, "
                    f"not the SOUP population this test uses, so a difference here "
                    f"may be legitimate — interpret it before changing anything. "
                    f"What it CANNOT mean is 'no deferral': every lookup-failed row "
                    f"is a memo record the per-loan lookup never reaches.",
                )

    def test_orphan_rows_are_named_not_counted(self):
        orphans = self.classified[
            self.classified["memo_row_class"] == ROW_CLASS_LOOKUP_FAILED
        ]
        for _, row in orphans.iterrows():
            print(f"LOOKUP-FAILED row {row.get('_excel_row', '?')}: "
                  f"{row.get('Borrower Name')!r} SBA={row.get('SBA Loan ID')!r} "
                  f"lender={row.get('Lender')!r} action={row.get('Securitization Action')!r}")
        self.assertEqual(len(orphans), self.counts["lookup_failed"])


# ---------------------------------------------------------------------------
# LAYER 3 — alias-table drift guard (no data needed)
# ---------------------------------------------------------------------------

APP_JSX = os.path.join(os.path.dirname(__file__), "..", "frontend", "src", "App.jsx")

_COMMENT_RE = re.compile(r"^\s*//.*$", re.MULTILINE)
_QUOTED_RE = re.compile(r"'([^']*)'")


def _jsx_block(source, declaration, open_char, close_char):
    """The literal between declaration's opening bracket and its match, with
    // comment lines stripped. Deliberately crude — this only has to read two
    flat literals, and a real JS parser would be a dependency for nothing."""
    start = source.index(declaration) + len(declaration)
    depth = 0
    for i in range(start, len(source)):
        if source[i] == open_char:
            depth += 1
        elif source[i] == close_char:
            depth -= 1
            if depth == 0:
                return _COMMENT_RE.sub("", source[start:i])
    raise AssertionError(f"unterminated literal for {declaration!r}")


@unittest.skipUnless(os.path.exists(APP_JSX), "frontend/src/App.jsx not present")
class AliasTableDriftTests(unittest.TestCase):
    """Every soup_key the UI can produce must have a MEMO_LENDER_ALIASES entry.

    Without this, adding a servicer slot in the frontend silently yields zero
    memo matches for that servicer — the exact failure that hid ALL Millennium
    Bank deferrals until a real memo file surfaced the spelling variant. That
    was found by accident; this finds the next one at test time.
    """

    @classmethod
    def setUpClass(cls):
        source = open(APP_JSX).read()
        slots_block = _jsx_block(source, "export const SERVICER_SLOTS =", "[", "]")
        keys_block = _jsx_block(source, "export const SERVICER_SOUP_KEYS =", "{", "}")
        cls.slots = _QUOTED_RE.findall(slots_block)
        pairs = _QUOTED_RE.findall(keys_block)
        cls.soup_key_map = dict(zip(pairs[0::2], pairs[1::2]))

    def test_parsed_the_frontend_maps(self):
        # Guards the crude parser itself: if App.jsx is restructured and these
        # come back empty, the coverage assertion below would vacuously pass.
        self.assertGreaterEqual(len(self.slots), 8)
        self.assertEqual(self.soup_key_map.get("Lendistry 2025"), "Lendistry")
        self.assertEqual(self.soup_key_map.get("Community Bank and Trust"),
                         "Community Bank & Trust")

    def test_every_frontend_soup_key_has_a_memo_alias_entry(self):
        missing = sorted({
            self.soup_key_map.get(slot, slot) for slot in self.slots
        } - set(file_readers.MEMO_LENDER_ALIASES))
        self.assertEqual(missing, [], (
            f"servicer(s) {missing} can be uploaded from the UI but have no "
            f"MEMO_LENDER_ALIASES entry, so checks._lookup_deferral falls back "
            f"to fuzzy substring matching for them — add an entry per the "
            f"table's own instructions in file_readers.py."
        ))

    def test_no_alias_key_carries_a_note_series_suffix(self):
        # A note-series-suffixed key ("Millennium 2025") can never match:
        # _lookup_deferral is called with the canonical soup_key, and the memo
        # file has no note-series field at all.
        suffixed = [k for k in file_readers.MEMO_LENDER_ALIASES if re.search(r"\b\d{4}\b", k)]
        self.assertEqual(suffixed, [])

    def test_no_spelling_is_claimed_by_two_servicers(self):
        seen = {}
        collisions = []
        for soup_key, spellings in file_readers.MEMO_LENDER_ALIASES.items():
            for spelling in spellings:
                norm = spelling.strip().lower()
                if norm in seen and seen[norm] != soup_key:
                    collisions.append((spelling, seen[norm], soup_key))
                seen[norm] = soup_key
        self.assertEqual(collisions, [], (
            "a memo Lender spelling claimed by two servicers makes attribution "
            "ambiguous; _build_alias_index keeps only the first binding"
        ))

    def test_every_alias_spelling_resolves_through_the_public_helper(self):
        for soup_key, spellings in file_readers.MEMO_LENDER_ALIASES.items():
            for spelling in spellings:
                with self.subTest(spelling=spelling):
                    self.assertEqual(file_readers.memo_lender_soup_key(spelling), soup_key)


if __name__ == "__main__":
    unittest.main()
