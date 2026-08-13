"""Servicer Memo attribution: SBA Loan ID -> Lender Loan Number fallback
(client-confirmed 2026-08-10).

    1. PRIMARY   memo SBA Loan ID -> SOUP SBA GP Number
    2. FALLBACK  memo SBA Loan ID -> SOUP Lender Loan Number

The fallback runs ONLY when the primary finds nothing, and only when it
resolves to exactly one loan. It exists because the delivered memo carries
rows whose SBA Loan ID field holds the servicer's own loan number rather than
the SBA GP Number.

DELIBERATELY NOT IMPLEMENTED: Lender Loan ID -> Lender Loan Number, Lender
Loan ID -> SBA GP Number, or any other probe. Asserted below, so a later
widening has to be a decision rather than a drift.

Exact matching only -- no fuzzy matching, no borrower-name matching, no
filename inference, no servicer guessing. The MATCHED LOAN determines the
servicer; the memo's own Lender cell is never consulted by the fallback.

ATTRIBUTION IS INDEPENDENT OF DATA QUALITY. A row that becomes attributable
keeps every defect it had.

Real-file tests skip unless the fixtures are present:
    FALCONBRIDGE_MEMO_XLSX=... FALCONBRIDGE_SOUP_XLSX=... \\
    python -m pytest test_memo_attribution_fallback.py
"""

import os
import unittest

import pandas as pd

import file_readers
import memo_audit


MEMO_PATH = os.environ.get("FALCONBRIDGE_MEMO_XLSX")
SOUP_PATH = os.environ.get("FALCONBRIDGE_SOUP_XLSX")
REAL = bool(MEMO_PATH and SOUP_PATH)


def _soup(rows):
    """rows: (gp, lender_loan_number, servicer)"""
    return pd.DataFrame([{"SBA GP Number": gp, "Lender Loan Number": lln,
                          "Servicer": servicer} for gp, lln, servicer in rows])


def _memo(sba_ids, lender_ids=None):
    lender_ids = lender_ids or [None] * len(sba_ids)
    df = pd.DataFrame([{"SBA Loan ID": s, "Lender Loan ID": l}
                       for s, l in zip(sba_ids, lender_ids)])
    df["gp_sba"] = df["SBA Loan ID"].apply(file_readers.normalize_memo_loan_id)
    df["gp_lender"] = df["Lender Loan ID"].apply(file_readers.normalize_memo_loan_id)
    return df


def _resolve(memo_df, soup_df):
    return file_readers.resolve_memo_attribution(
        memo_df, file_readers.build_memo_loan_index(soup_df))


SOUP = _soup([("1000000001", "5555000001", "Bank A"),
              ("1000000002", "5555000002", "Bank B")])


class PrimaryMatchTakesPrecedenceTests(unittest.TestCase):

    def test_an_sba_gp_match_resolves_by_the_primary_probe(self):
        out = _resolve(_memo(["1000000001"]), SOUP)
        self.assertEqual(out["resolved_gp"][0], "1000000001")
        self.assertEqual(out["attribution_method"][0], file_readers.ATTRIBUTION_BY_SBA_GP)
        self.assertEqual(out["resolved_servicer"][0], "Bank A")

    def test_the_primary_wins_when_the_value_is_also_a_lender_loan_number(self):
        # A value that is BOTH a valid GP and a valid Lender Loan Number for a
        # DIFFERENT loan must resolve to the GP. 158 such values exist in the
        # real population, so precedence is not academic.
        soup = _soup([("1000000001", "9999999999", "Bank A"),
                      ("2000000002", "1000000001", "Bank B")])
        out = _resolve(_memo(["1000000001"]), soup)
        self.assertEqual(out["resolved_gp"][0], "1000000001")
        self.assertEqual(out["attribution_method"][0], file_readers.ATTRIBUTION_BY_SBA_GP)

    def test_that_collision_is_surfaced_as_a_conflict(self):
        soup = _soup([("1000000001", "9999999999", "Bank A"),
                      ("2000000002", "1000000001", "Bank B")])
        out = _resolve(_memo(["1000000001"]), soup)
        self.assertEqual(out["attribution_issue"][0], file_readers.ATTRIBUTION_ISSUE_CONFLICT)

    def test_no_conflict_is_raised_when_both_probes_agree(self):
        soup = _soup([("1000000001", "1000000001", "Bank A")])
        out = _resolve(_memo(["1000000001"]), soup)
        self.assertEqual(out["attribution_issue"][0], "")


class FallbackTests(unittest.TestCase):

    def test_an_exact_unique_lender_loan_number_match_resolves(self):
        out = _resolve(_memo(["5555000002"]), SOUP)
        self.assertEqual(out["resolved_gp"][0], "1000000002")
        self.assertEqual(out["attribution_method"][0],
                         file_readers.ATTRIBUTION_BY_SBA_LENDER_LOAN_NUMBER)

    def test_the_servicer_comes_from_the_matched_loan(self):
        out = _resolve(_memo(["5555000002"]), SOUP)
        self.assertEqual(out["resolved_servicer"][0], "Bank B")

    def test_the_memo_lender_cell_is_never_consulted(self):
        # Every row the fallback resolves in the real drop has a BLANK Lender
        # cell. Attribution must not depend on it -- nor be misled by a wrong one.
        memo = _memo(["5555000002"])
        memo["Lender"] = "Some Other Bank Entirely"
        out = _resolve(memo, SOUP)
        self.assertEqual(out["resolved_gp"][0], "1000000002")
        self.assertEqual(out["resolved_servicer"][0], "Bank B")

    def test_normalisation_covers_only_harmless_representation(self):
        for value in ("5555000002", 5555000002, 5555000002.0, " 5555000002 ", "5555000002.0"):
            with self.subTest(value=value):
                out = _resolve(_memo([value]), SOUP)
                self.assertEqual(out["resolved_gp"][0], "1000000002")

    def test_a_near_miss_does_not_resolve(self):
        # No fuzzy matching: one digit out is not a match.
        out = _resolve(_memo(["5555000003"]), SOUP)
        self.assertIsNone(out["resolved_gp"][0])


class AmbiguityAndRefusalTests(unittest.TestCase):

    def test_a_lender_loan_number_matching_two_loans_is_refused(self):
        soup = _soup([("1000000001", "7777777777", "Bank A"),
                      ("1000000002", "7777777777", "Bank B")])
        out = _resolve(_memo(["7777777777"]), soup)
        self.assertIsNone(out["resolved_gp"][0])
        self.assertEqual(out["attribution_issue"][0], file_readers.ATTRIBUTION_ISSUE_AMBIGUOUS)
        self.assertEqual(out["resolved_servicer"][0], "")

    def test_an_unknown_identifier_stays_unresolved(self):
        out = _resolve(_memo(["8888888888"]), SOUP)
        self.assertIsNone(out["resolved_gp"][0])
        self.assertEqual(out["attribution_issue"][0], "")
        self.assertEqual(out["attribution_method"][0], "")

    def test_a_blank_identifier_stays_unresolved(self):
        for value in (None, "", "   "):
            with self.subTest(value=value):
                self.assertIsNone(_resolve(_memo([value]), SOUP)["resolved_gp"][0])


class OnlyProbeThreeIsImplementedTests(unittest.TestCase):
    """The other probes are deliberately absent; widening must be a decision."""

    def test_lender_loan_id_to_lender_loan_number_is_NOT_implemented(self):
        # SBA field unknown, Lender Loan ID would match a Lender Loan Number.
        out = _resolve(_memo(["8888888888"], ["5555000002"]), SOUP)
        self.assertIsNone(out["resolved_gp"][0])

    def test_lender_loan_id_to_sba_gp_is_NOT_implemented_by_this_resolver(self):
        out = _resolve(_memo(["8888888888"], ["1000000001"]), SOUP)
        self.assertIsNone(out["resolved_gp"][0])


class DegradesSafelyTests(unittest.TestCase):

    def test_soup_without_a_lender_loan_number_column_disables_only_the_fallback(self):
        soup = pd.DataFrame([{"SBA GP Number": "1000000001", "Servicer": "Bank A"}])
        out = _resolve(_memo(["1000000001", "5555000002"]), soup)
        self.assertEqual(out["resolved_gp"][0], "1000000001")     # primary still works
        self.assertIsNone(out["resolved_gp"][1])                  # fallback has no index

    def test_an_empty_population_resolves_nothing(self):
        out = _resolve(_memo(["1000000001"]), pd.DataFrame())
        self.assertIsNone(out["resolved_gp"][0])

    def test_the_input_frame_is_not_mutated(self):
        memo = _memo(["5555000002"])
        _resolve(memo, SOUP)
        for column in file_readers.MEMO_ATTRIBUTION_COLUMNS:
            self.assertNotIn(column, memo.columns)


class DataQualityIsIndependentOfAttributionTests(unittest.TestCase):
    """A row becoming attributable must not erase any existing defect.

    "Defect" means what the row says about ITSELF -- its dates, its labels.
    lender-unmapped is the one flag that is not that: it is an attribution
    verdict ("nothing identifies this row's servicer"), so attribution is
    exactly what is supposed to clear it. Asserting whole-flag-string equality
    swept the two together and, until 2026-08-12, pinned the incomplete wiring
    in place -- see test_memo_attribution_wiring.py."""

    def _classified(self, attribute):
        memo = pd.DataFrame([{
            "Borrower Name": "T", "SBA Loan ID": "5555000002", "Lender Loan ID": "",
            "Securitization Action": "Deferred", "Servicer Action": "3 month deferment",
            "Deferral Start Date": pd.Timestamp("2026-05-01"),
            "Deferral End Date": pd.Timestamp("2026-03-01"),      # end BEFORE start
            "Regular Payment Resume Date": pd.NaT, "Lender": "",
        }])
        memo["gp_sba"] = memo["SBA Loan ID"].apply(file_readers.normalize_memo_loan_id)
        memo["gp_lender"] = memo["Lender Loan ID"].apply(file_readers.normalize_memo_loan_id)
        if attribute:
            memo = _resolve(memo, SOUP)
        return memo_audit.classify_memo_rows(memo, ["1000000001", "1000000002"])

    def test_attribution_moves_the_row_class_to_joined(self):
        self.assertEqual(self._classified(False)["memo_row_class"][0],
                         memo_audit.ROW_CLASS_LOOKUP_FAILED)
        self.assertEqual(self._classified(True)["memo_row_class"][0],
                         memo_audit.ROW_CLASS_JOINED)

    @staticmethod
    def _flags(classified):
        return set(f for f in classified["memo_quality_flags"][0].split("; ") if f)

    def test_the_date_anomaly_survives_attribution(self):
        before = self._classified(False)
        after = self._classified(True)
        self.assertGreater(int(before["date_order_violations"][0]), 0)
        self.assertEqual(int(after["date_order_violations"][0]),
                         int(before["date_order_violations"][0]))
        self.assertIn(memo_audit.FLAG_DATE_ANOMALY, self._flags(after))

    def test_attribution_clears_only_the_lender_unmapped_flag(self):
        # The precise statement the old whole-string assertion was reaching
        # for: attribution changes the attribution verdict and NOTHING else.
        before, after = self._flags(self._classified(False)), self._flags(self._classified(True))
        self.assertIn(memo_audit.FLAG_LENDER_UNMAPPED, before)
        self.assertNotIn(memo_audit.FLAG_LENDER_UNMAPPED, after)
        self.assertEqual(before - {memo_audit.FLAG_LENDER_UNMAPPED}, after)


@unittest.skipUnless(REAL, "set FALCONBRIDGE_MEMO_XLSX and FALCONBRIDGE_SOUP_XLSX")
class RealJuneFixtureTests(unittest.TestCase):
    """The six known June rows, and the two that must stay unresolved."""

    EXPECTED_ROWS = [113, 114, 142, 143, 144, 145]
    EXPECTED_LOANS = {("2830919106", "Mission Valley Bank"),
                      ("6806045010", "Poppy Bank"),
                      ("7438795001", "Poppy Bank")}

    @classmethod
    def setUpClass(cls):
        cls.soup = file_readers.read_soup_file(open(SOUP_PATH, "rb").read())
        cls.memo = file_readers.read_service_memo(open(MEMO_PATH, "rb").read())
        cls.resolved = _resolve(cls.memo, cls.soup)

    def test_exactly_six_rows_resolve_through_the_fallback(self):
        fallback = self.resolved[
            self.resolved["attribution_method"]
            == file_readers.ATTRIBUTION_BY_SBA_LENDER_LOAN_NUMBER]
        self.assertEqual(sorted(int(r) for r in fallback["_excel_row"]), self.EXPECTED_ROWS)

    def test_they_resolve_to_exactly_three_loans_with_the_expected_servicers(self):
        fallback = self.resolved[
            self.resolved["attribution_method"]
            == file_readers.ATTRIBUTION_BY_SBA_LENDER_LOAN_NUMBER]
        pairs = {(r["resolved_gp"], r["resolved_servicer"]) for _, r in fallback.iterrows()}
        self.assertEqual(pairs, self.EXPECTED_LOANS)

    def test_rows_96_and_112_remain_unresolved(self):
        unresolved = self.resolved[self.resolved["resolved_gp"].isna()]
        self.assertEqual(sorted(int(r) for r in unresolved["_excel_row"]), [96, 112])

    def test_no_conflicts_or_ambiguity_in_the_delivered_memo(self):
        issues = [i for i in self.resolved["attribution_issue"] if i]
        self.assertEqual(issues, [])

    def test_coverage_moves_192_8_to_198_2(self):
        population = self.soup["SBA GP Number"].dropna().tolist()
        before = memo_audit.memo_partition_counts(
            memo_audit.classify_memo_rows(self.memo, population))
        after = memo_audit.memo_partition_counts(
            memo_audit.classify_memo_rows(self.resolved, population))
        self.assertEqual((before["joined"], before["lookup_failed"]), (192, 8))
        self.assertEqual((after["joined"], after["lookup_failed"]), (198, 2))

    def _before_after(self):
        population = self.soup["SBA GP Number"].dropna().tolist()
        return (memo_audit.memo_partition_counts(
                    memo_audit.classify_memo_rows(self.memo, population)),
                memo_audit.memo_partition_counts(
                    memo_audit.classify_memo_rows(self.resolved, population)))

    def test_every_data_quality_flag_count_except_attribution_is_unchanged(self):
        # AMENDED 2026-08-12, same reason as the synthetic twin above and as
        # test_memo_attribution_wiring.py: lender_unmapped is an attribution
        # verdict, not something the row asserts about itself, so freezing it
        # here protected the incomplete wiring instead of guarding it. Pinned
        # explicitly in the next test.
        before, after = self._before_after()
        for flag in ("deferral_dates_missing", "date_anomaly_violations",
                     "date_anomaly_rows", "label_conflict",
                     "total_rows", "content", "phantom"):
            with self.subTest(flag=flag):
                self.assertEqual(after[flag], before[flag])

    def test_lender_unmapped_moves_8_to_2(self):
        # The six fallback-resolved rows all have a blank Lender cell, so
        # before attribution reaches memo_audit they are indistinguishable from
        # the two genuinely unattributable ones. Only rows 96 and 112 -- whose
        # IDs appear nowhere in SOUP under either identifier -- may remain.
        before, after = self._before_after()
        self.assertEqual(before["lender_unmapped"], 8)
        self.assertEqual(after["lender_unmapped"], 2)

    def test_only_rows_96_and_112_stay_unattributable_in_the_report(self):
        population = self.soup["SBA GP Number"].dropna().tolist()
        records = memo_audit.memo_coverage_records(
            memo_audit.classify_memo_rows(self.resolved, population))
        unattributable = sorted(int(r["Memo Sheet Row"]) for r in records
                                if r["Memo Scope"] == memo_audit.MEMO_SCOPE_UNATTRIBUTABLE)
        self.assertEqual(unattributable, [96, 112])

    def test_the_fallback_rows_carry_their_matched_servicer_in_the_report(self):
        population = self.soup["SBA GP Number"].dropna().tolist()
        classified = memo_audit.classify_memo_rows(self.resolved, population)
        by_row = {int(r["_excel_row"]): r for _, r in classified.iterrows()}
        expected = {113: "Mission Valley Bank", 114: "Mission Valley Bank",
                    142: "Poppy Bank", 143: "Poppy Bank",
                    144: "Poppy Bank", 145: "Poppy Bank"}
        for row, servicer in expected.items():
            with self.subTest(memo_row=row):
                self.assertEqual(memo_audit._row_soup_key(by_row[row]), servicer)


if __name__ == "__main__":
    unittest.main()
