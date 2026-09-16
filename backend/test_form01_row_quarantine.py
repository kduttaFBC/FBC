"""
FORM-01 row-scoped quarantine — regression tests.

WHAT CHANGED

    preflight.preflight_workbook is UNTOUCHED: it still detects and reports
    every defect exactly as before. What changed is the DISPOSITION main.py
    applies to a failure:

        before   any FAIL  -> HTTP 400, whole workbook rejected
        after    FORM-01 only, rows named -> workbook ACCEPTED, the affected
                                             LOANS quarantined
                 anything else            -> HTTP 400, unchanged

    A quarantined loan keeps its record visible in STAT-01, carries the
    SOURCE FORMULA INTEGRITY DEFECT reason, has its Calculated DQ Status
    withdrawn to UNVERIFIED, and is excluded from every correction candidate
    via corrections.WARNING_REASONS. No lender value is repaired, inferred,
    recalculated or overwritten, and no workbook byte is modified.

    FORM-02 and FORM-03 remain file-level failures. That is the property most
    worth protecting here, so it is asserted directly rather than assumed.

Run with:
    python -m unittest test_form01_row_quarantine.py
"""

import hashlib
import os
import unittest

import pandas as pd

import checks
import corrections
import preflight


DOWNLOADS = os.path.expanduser("~/Downloads")

PEOPLES = os.path.join(
    DOWNLOADS, "2026.08_Peoples Bank_SOUP 2024-1 SRF August 2026 v1.xlsx")
# Two Sep-11 Mission Valley saves whose formula columns carry no value cache
# (FORM-02). These must still be rejected outright.
FORM02_FILES = [
    os.path.join(DOWNLOADS,
                 "2026.08_Mission Valley Bank_SOUP 2024-1 SRF Template_August 2026.xlsx"),
    os.path.join(DOWNLOADS,
                 "2026.08_Mission Valley Bank_SOUP 2024-1 SRF Template_August 2026-5.xlsx"),
]
# The other seven current submissions, which must be unaffected.
CLEAN_FILES = [
    "2026.08_Community Bank and Trust_SOUP 2025-1 SRF Template_August 2026.xlsx",
    "2026.08_Lendistry_SOUP 2024-1 SRF Template_August 2026.xlsx",
    "2026.08_Lendistry_SOUP 2025-1 SRF Template_August 2026.xlsx",
    "2026.08_Millennium_SOUP 2024-1 SRF Template_August 2026.xlsx",
    "2026.08_Millennium_SOUP 2025-1 SFR Template_August 2026.xlsx",
    "2026.08_Mission Valley Bank_SOUP 2024-1 SRF Template_August 2026-7.xlsx",
    "2026.08_Poppy Bank_SOUP 2024-1 SRF August 2026.xlsx",
]


def _finding(code, severity="FAIL", column="Loan Payment Instance", rows=(19,)):
    return {"code": code, "severity": severity, "column": column,
            "rows": list(rows), "detail": "x"}


def _result(ok, findings):
    return {"ok": ok, "sheet": "Servicer Inputs", "data_rows": 56,
            "columns_checked": [], "findings": findings, "rows_affected": []}


class IsolatableRowsTests(unittest.TestCase):
    """preflight.isolatable_rows is the whole disposition decision. It is
    fail-closed: every answer but 'FORM-01 only, with rows' is None, and None
    means the caller rejects the file exactly as it always did."""

    def test_clean_workbook_yields_nothing_to_quarantine(self):
        self.assertIsNone(preflight.isolatable_rows(_result(True, [])))

    def test_form01_only_returns_its_rows(self):
        r = _result(False, [_finding("FORM-01", rows=[19]),
                            _finding("FORM-01", column="Servicing Fee", rows=[19, 22])])
        self.assertEqual(preflight.isolatable_rows(r), [19, 22])

    def test_form02_is_never_row_isolatable(self):
        r = _result(False, [_finding("FORM-02", rows=list(range(10, 123)))])
        self.assertIsNone(preflight.isolatable_rows(r))

    def test_form03_is_never_row_isolatable(self):
        r = _result(False, [_finding("FORM-03", column=None, rows=[])])
        self.assertIsNone(preflight.isolatable_rows(r))

    def test_form04_is_never_row_isolatable(self):
        r = _result(False, [_finding("FORM-04", column=None, rows=[])])
        self.assertIsNone(preflight.isolatable_rows(r))

    def test_one_non_isolatable_finding_condemns_the_whole_file(self):
        """The critical case: a FORM-01 sitting beside a FORM-02 must NOT be
        downgraded to a row quarantine. The file-level fault wins."""
        r = _result(False, [_finding("FORM-01", rows=[19]),
                            _finding("FORM-02", rows=[10, 11])])
        self.assertIsNone(preflight.isolatable_rows(r))

    def test_form01_naming_no_rows_cannot_be_quarantined(self):
        r = _result(False, [_finding("FORM-01", rows=[])])
        self.assertIsNone(preflight.isolatable_rows(r))

    def test_warn_only_findings_are_not_blocking(self):
        r = _result(False, [_finding("FORM-05", severity="WARN", column=None, rows=[])])
        self.assertIsNone(preflight.isolatable_rows(r))


class QuarantineTagSpellingTests(unittest.TestCase):
    """corrections.py is deliberately isolated from checks.py, so the tag is
    spelled as a literal in both. This is the guard that keeps the two copies
    identical -- without it a rename in checks.py would silently stop
    excluding quarantined loans from correction."""

    def test_tag_is_in_warning_reasons(self):
        self.assertIn(checks.SOURCE_FORMULA_INTEGRITY_QUARANTINE,
                      corrections.WARNING_REASONS)


def _si(gp, loan_status="ACTIVE", dq="CURRENT", received=500.0, scheduled=500.0):
    return pd.DataFrame([{
        "SBA GP Number": gp, "Loan Status": loan_status, "Delinquency Status": dq,
        "Payment Received": received, "Scheduled Payment": scheduled,
        "Beginning Balance": 1000.0, "Ending Balance": 500.0,
        "Loan Payment Instance": 1, "Last Payment Instance?": "Y",
    }])


def _soup(gp, loan_status="ACTIVE", dq="CURRENT"):
    return pd.DataFrame([{
        "SBA GP Number": gp, "Servicer": "S", "Payment Month": "2026-07-31",
        "Loan Status": loan_status, "Delinquency Status": dq,
        "Beginning Actual Balance": 1500.0, "Ending Actual Balance": 1000.0,
    }])


class QuarantineEffectTests(unittest.TestCase):
    """The behaviour on a loan, held against an otherwise clean fixture so the
    quarantine is the only variable."""

    GP = "1234567890"

    def _run(self, quarantined):
        return checks.status_match_check(
            _si(self.GP), _soup(self.GP), servicer_name="S",
            quarantined_gps=quarantined,
        )

    def test_without_quarantine_a_clean_loan_is_not_flagged(self):
        """Baseline. A fully-matching loan produces no record at all."""
        r = self._run(None)
        self.assertEqual(r["records"], [])
        self.assertEqual(r["audited_status_by_loan"][self.GP], "CURRENT")

    def test_quarantine_makes_a_clean_loan_visible(self):
        """The loan matched on every field, so it would have been skipped.
        Quarantine must keep it visible -- an invisible quarantine is the
        failure mode this whole change exists to avoid."""
        r = self._run({self.GP: "FORM-01 on row 19"})
        self.assertEqual(len(r["records"]), 1)

    def test_quarantine_withdraws_the_calculated_figure(self):
        r = self._run([self.GP])
        rec = r["records"][0]
        self.assertEqual(rec["Calculated DQ Status"], "UNVERIFIED")
        self.assertEqual(rec["DQ Status Match"], "REVIEW")

    def test_quarantine_tags_the_reason(self):
        rec = self._run([self.GP])["records"][0]
        self.assertIn(checks.SOURCE_FORMULA_INTEGRITY_QUARANTINE, rec["Reason"])

    def test_detail_is_carried_verbatim_to_the_reviewer(self):
        detail = "preflight FORM-01 flagged row 19, column Servicing Fee"
        rec = self._run({self.GP: detail})["records"][0]
        self.assertEqual(rec["Source Integrity Detail"], detail)

    def test_unquarantined_loan_has_blank_detail(self):
        gp2 = "9999999999"
        si = pd.concat([_si(self.GP), _si(gp2, dq="DELINQUENT")], ignore_index=True)
        soup = pd.concat([_soup(self.GP), _soup(gp2)], ignore_index=True)
        r = checks.status_match_check(si, soup, servicer_name="S",
                                      quarantined_gps=[self.GP])
        by_gp = {rec["SBA GP Number"]: rec for rec in r["records"]}
        self.assertEqual(by_gp[gp2]["Source Integrity Detail"], "")
        self.assertNotIn(checks.SOURCE_FORMULA_INTEGRITY_QUARANTINE,
                         by_gp[gp2]["Reason"])

    def test_audited_summary_never_sees_the_defective_figure(self):
        """The withdrawal happens BEFORE audited_status_by_loan is captured,
        so the Audited Portfolio Summary cannot bucket the loan on a value
        derived from overwritten cells."""
        r = self._run([self.GP])
        self.assertEqual(r["audited_status_by_loan"][self.GP], "UNVERIFIED")
        self.assertIsNone(checks._classify_delinquency_bucket("UNVERIFIED"))

    def test_expected_loan_status_is_not_withdrawn(self):
        """Loan Status comes from SOUP history and the memo -- independent
        evidence the overwritten SRF cells cannot corrupt. Blanking it would
        discard a sound finding."""
        rec = self._run([self.GP])["records"][0]
        self.assertTrue(str(rec["Expected Loan Status"]).strip())

    def test_other_loans_are_untouched_by_a_quarantine(self):
        gp2 = "9999999999"
        si = pd.concat([_si(self.GP), _si(gp2, dq="DELINQUENT")], ignore_index=True)
        soup = pd.concat([_soup(self.GP), _soup(gp2)], ignore_index=True)
        with_q = checks.status_match_check(si, soup, servicer_name="S",
                                           quarantined_gps=[self.GP])
        without = checks.status_match_check(si, soup, servicer_name="S")
        self.assertEqual(with_q["audited_status_by_loan"][gp2],
                         without["audited_status_by_loan"][gp2])

    def test_quarantine_does_not_alter_the_ladder(self):
        """Nothing about the delinquency calculation itself moved."""
        self.assertEqual(checks.PAYMENT_SUFFICIENCY_THRESHOLD, 0.80)
        self.assertEqual(checks._missed_count_to_status(2), "30 DAYS")
        self.assertEqual(checks._status_to_missed_count("30 DAYS"), 2)


@unittest.skipUnless(os.path.isfile(PEOPLES), f"lender file not present: {PEOPLES}")
class RealPeoplesBankAcceptanceTests(unittest.TestCase):
    """The live case. Skipped when the lender files are not on this machine --
    they are not in the repository."""

    @classmethod
    def setUpClass(cls):
        cls.raw = open(PEOPLES, "rb").read()
        cls.sha_before = hashlib.sha256(cls.raw).hexdigest()
        cls.pf = preflight.preflight_workbook(cls.raw, filename=os.path.basename(PEOPLES))

    def test_form01_is_still_detected(self):
        """The gate was not disabled or weakened."""
        self.assertFalse(self.pf["ok"])
        codes = {f["code"] for f in self.pf["findings"] if f["severity"] == "FAIL"}
        self.assertEqual(codes, {"FORM-01"})

    def test_workbook_is_now_accepted_and_scoped_to_its_rows(self):
        self.assertEqual(preflight.isolatable_rows(self.pf), [19])

    def test_rows_resolve_to_exactly_one_loan(self):
        import file_readers
        _, _, _, si, _ = file_readers.read_servicer_file(
            self.raw, filename=os.path.basename(PEOPLES))
        hit = si[si["_excel_row"].isin(set(preflight.isolatable_rows(self.pf)))]
        gps = sorted(hit["SBA GP Number"].dropna().astype(str).str.strip().unique())
        self.assertEqual(gps, ["2786359008"])

    def test_lender_bytes_are_never_modified(self):
        self.assertEqual(hashlib.sha256(open(PEOPLES, "rb").read()).hexdigest(),
                         self.sha_before)


class Form02And03StayFileLevelTests(unittest.TestCase):
    """The prohibition that matters most: this change must not have opened a
    hole in the other two codes."""

    @unittest.skipUnless(any(os.path.isfile(f) for f in FORM02_FILES),
                         "no FORM-02 fixture file present")
    def test_known_form02_files_are_still_rejected_whole(self):
        checked = 0
        for path in FORM02_FILES:
            if not os.path.isfile(path):
                continue
            r = preflight.preflight_workbook(open(path, "rb").read(),
                                             filename=os.path.basename(path))
            self.assertFalse(r["ok"], path)
            self.assertIn("FORM-02", {f["code"] for f in r["findings"]}, path)
            self.assertIsNone(preflight.isolatable_rows(r), path)
            checked += 1
        self.assertGreater(checked, 0)

    def test_form03_disposition_is_unchanged(self):
        r = _result(False, [_finding("FORM-03", column=None, rows=[])])
        self.assertIsNone(preflight.isolatable_rows(r))


class OtherCurrentSubmissionsUnaffectedTests(unittest.TestCase):
    """The seven other current lender files must pass preflight and produce
    no quarantine at all."""

    def test_each_clean_file_passes_with_no_quarantine(self):
        checked = 0
        for name in CLEAN_FILES:
            path = os.path.join(DOWNLOADS, name)
            if not os.path.isfile(path):
                continue
            r = preflight.preflight_workbook(open(path, "rb").read(), filename=name)
            self.assertTrue(r["ok"], f"{name} unexpectedly failed preflight")
            self.assertIsNone(preflight.isolatable_rows(r), name)
            checked += 1
        if checked == 0:
            self.skipTest("no current lender files present")


if __name__ == "__main__":
    unittest.main()
