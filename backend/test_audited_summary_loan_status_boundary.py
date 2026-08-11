"""The Audited Portfolio Summary's Loan Status boundary (Finding 9 regression).

The three Portfolio Summary layers deliberately answer three different
questions, and this file pins the boundary between them for the one case that
makes the distinction visible: a loan the audit says is ACTIVE while the
servicer reports LIQUIDATION.

    Original         what the LENDER reported
    Audited          what the AUDIT calculated -- CLIENT-CONFIRMED (2026-08-04)
                     as "identical to portfolio_summary_check in every respect
                     EXCEPT the aging-bucket source". The audit's calculated
                     DELINQUENCY status drives the aging buckets; the LENDER's
                     Loan Status still drives Outstanding, Liquidation,
                     Deferred, Bankruptcy, Foreclosure and Paid In Full,
                     because a proposed Loan Status correction is a
                     RECOMMENDATION that has not been approved yet.
    Final Audited    what the REVIEWER verified -- the only layer where an
                     approved Loan Status correction moves money.

WHY THIS FILE EXISTS. Until 2026-08-08 the audit could not derive a Loan
Status that contradicted a servicer-reported terminal status: the terminal
branch copied the servicer's value, so expected and reported could never
disagree. That capability now exists (the confirmed-resume derivation), which
makes it possible for the Audited summary to report a loan in Liquidation
while the same audit run recommends moving it to ACTIVE. That is CORRECT --
the correction requires reviewer approval and the Audited summary is generated
during /api/audit, before any reviewer decision exists -- but nothing pinned
it, so a future change could silently move money into Outstanding without an
approval and no test would fail.

The real-world case is GP 6717929105 (July 2026, Community Bank and Trust):
servicer LIQUIDATION / 60 DAYS, memo-confirmed resume 2026-07-05, audit
derives ACTIVE / DELINQUENT, balance 843,761.03. The synthetic fixture below
reproduces that shape exactly; the real-file test at the end pins the actual
loan when the fixtures are available.

NO PRODUCTION BEHAVIOUR IS CHANGED BY THIS FILE. It asserts what the code
already does.

Run with:
    python -m pytest test_audited_summary_loan_status_boundary.py
"""

import os
import unittest

import pandas as pd

import checks
import corrections
import file_readers


GP = "1000000001"          # the contested loan: servicer LIQUIDATION, audit ACTIVE
OTHER = "1000000002"       # an ordinary loan, so Outstanding is never trivially zero
BALANCE = 843761.03        # GP 6717929105's real balance
OTHER_BALANCE = 100000.00


def _servicer_inputs():
    return pd.DataFrame([
        {"SBA GP Number": GP, "Lender Loan Number": "L-1",
         "Loan Status": "LIQUIDATION", "Delinquency Status": "60 DAYS",
         "Beginning Balance": BALANCE, "Ending Balance": BALANCE,
         "Payment Received": 0.0, "Scheduled Payment": 0.0, "Servicing Fee": 0.0,
         "_excel_row": 10, "_sheet_name": "Servicer Inputs", "_header_excel_row": 9},
        {"SBA GP Number": OTHER, "Lender Loan Number": "L-2",
         "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
         "Beginning Balance": OTHER_BALANCE, "Ending Balance": OTHER_BALANCE,
         "Payment Received": 500.0, "Scheduled Payment": 500.0, "Servicing Fee": 0.0,
         "_excel_row": 11, "_sheet_name": "Servicer Inputs", "_header_excel_row": 9},
    ])


def _soup():
    """SOUP through 2026-06 so the inferred audit cycle is 2026-07 -- the month
    the memo's Regular Payment Resume Date falls in."""
    rows = [{"SBA GP Number": GP, "Servicer": "Test Bank", "Payment Month": month,
             "Loan Status": "DEFERRED", "Delinquency Status": dq,
             "Beginning Actual Balance": BALANCE, "Ending Actual Balance": BALANCE}
            for month, dq in (("2026-04-01", "DELINQUENT"), ("2026-05-01", "30 DAYS"),
                              ("2026-06-01", "60 DAYS"))]
    rows.append({"SBA GP Number": OTHER, "Servicer": "Test Bank",
                 "Payment Month": "2026-06-01", "Loan Status": "ACTIVE",
                 "Delinquency Status": "CURRENT",
                 "Beginning Actual Balance": OTHER_BALANCE,
                 "Ending Actual Balance": OTHER_BALANCE})
    return pd.DataFrame(rows)


def _memo():
    df = pd.DataFrame([{
        "Lender": "Test Bank", "SBA Loan ID": GP, "Lender Loan ID": GP,
        "Borrower Name": "T", "Securitization Action": "Deferred",
        "Servicer Action": "3-month full principal and interest deferment",
        "Deferral Start Date": pd.Timestamp("2026-04-05"),
        "Deferral End Date": pd.Timestamp("2026-06-05"),
        "Regular Payment Resume Date": pd.Timestamp("2026-07-05"),
        "Report Date": pd.Timestamp("2026-06-01"),
    }])
    df["gp_sba"] = df["SBA Loan ID"].apply(file_readers.normalize_memo_loan_id)
    df["gp_lender"] = df["Lender Loan ID"].apply(file_readers.normalize_memo_loan_id)
    df["deferral_type"] = df["Servicer Action"].apply(file_readers.classify_deferral_type)
    return df


class _Scenario(unittest.TestCase):
    """Shared setup: STAT-01 plus all three summaries, with and without approval."""

    @classmethod
    def setUpClass(cls):
        cls.si = _servicer_inputs()
        cls.stat01 = checks.status_match_check(
            cls.si, _soup(), memo_df=_memo(), servicer_name="Test Bank")
        cls.cd = next(c for c in cls.stat01["correction_data"] if c["gp"] == GP)
        cls.record = next(r for r in cls.stat01["records"]
                          if str(r.get("SBA GP Number")).strip() == GP)
        cls.audited_map = cls.stat01["audited_status_by_loan"]
        cls.recs, cls.cands, _ = corrections.build_recommendations(cls.stat01, "Test Bank")
        cls.approved = {(c["gp"], c["field"]): c["corrected_value"] for c in cls.cands}

        cls.original = checks.portfolio_summary_check(cls.si, "Test Bank")["records"][0]
        cls.audited = checks.audited_portfolio_summary_record(
            cls.si, cls.audited_map, "Test Bank")
        cls.final_unapproved = checks.verified_portfolio_summary_views(
            cls.si, {}, "Test Bank")["dollars"]
        cls.final_approved = checks.verified_portfolio_summary_views(
            cls.si, cls.approved, "Test Bank")["dollars"]


class TheFixtureReproducesTheCondition(_Scenario):
    """Guards the premise. If STAT-01 stops deriving ACTIVE here, the
    assertions below would pass vacuously."""

    def test_servicer_reports_liquidation_and_60_days(self):
        self.assertEqual(self.cd["srf_status"], "LIQUIDATION")
        self.assertEqual(self.cd["srf_dq"], "60 DAYS")

    def test_the_audit_derives_active_and_delinquent(self):
        self.assertEqual(self.cd["expected_loan_status_effective"], "ACTIVE")
        self.assertEqual(self.cd["calculated_dq"], "DELINQUENT")

    def test_it_is_a_confirmed_documented_resume(self):
        self.assertEqual(self.cd["deferral_phase"], "resume")
        self.assertEqual(self.cd["evidence_status"], "documented")

    def test_stat01_raises_a_loan_status_mismatch(self):
        self.assertIn("LOAN STATUS MISMATCH", self.cd["reasons"])
        self.assertEqual(self.record["Loan Status Match"], "NO")


class BeforeReviewerApproval(_Scenario):
    """The Audited summary is generated during /api/audit, before any reviewer
    decision exists. It must therefore keep the LENDER's Loan Status."""

    def test_audited_uses_the_lender_loan_status_for_liquidation(self):
        self.assertEqual(self.audited["Liquidation ($)"], BALANCE)
        self.assertEqual(self.original["Liquidation ($)"], BALANCE)

    def test_the_loan_remains_excluded_from_outstanding(self):
        # Outstanding excludes PAID-IN-FULL and LIQUIDATION. The audit's
        # ACTIVE has not been approved, so the loan stays out.
        self.assertEqual(self.audited["Outstanding ($)"], OTHER_BALANCE)
        self.assertNotIn(BALANCE, (self.audited["Outstanding ($)"],))

    def test_audited_uses_the_calculated_delinquency_for_aging(self):
        self.assertEqual(self.audited["Delinquent ($)"], BALANCE)
        self.assertEqual(self.audited["60 Days ($)"], 0.0)

    def test_the_original_summary_uses_the_lender_aging(self):
        self.assertEqual(self.original["60 Days ($)"], BALANCE)
        self.assertEqual(self.original["Delinquent ($)"], 0.0)

    def test_aging_is_the_only_difference_between_original_and_audited(self):
        # The client-confirmed scope: "identical in every respect EXCEPT the
        # aging-bucket source".
        for category in ("Outstanding", "Not Making Payment", "Deferred",
                         "Bankruptcy", "Foreclosure", "Liquidation", "Paid In Full"):
            with self.subTest(category=category):
                self.assertEqual(self.original[f"{category} ($)"],
                                 self.audited[f"{category} ($)"])

    def test_audited_status_by_loan_carries_delinquency_status_only(self):
        # There is no audited LOAN STATUS anywhere in the map -- which is why
        # the Audited summary cannot act on one.
        self.assertEqual(self.audited_map[GP], "DELINQUENT")
        loan_status_values = {"ACTIVE", "DEFERRED", "LIQUIDATION", "PAID-IN-FULL"}
        self.assertEqual(
            {v for v in self.audited_map.values() if str(v).upper() in loan_status_values},
            set())

    def test_the_loan_status_recommendation_is_active_and_needs_approval(self):
        loan_status = [r for r in self.recs if r["field"] == "Loan Status"]
        self.assertEqual(len(loan_status), 1)
        self.assertEqual(loan_status[0]["reported"], "LIQUIDATION")
        self.assertEqual(loan_status[0]["calculated"], "ACTIVE")
        self.assertTrue(loan_status[0]["reviewFlag"])
        candidate = next(c for c in self.cands if c["field"] == "Loan Status")
        self.assertTrue(candidate["review_required"])

    def test_an_undecided_recommendation_moves_nothing(self):
        # The Final Audited summary with NO decisions must match the lender's
        # Loan Status treatment exactly.
        self.assertEqual(self.final_unapproved["Liquidation ($)"], BALANCE)
        self.assertEqual(self.final_unapproved["Outstanding ($)"], OTHER_BALANCE)


class AfterReviewerApproval(_Scenario):
    """Only the Final Audited summary reflects an approved correction."""

    def test_the_approved_loan_status_is_active(self):
        self.assertEqual(self.approved[(GP, "Loan Status")], "ACTIVE")

    def test_the_loan_enters_outstanding(self):
        self.assertEqual(self.final_approved["Outstanding ($)"],
                         round(BALANCE + OTHER_BALANCE, 2))

    def test_it_leaves_liquidation(self):
        self.assertEqual(self.final_approved["Liquidation ($)"], 0.0)

    def test_the_approved_delinquency_drives_aging(self):
        self.assertEqual(self.final_approved["Delinquent ($)"], BALANCE)
        self.assertEqual(self.final_approved["60 Days ($)"], 0.0)

    def test_the_count_view_moves_with_the_dollars(self):
        counts = checks.verified_portfolio_summary_views(
            self.si, self.approved, "Test Bank")["counts"]
        self.assertEqual(counts["Outstanding"], 2)
        self.assertEqual(counts["Liquidation"], 0)
        self.assertEqual(counts["Delinquent"], 1)

    def test_the_three_layers_are_genuinely_distinct(self):
        # The whole point: three different answers to three different questions.
        self.assertEqual(self.original["Outstanding ($)"], OTHER_BALANCE)
        self.assertEqual(self.audited["Outstanding ($)"], OTHER_BALANCE)
        self.assertEqual(self.final_approved["Outstanding ($)"],
                         round(BALANCE + OTHER_BALANCE, 2))
        self.assertNotEqual(self.original["Delinquent ($)"], self.audited["Delinquent ($)"])


MEMO_PATH = os.environ.get("FALCONBRIDGE_MEMO_XLSX")
JULY_SOUP = os.environ.get("FALCONBRIDGE_JULY_SOUP_XLSX")
JULY_SRF = os.environ.get("FALCONBRIDGE_JULY_CBT_SRF_XLSX")
REAL = all([MEMO_PATH, JULY_SOUP, JULY_SRF])


@unittest.skipUnless(REAL, "set FALCONBRIDGE_MEMO_XLSX / _JULY_SOUP_XLSX / _JULY_CBT_SRF_XLSX")
class RealLoanGP6717929105(unittest.TestCase):
    """The actual loan the boundary was discovered on."""

    GP = "6717929105"
    BALANCE = 843761.03

    @classmethod
    def setUpClass(cls):
        soup = file_readers.read_soup_file(open(JULY_SOUP, "rb").read())
        memo = file_readers.resolve_memo_attribution(
            file_readers.read_service_memo(open(MEMO_PATH, "rb").read()),
            file_readers.build_memo_loan_index(soup))
        sk, dl, lt, si, deal = file_readers.read_servicer_file(
            open(JULY_SRF, "rb").read(),
            filename="Community Bank & Trust||Community Bank and Trust__x.xlsx")
        cls.si = si
        results = checks.run_servicer_reconciliation(
            dl, lt, si, file_readers.filter_soup_by_servicer(soup, sk),
            memo_df=memo, memo_lender_key=sk, deal=deal)
        cls.stat01 = next(r for r in results if r["id"].split("::")[0] == "STAT-01")
        cls.cd = next(c for c in cls.stat01["correction_data"]
                      if str(c.get("gp")).strip() == cls.GP)
        cls.excluded = [t["SBA GP Number"]
                        for t in cls.stat01.get("suppressed_closed_loans") or []]
        cls.audited = checks.audited_portfolio_summary_record(
            si, cls.stat01["audited_status_by_loan"], "CBT", excluded_gps=cls.excluded)
        _, cands, _ = corrections.build_recommendations(cls.stat01, "CBT")
        cls.approved = {(c["gp"], c["field"]): c["corrected_value"]
                        for c in cands if str(c.get("gp")).strip() == cls.GP}
        cls.final_approved = checks.verified_portfolio_summary_views(
            si, cls.approved, "CBT", excluded_gps=cls.excluded)["dollars"]
        cls.final_unapproved = checks.verified_portfolio_summary_views(
            si, {}, "CBT", excluded_gps=cls.excluded)["dollars"]

    def test_the_audit_derives_active_against_a_reported_liquidation(self):
        self.assertEqual(self.cd["srf_status"], "LIQUIDATION")
        self.assertEqual(self.cd["expected_loan_status_effective"], "ACTIVE")
        self.assertIn("LOAN STATUS MISMATCH", self.cd["reasons"])

    def test_before_approval_the_balance_sits_in_liquidation(self):
        self.assertAlmostEqual(self.audited["Liquidation ($)"], 895456.32, places=2)
        self.assertAlmostEqual(self.final_unapproved["Liquidation ($)"], 895456.32, places=2)

    def test_approval_moves_exactly_this_loans_balance_into_outstanding(self):
        delta = self.final_approved["Outstanding ($)"] - self.final_unapproved["Outstanding ($)"]
        self.assertAlmostEqual(delta, self.BALANCE, places=2)

    def test_approval_removes_exactly_this_loans_balance_from_liquidation(self):
        delta = self.final_unapproved["Liquidation ($)"] - self.final_approved["Liquidation ($)"]
        self.assertAlmostEqual(delta, self.BALANCE, places=2)


if __name__ == "__main__":
    unittest.main()
