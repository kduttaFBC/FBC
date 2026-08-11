"""CONT-01 — Prior Period Loan Continuity (client-confirmed 2026-08-11).

A loan that carried an Ending Actual Balance above BALANCE_TOLERANCE in the
immediately prior finalized reporting month is expected to appear in the
current cycle's Servicer Inputs. SOUP is the authoritative prior finalized
period; the current submission's "Loan Terms & Prior Period" sheet is
corroborating reviewer evidence only, never an input to the test.

CONT-01 is deliberately narrow. It answers ONLY "did the loan vanish". Whether
a loan that IS present opens at the right balance is BAL-01's question, and
BAL-01 is not modified by this work. POP-01 is not modified either -- it
compares the Loan Terms population against SOUP, and both sources can agree a
loan exists while the loan is missing from the postings that report on it,
which is exactly the real July 2026 Lendistry 2024 case.

Run with:
    python -m pytest test_prior_period_continuity.py
"""

import copy
import os
import unittest

import pandas as pd

import checks
import file_readers
import main


PRIOR_MONTH = "2026-06-01"


def _soup(rows):
    """Rows are (gp, month, status, dq, ending_balance) with an optional 6th
    Deal element; omitting it puts every row in one partition."""
    return pd.DataFrame([{
        "SBA GP Number": row[0], "Servicer": "T", "Payment Month": row[1],
        "Loan Status": row[2], "Delinquency Status": row[3],
        "Beginning Actual Balance": row[4], "Ending Actual Balance": row[4],
        "Deal": row[5] if len(row) > 5 else "",
    } for row in rows])


def _si(gps):
    """Current-cycle Servicer Inputs. Each entry is a GP, or a
    (gp, beginning_balance, ending_balance) tuple. Repeating a GP produces
    multiple postings for that loan in one cycle."""
    rows = []
    for entry in gps:
        gp, beginning, ending = entry if isinstance(entry, tuple) else (entry, 1000.0, 900.0)
        rows.append({
            "SBA GP Number": gp, "Lender Loan Number": f"L-{gp}",
            "Loan Status": "ACTIVE", "Delinquency Status": "CURRENT",
            "Beginning Balance": beginning, "Ending Balance": ending,
            "Payment Received": 100.0, "Scheduled Payment": 100.0,
            "Servicing Fee": 0.0,
        })
    return pd.DataFrame(rows)


def _lt(rows):
    """Loan Terms & Prior Period. Rows are (gp, prior_ending_balance)."""
    return pd.DataFrame([{
        "SBA GP Number": gp, "Lender Name": "T", "Loan Status": "ACTIVE",
        "Delinquency Status": "CURRENT", "Ending Balance": bal,
    } for gp, bal in rows])


def _run(soup_df, si_df, **kwargs):
    kwargs.setdefault("servicer_name", "T")
    return checks.prior_period_continuity_check(si_df, soup_df, **kwargs)


def _gps(result):
    return [r["SBA GP Number"] for r in result["records"]]


# ---------------------------------------------------------------------------
# Core detection
# ---------------------------------------------------------------------------

class CoreDetectionTests(unittest.TestCase):

    def test_prior_balance_above_tolerance_and_absent_is_flagged(self):
        soup = _soup([("1", PRIOR_MONTH, "BANKRUPTCY", "CURRENT", 12675.95)])
        out = _run(soup, _si(["2"]), deal=None)

        self.assertEqual(out["status"], "FAIL")
        self.assertEqual(out["issues_found"], 1)
        record = out["records"][0]
        self.assertEqual(record["SBA GP Number"], "1")
        self.assertEqual(record["Prior Ending Balance"], 12675.95)
        self.assertEqual(record["Prior Loan Status"], "BANKRUPTCY")
        self.assertEqual(record["Prior Delinquency Status"], "CURRENT")
        self.assertEqual(record["Prior Reporting Month"], "06/2026")
        self.assertEqual(record["Reported This Cycle"], "NO")
        self.assertIn("$12,675.95", record["Issue"])

    def test_loan_present_this_cycle_is_never_flagged(self):
        """Even with a blatant balance break -- that is BAL-01's finding, and
        reporting it here too would double-count a single real problem."""
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 10000.00)])
        out = _run(soup, _si([("1", 1.00, 0.00)]), deal=None)

        self.assertEqual(out["status"], "PASS")
        self.assertEqual(out["issues_found"], 0)

    def test_prior_balance_exactly_zero_is_a_legitimate_exit(self):
        soup = _soup([("1", PRIOR_MONTH, "PAID-IN-FULL", "CURRENT", 0.00)])
        out = _run(soup, _si(["2"]), deal=None)

        self.assertEqual(out["status"], "PASS")

    def test_floating_point_residue_is_not_a_live_balance(self):
        """Regression-locks the two real false positives found in June 2026
        data: Lendistry 5214559105 (2.04e-12) and Poppy 6625075009 (0.0007),
        both PAID-IN-FULL loans whose stale rows carry arithmetic residue."""
        soup = _soup([
            ("5214559105", PRIOR_MONTH, "PAID-IN-FULL", "CURRENT", 2.037268e-12),
            ("6625075009", PRIOR_MONTH, "PAID-IN-FULL", "CURRENT", 0.0007),
        ])
        out = _run(soup, _si(["other"]), deal=None)

        self.assertEqual(out["status"], "PASS")
        self.assertEqual(out["issues_found"], 0)

    def test_balance_tolerance_boundary(self):
        """Zero is judged with BALANCE_TOLERANCE, the same as every other
        balance comparison in the module -- not a strict `> 0`."""
        self.assertEqual(checks.BALANCE_TOLERANCE, 1.00)

        inside = _run(_soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 0.99)]),
                      _si(["2"]), deal=None)
        self.assertEqual(inside["status"], "PASS")

        at_boundary = _run(_soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 1.00)]),
                           _si(["2"]), deal=None)
        self.assertEqual(at_boundary["status"], "PASS")

        outside = _run(_soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 1.01)]),
                       _si(["2"]), deal=None)
        self.assertEqual(outside["status"], "FAIL")

    def test_negative_prior_balance_is_not_above_tolerance(self):
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", -500.00)])
        self.assertEqual(_run(soup, _si(["2"]), deal=None)["status"], "PASS")


# ---------------------------------------------------------------------------
# Legitimate exits -- and, just as important, what is NOT one
# ---------------------------------------------------------------------------

class ExitHandlingTests(unittest.TestCase):

    def test_paid_in_full_at_zero_is_a_legitimate_exit(self):
        soup = _soup([("1", PRIOR_MONTH, "PAID-IN-FULL", "CURRENT", 0.00)])
        self.assertEqual(_run(soup, _si(["2"]), deal=None)["status"], "PASS")

    def test_liquidation_at_zero_is_a_legitimate_exit(self):
        soup = _soup([("1", PRIOR_MONTH, "LIQUIDATION", "CURRENT", 0.00)])
        self.assertEqual(_run(soup, _si(["2"]), deal=None)["status"], "PASS")

    def test_substitution_out_is_an_exit_even_with_a_live_balance(self):
        """Reuses the existing, already-validated
        _POPULATION_EXIT_STATUS_KEYWORDS rather than a second list."""
        soup = _soup([("1", PRIOR_MONTH, "SUBSTITUTION-OUT", "CURRENT", 50000.00)])
        self.assertEqual(_run(soup, _si(["2"]), deal=None)["status"], "PASS")

    def test_terminal_status_with_a_live_balance_is_still_flagged(self):
        """Guards against anyone widening the exit list to
        _TERMINAL_STATUS_KEYWORDS. A PAID-IN-FULL or LIQUIDATION loan still
        carrying a real balance is precisely a loan that must keep reporting
        until it closes to zero -- suppressing it would silence the finding
        this check exists to make."""
        soup = _soup([
            ("1", PRIOR_MONTH, "PAID-IN-FULL", "CURRENT", 9000.00),
            ("2", PRIOR_MONTH, "LIQUIDATION", "CURRENT", 7000.00),
        ])
        out = _run(soup, _si(["other"]), deal=None)

        self.assertEqual(out["issues_found"], 2)
        self.assertEqual(_gps(out), ["1", "2"])

    def test_bankruptcy_is_never_an_exit(self):
        """The real July 2026 Lendistry 2024 case, unit-sized."""
        soup = _soup([("1", PRIOR_MONTH, "BANKRUPTCY", "CURRENT", 12675.95)])
        self.assertEqual(_run(soup, _si(["2"]), deal=None)["status"], "FAIL")

    def test_foreclosure_and_reo_are_never_exits(self):
        soup = _soup([
            ("1", PRIOR_MONTH, "FORECLOSURE", "90 DAYS", 4000.00),
            ("2", PRIOR_MONTH, "REO", "120 DAYS", 6000.00),
        ])
        out = _run(soup, _si(["other"]), deal=None)

        self.assertEqual(out["issues_found"], 2)
        self.assertEqual(_gps(out), ["1", "2"])

    def test_closed_at_zero_suppresses_and_is_traced(self):
        """The shared closed-at-zero gate applies here too, and suppression
        stays auditable via suppressed_closed_gps -- the same sibling key
        BAL-01/BAL-02 use, which run_servicer_reconciliation folds into
        STAT-01's suppression trace."""
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 8000.00)])
        out = _run(soup, _si(["2"]), deal=None,
                   closed_at_zero={"1": {"closed_status": "PAID-IN-FULL",
                                         "closed_month": "2026-01",
                                         "closed_balance": 0.0,
                                         "closed_deal": ""}})

        self.assertEqual(out["status"], "PASS")
        self.assertEqual(out["issues_found"], 0)
        self.assertEqual(out["suppressed_closed_gps"], ["1"])

    def test_closed_at_zero_trace_only_lists_loans_that_would_have_flagged(self):
        """Applied last rather than first (unlike BAL-01/BAL-02, which iterate
        the current population): a trace entry here means the loan really
        would have been a CONT-01 finding, not merely that it was closed."""
        soup = _soup([
            ("reported", PRIOR_MONTH, "ACTIVE", "CURRENT", 8000.00),
            ("zero-balance", PRIOR_MONTH, "PAID-IN-FULL", "CURRENT", 0.00),
            ("would-flag", PRIOR_MONTH, "ACTIVE", "CURRENT", 8000.00),
        ])
        closed = {gp: {"closed_status": "PAID-IN-FULL", "closed_month": "2026-01",
                       "closed_balance": 0.0, "closed_deal": ""}
                  for gp in ("reported", "zero-balance", "would-flag")}
        out = _run(soup, _si(["reported"]), deal=None, closed_at_zero=closed)

        self.assertEqual(out["suppressed_closed_gps"], ["would-flag"])


# ---------------------------------------------------------------------------
# Scoping, identity and prior-month anchoring
# ---------------------------------------------------------------------------

class ScopingAndIdentityTests(unittest.TestCase):

    def test_deal_scoping_isolates_note_series(self):
        """SOUP files every note series under one Servicer value. Without
        Deal scoping, the 2025 file would be compared against the 2024 deal's
        loans too -- the same trap POP-01 documents (1,503 false exceptions
        for Lendistry)."""
        soup = _soup([
            ("shared", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00, "SOUP2024_1"),
            ("shared", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00, "SOUP2025_1"),
            ("only-2024", PRIOR_MONTH, "ACTIVE", "CURRENT", 3000.00, "SOUP2024_1"),
        ])
        si_2025 = _si(["shared"])

        scoped = _run(soup, si_2025, deal="SOUP2025_1")
        self.assertEqual(scoped["status"], "PASS")

        # The same submission, compared without Deal scoping, wrongly reports
        # the other note series' loan as missing.
        unscoped = _run(soup, si_2025, deal=None)
        self.assertEqual(_gps(unscoped), ["only-2024"])

    def test_deal_scoping_still_flags_within_its_own_deal(self):
        soup = _soup([
            ("shared", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00, "SOUP2024_1"),
            ("shared", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00, "SOUP2025_1"),
        ])
        out = _run(soup, _si(["other"]), deal="SOUP2024_1")

        self.assertEqual(_gps(out), ["shared"])
        self.assertEqual(out["records"][0]["Deal"], "SOUP2024_1")
        self.assertIn("scoped to Deal 'SOUP2024_1'", out["details"])

    def test_undeterminable_deal_falls_back_with_an_explicit_note(self):
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00, "SOUP2024_1")])
        out = _run(soup, _si(["2"]), deal=None)

        self.assertEqual(out["issues_found"], 1)
        self.assertEqual(out["records"][0]["Deal"], "unknown")
        self.assertIn("Deal could not be determined", out["details"])

    def test_gp_number_normalization(self):
        """'5561999102.0' off a float-typed cell must match '5561999102'."""
        soup = _soup([("5561999102", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00)])
        out = _run(soup, _si(["5561999102.0"]), deal=None)

        self.assertEqual(out["status"], "PASS")

    def test_prior_month_anchoring_ignores_stale_history(self):
        """The prior population is SOUP's LATEST month, not the latest row per
        loan. A loan whose history ended months ago is PAY-01's stopped-
        reporting case; anchoring per loan would re-flag it here forever."""
        soup = _soup([
            ("stale", "2026-03-01", "ACTIVE", "CURRENT", 5000.00),
            ("current", PRIOR_MONTH, "ACTIVE", "CURRENT", 4000.00),
        ])
        out = _run(soup, _si(["current"]), deal=None)

        self.assertEqual(out["status"], "PASS")

    def test_prior_month_anchoring_uses_only_the_latest_month_row(self):
        """A loan present in earlier months with a live balance but closed to
        zero in the prior month is a legitimate exit, judged on the prior
        month's row alone."""
        soup = _soup([
            ("1", "2026-05-01", "ACTIVE", "CURRENT", 5000.00),
            ("1", PRIOR_MONTH, "PAID-IN-FULL", "CURRENT", 0.00),
        ])
        self.assertEqual(_run(soup, _si(["2"]), deal=None)["status"], "PASS")

    def test_multiple_postings_this_cycle_count_as_reported(self):
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00)])
        si = _si([("1", 5000.0, 4800.0), ("1", 4800.0, 4600.0), ("1", 4600.0, 4400.0)])
        self.assertEqual(len(si), 3)

        self.assertEqual(_run(soup, si, deal=None)["status"], "PASS")


# ---------------------------------------------------------------------------
# Corroborating evidence from the current submission's own prior-period sheet
# ---------------------------------------------------------------------------

class CorroboratingEvidenceTests(unittest.TestCase):

    def test_loan_terms_prior_balance_is_carried_as_evidence(self):
        soup = _soup([("1", PRIOR_MONTH, "BANKRUPTCY", "CURRENT", 12675.95)])
        out = _run(soup, _si(["2"]), loan_terms_df=_lt([("1", 12675.95)]), deal=None)

        self.assertEqual(out["records"][0]["Prior Period Sheet Balance"], 12675.95)

    def test_missing_loan_terms_evidence_leaves_the_field_blank(self):
        soup = _soup([("1", PRIOR_MONTH, "BANKRUPTCY", "CURRENT", 12675.95)])
        out = _run(soup, _si(["2"]), loan_terms_df=_lt([("other", 1.0)]), deal=None)

        self.assertEqual(out["records"][0]["Prior Period Sheet Balance"], "")

    def test_loan_terms_evidence_never_changes_the_finding(self):
        """Evidence for a human, never an input to the test -- a servicer that
        drops a loan from Servicer Inputs could equally drop it from its own
        prior-period sheet, so SOUP alone decides."""
        soup = _soup([("1", PRIOR_MONTH, "BANKRUPTCY", "CURRENT", 12675.95)])

        without = _run(soup, _si(["2"]), deal=None)
        contradicting = _run(soup, _si(["2"]), loan_terms_df=_lt([("1", 0.00)]), deal=None)
        absent = _run(soup, _si(["2"]), loan_terms_df=pd.DataFrame(), deal=None)

        for out in (without, contradicting, absent):
            self.assertEqual(out["issues_found"], 1)
            self.assertEqual(out["records"][0]["Prior Ending Balance"], 12675.95)


# ---------------------------------------------------------------------------
# Result shape and safe degradation
# ---------------------------------------------------------------------------

class ResultShapeTests(unittest.TestCase):

    def test_result_dict_shape(self):
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00)])
        out = _run(soup, _si(["2"]), deal=None)

        self.assertEqual(out["id"], "CONT-01")
        self.assertEqual(out["name"], "Prior Period Loan Continuity")
        self.assertEqual(out["category"], "SBA Reconciliation")
        self.assertEqual(out["severity"], "FAIL")
        self.assertEqual(out["status"], "FAIL")
        self.assertEqual(out["issues_found"], len(out["records"]))
        self.assertIn("suppressed_closed_gps", out)

    def test_record_field_order_is_stable(self):
        soup = _soup([("1", PRIOR_MONTH, "BANKRUPTCY", "CURRENT", 12675.95, "SOUP2024_1")])
        out = _run(soup, _si(["2"]), deal="SOUP2024_1")

        self.assertEqual(list(out["records"][0].keys()), [
            "SBA GP Number", "Deal", "Prior Reporting Month", "Prior Loan Status",
            "Prior Delinquency Status", "Prior Ending Balance", "Reported This Cycle",
            "Issue", "Prior Period Sheet Balance",
        ])

    def test_records_are_sorted_by_gp(self):
        soup = _soup([
            ("3", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00),
            ("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00),
            ("2", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00),
        ])
        self.assertEqual(_gps(_run(soup, _si(["other"]), deal=None)), ["1", "2", "3"])

    def test_clean_result_says_no_issues_found(self):
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00)])
        out = _run(soup, _si(["1"]), deal=None)

        self.assertEqual(out["details"], "No issues found.")
        self.assertEqual(out["records"], [])


class SafeDegradationTests(unittest.TestCase):
    """Every source this check can't evaluate degrades to "nothing to
    report" rather than raising, as the rest of the module does."""

    def test_empty_soup(self):
        self.assertEqual(_run(pd.DataFrame(), _si(["1"]), deal=None)["issues_found"], 0)

    def test_soup_missing_payment_month(self):
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00)]).drop(columns=["Payment Month"])
        self.assertEqual(_run(soup, _si(["2"]), deal=None)["issues_found"], 0)

    def test_soup_missing_ending_actual_balance(self):
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00)]).drop(columns=["Ending Actual Balance"])
        self.assertEqual(_run(soup, _si(["2"]), deal=None)["issues_found"], 0)

    def test_soup_with_unparseable_payment_months(self):
        soup = _soup([("1", "not a date", "ACTIVE", "CURRENT", 5000.00)])
        self.assertEqual(_run(soup, _si(["2"]), deal=None)["issues_found"], 0)

    def test_null_prior_balance_cannot_be_evaluated(self):
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", None)])
        self.assertEqual(_run(soup, _si(["2"]), deal=None)["issues_found"], 0)

    def test_empty_servicer_inputs(self):
        """A submission with no readable postings is a file-level failure the
        readers and preflight already surface. Reporting every loan in the
        portfolio as missing would bury that, not reveal it."""
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00)])
        self.assertEqual(_run(soup, pd.DataFrame(), deal=None)["issues_found"], 0)

    def test_servicer_inputs_missing_gp_column(self):
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00)])
        si = _si(["1"]).drop(columns=["SBA GP Number"])
        self.assertEqual(_run(soup, si, deal=None)["issues_found"], 0)

    def test_deal_scoping_that_matches_nothing(self):
        soup = _soup([("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00, "SOUP2024_1")])
        self.assertEqual(_run(soup, _si(["2"]), deal="SOUP2099_1")["issues_found"], 0)


# ---------------------------------------------------------------------------
# Registration, wiring, and non-interference with existing checks
# ---------------------------------------------------------------------------

def _full_soup():
    return _soup([
        ("1", "2026-05-01", "ACTIVE", "CURRENT", 5100.00, "SOUP2024_1"),
        ("1", PRIOR_MONTH, "ACTIVE", "CURRENT", 5000.00, "SOUP2024_1"),
        ("2", "2026-05-01", "ACTIVE", "CURRENT", 4100.00, "SOUP2024_1"),
        ("2", PRIOR_MONTH, "BANKRUPTCY", "CURRENT", 4000.00, "SOUP2024_1"),
    ])


class RunnerIntegrationTests(unittest.TestCase):

    def test_registered_in_all_checks_meta(self):
        self.assertIn(("CONT-01", "Prior Period Loan Continuity"), checks.ALL_CHECKS_META)
        self.assertEqual(len(checks.ALL_CHECKS_META), 12)
        self.assertEqual(len({c[0] for c in checks.ALL_CHECKS_META}), 12)

    def test_runner_tags_the_result_with_the_display_label(self):
        results = checks.run_servicer_reconciliation(
            "Lendistry 2024", _lt([("1", 5000.00), ("2", 4000.00)]),
            _si([("1", 5000.0, 4900.0)]), _full_soup(), deal="SOUP2024_1",
        )
        cont = [r for r in results if r["id"].split("::")[0] == "CONT-01"]

        self.assertEqual(len(cont), 1)
        self.assertEqual(cont[0]["id"], "CONT-01::Lendistry 2024")
        self.assertEqual(cont[0]["servicer"], "Lendistry 2024")
        self.assertEqual(_gps(cont[0]), ["2"])
        self.assertEqual(cont[0]["records"][0]["Prior Loan Status"], "BANKRUPTCY")
        self.assertEqual(cont[0]["records"][0]["Prior Period Sheet Balance"], 4000.00)

    def test_runner_strips_the_suppression_sibling_key(self):
        results = checks.run_servicer_reconciliation(
            "T", _lt([("1", 5000.00)]), _si(["1", "2"]), _full_soup(), deal="SOUP2024_1",
        )
        cont = next(r for r in results if r["id"].split("::")[0] == "CONT-01")

        self.assertNotIn("suppressed_closed_gps", cont)

    def test_does_not_change_any_existing_check_result(self):
        """The guarantee behind "do not modify the existing population check":
        every other check must return exactly what it returned before CONT-01
        existed -- which also proves CONT-01 does not mutate the shared
        SOUP/Servicer Inputs frames its neighbours read."""
        args = ("T", _lt([("1", 5000.00), ("2", 4000.00)]),
                _si([("1", 5000.0, 4900.0)]), _full_soup())
        kwargs = {"deal": "SOUP2024_1"}

        with_cont = checks.run_servicer_reconciliation(*copy.deepcopy(args), **kwargs)

        original = checks.prior_period_continuity_check
        checks.prior_period_continuity_check = lambda *a, **k: {
            "id": "CONT-STUB", "name": "stub", "category": "SBA Reconciliation",
            "severity": "FAIL", "status": "PASS", "issues_found": 0,
            "details": "No issues found.", "records": [],
        }
        try:
            without_cont = checks.run_servicer_reconciliation(*copy.deepcopy(args), **kwargs)
        finally:
            checks.prior_period_continuity_check = original

        def _others(results):
            return [r for r in results if r["id"].split("::")[0] not in ("CONT-01", "CONT-STUB")]

        self.assertEqual(len(_others(with_cont)), 11)
        self.assertEqual(_others(with_cont), _others(without_cont))

    def test_population_check_still_passes_a_loan_missing_only_from_postings(self):
        """POP-01 is untouched and must stay untouched: it compares the Loan
        Terms population against SOUP, and a loan present in both while absent
        from Servicer Inputs is legitimately not its finding. That gap is
        precisely what CONT-01 exists to close."""
        results = checks.run_servicer_reconciliation(
            "T", _lt([("1", 5000.00), ("2", 4000.00)]),
            _si([("1", 5000.0, 4900.0)]), _full_soup(), deal="SOUP2024_1",
        )
        pop = next(r for r in results if r["id"].split("::")[0] == "POP-01")
        cont = next(r for r in results if r["id"].split("::")[0] == "CONT-01")

        self.assertEqual(pop["status"], "PASS")
        self.assertEqual(pop["issues_found"], 0)
        self.assertEqual(cont["issues_found"], 1)


class ReportTierTests(unittest.TestCase):

    def test_cont01_is_a_conditional_report_tab(self):
        self.assertIn("CONT-01", main.REPORT_CONDITIONAL_CHECK_IDS)
        self.assertNotIn("CONT-01", main.REPORT_ALWAYS_EXCLUDED_CHECK_IDS)

    def test_no_tab_when_clean_and_a_tab_once_it_finds_something(self):
        clean = {"id": "CONT-01::T", "issues_found": 0, "records": []}
        found = {"id": "CONT-01::T", "issues_found": 1, "records": [{"SBA GP Number": "1"}]}

        self.assertEqual(main._filter_report_results([clean]), [])
        self.assertEqual(main._filter_report_results([found]), [found])


# ---------------------------------------------------------------------------
# Real fixture — July 2026 Lendistry 2024
# ---------------------------------------------------------------------------

JUNE_FINAL = os.path.expanduser(
    "~/Downloads/2026.06/Final/2026.06_Lendistry_SOUP 2024-1 SRF Template_June 2026 - vF_2026.07.16.xlsx"
)
JULY_INITIAL = os.path.expanduser(
    "~/Downloads/2026.07_Lendistry_SOUP 2024-1 SRF Template_July 2026.xlsx"
)
FIXTURES_PRESENT = os.path.exists(JUNE_FINAL) and os.path.exists(JULY_INITIAL)


@unittest.skipUnless(FIXTURES_PRESENT, "Lendistry 2024 June/July workbooks not present")
class RealFixtureTests(unittest.TestCase):
    """GP 5561999102 carried $12,675.95 at BANKRUPTCY in the June 2026 final
    report and does not appear in the July 2026 submission at all. It is the
    ONLY loan in the deal that qualifies -- the two other June dropouts
    (4709119110, 5499889102) are PAID-IN-FULL at $0.00.

    Neither workbook is in the repo, so this skips when they are absent. SOUP
    for the July cycle would carry June as its latest month; here that frame is
    reconstructed from June's own finalized Servicer Inputs, which is what SOUP
    is built from -- validated as an exact mirror (SOUP 2026-05 vs the June
    file's prior-period sheet: 122 loans in common, 0 mismatches)."""

    @classmethod
    def setUpClass(cls):
        _, _, _, june_si, cls.deal = file_readers.read_servicer_file(
            open(JUNE_FINAL, "rb").read(), os.path.basename(JUNE_FINAL))
        _, _, cls.july_lt, cls.july_si, _ = file_readers.read_servicer_file(
            open(JULY_INITIAL, "rb").read(), os.path.basename(JULY_INITIAL))

        consolidated = checks.consolidate_cycle_postings(june_si)
        cls.prior_soup = pd.DataFrame([{
            "SBA GP Number": gp,
            "Servicer": "Lendistry",
            "Deal": "SOUP2024_1",
            "Payment Month": "2026-06-01",
            "Loan Status": row.get("loan_status"),
            "Delinquency Status": row.get("dq_status"),
            "Beginning Actual Balance": row.get("beginning_balance"),
            "Ending Actual Balance": row.get("ending_balance"),
        } for gp, row in consolidated.iterrows()])

    def _result(self):
        return checks.prior_period_continuity_check(
            self.july_si, self.prior_soup, loan_terms_df=self.july_lt,
            servicer_name="Lendistry 2024", deal="SOUP2024_1",
            closed_at_zero=checks.loans_closed_at_zero(self.prior_soup),
        )

    def test_deal_is_read_from_the_workbook(self):
        self.assertEqual(self.deal, "SOUP2024_1")

    def test_5561999102_is_the_single_finding(self):
        out = self._result()

        self.assertEqual(out["status"], "FAIL")
        self.assertEqual(out["issues_found"], 1)
        record = out["records"][0]
        self.assertEqual(record["SBA GP Number"], "5561999102")
        self.assertEqual(record["Prior Ending Balance"], 12675.95)
        self.assertEqual(record["Prior Loan Status"], "BANKRUPTCY")
        self.assertEqual(record["Prior Reporting Month"], "06/2026")
        self.assertEqual(record["Deal"], "SOUP2024_1")
        self.assertEqual(record["Reported This Cycle"], "NO")

    def test_paid_in_full_dropouts_are_not_flagged(self):
        """Both left the portfolio in June at $0.00 -- legitimate exits."""
        out = self._result()

        for gp in ("4709119110", "5499889102"):
            self.assertNotIn(gp, self.july_si["SBA GP Number"].values)
            self.assertNotIn(gp, _gps(out))

    def test_the_submission_itself_corroborates_the_omission(self):
        """July's own "Loan Terms & Prior Period" sheet states the $12,675.95
        closing balance the submission then failed to report on."""
        out = self._result()

        self.assertEqual(out["records"][0]["Prior Period Sheet Balance"], 12675.95)

    def test_details_names_the_servicer_and_deal(self):
        out = self._result()

        self.assertIn("Lendistry 2024", out["details"])
        self.assertIn("scoped to Deal 'SOUP2024_1'", out["details"])


if __name__ == "__main__":
    unittest.main()
