"""Review-evidence amounts must reach the payload as None, never NaN.

THE PRODUCTION DEFECT THIS PINS

    ValueError: Out of range float values are not JSON compliant: nan
      when serializing dict item 'liquidationLoss'
      when serializing list item 0
      when serializing dict item 'recommendations'

POST /api/audit returned HTTP 400. Not a client error: Starlette's JSONResponse
renders with allow_nan=False and __init__ renders EAGERLY, so the ValueError is
raised at `return JSONResponse({...})` inside the handler body and main.py's
outer `except Exception` reports it as 400.

ROOT CAUSE. Nothing calculated a NaN. _cycle_review_evidence deliberately sets
each amount to a literal None when the source column carried nothing, because
the client-confirmed rule is that an all-blank column stays None and must not
become 0.00. `pd.DataFrame(rows).set_index("SBA GP Number")` then infers the
column dtype from the rows handed to it:

    every loan None        -> object  dtype, None survives
    ANY loan a real value  -> float64 dtype, EVERY None becomes NaN

So ONE liquidating loan converts every other loan's "nothing reported" into NaN.

WHY THE SUITE MISSED IT, and what these tests do about it:

  1. No fixture had a populated Liquidation Loss. Measured across all four real
     files available (June/July x CB&T/Lendistry, 2,772 loans): zero real
     values. MixedColumnTests below supplies the missing shape.
  2. The trigger needs a MIXED column -- at least one real value AND at least
     one None -- in ONE servicer file. All-None stays object; single-loan
     all-real has no None to corrupt. Neither reproduces.
  3. No test asserted strict JSON serializability of the payload. In Python a
     NaN is an ordinary float; the failure exists only at the serialization
     boundary. StrictSerializationTests crosses it.

SCOPE. NaN -> None for the three review-evidence amounts only. Infinities are
deliberately out of scope and InfinityIsOutOfScopeTests pins that on purpose, so
a later infinity fix is a visible decision rather than an accident here.
"""

import io
import json
import math
import os
import unittest

import pandas as pd

import checks
import corrections


HEADER_ROW = 9
SERVICER = "Test Bank"
GP_PLAIN = "1000000001"      # no liquidation amounts at all
GP_LIQUIDATING = "2000000002"  # carries a real Liquidation Loss
BAL = 80000.0

AMOUNT_FIELDS = ("liquidation_proceeds", "liquidation_loss", "subsequent_recoveries")
PAYLOAD_FIELDS = ("liquidationProceeds", "liquidationLoss", "subsequentRecoveries")

SOUP_XLSX = os.environ.get("FALCONBRIDGE_JULY_SOUP_XLSX")
MEMO_XLSX = os.environ.get("FALCONBRIDGE_MEMO_XLSX")
CBT_XLSX = os.environ.get("FALCONBRIDGE_JULY_CBT_SRF_XLSX")
REAL = all([SOUP_XLSX, MEMO_XLSX, CBT_XLSX])


def _soup(rows):
    """rows: (gp, month, loan_status, dq, ending_balance)"""
    return pd.DataFrame([{
        "SBA GP Number": gp, "Payment Month": m, "Loan Status": s,
        "Delinquency Status": dq, "Ending Actual Balance": bal,
        "Beginning Actual Balance": bal, "Deal": "SOUP2025_1",
    } for gp, m, s, dq, bal in rows])


def _row(gp, status="ACTIVE", dq="CURRENT", bal=BAL, received=0.0, scheduled=500.0,
         excel_row=HEADER_ROW + 1, instance=1, last="Y", **extra):
    row = {
        "SBA GP Number": gp, "Lender Loan Number": "L-" + gp, "Loan Status": status,
        "Delinquency Status": dq, "Payment Received": received, "Principal": 0.0,
        "Net Interest": 0.0, "Servicing Fee": 0.0, "Scheduled Payment": scheduled,
        "Beginning Balance": bal, "Ending Balance": bal,
        "Loan Payment Instance": instance, "Last Payment Instance?": last,
        "_excel_row": excel_row, "_sheet_name": "Servicer Inputs",
        "_header_excel_row": HEADER_ROW,
        # Present on every row so the column EXISTS; blank for the plain loan.
        "Liquidation Proceeds": None,
        "Liquidation Loss": None,
        "Subsequent Recoveries/(Losses)": None,
    }
    row.update(extra)
    return row


def _run(srf, soup_rows):
    """Full pipeline -> (stat01, recommendations)."""
    results = checks.run_servicer_reconciliation(
        SERVICER, pd.DataFrame(), srf, _soup(soup_rows),
        memo_df=None, memo_lender_key=SERVICER)
    stat01 = next(r for r in results if r["id"].split("::")[0] == "STAT-01")
    recs, _cands, _skipped = corrections.build_recommendations(stat01, SERVICER)
    return stat01, recs


def _mixed_frame(loss=4321.55):
    """The shape that reproduces production: two loans in ONE file, one with a
    real Liquidation Loss and one with nothing. Both need a real audit finding
    so both produce a recommendation to inspect."""
    return pd.DataFrame([
        # DQ mismatch -> a Delinquency Status recommendation for each loan.
        _row(GP_PLAIN, dq="CURRENT", received=0.0, scheduled=500.0,
             excel_row=HEADER_ROW + 1),
        _row(GP_LIQUIDATING, dq="CURRENT", received=0.0, scheduled=500.0,
             excel_row=HEADER_ROW + 2, **{"Liquidation Loss": loss}),
    ])


def _mixed_soup():
    return [(GP_PLAIN, "2026-06-01", "ACTIVE", "CURRENT", BAL),
            (GP_LIQUIDATING, "2026-06-01", "ACTIVE", "CURRENT", BAL)]


def _by_gp(recs):
    out = {}
    for r in recs:
        out.setdefault(str(r["gp"]), []).append(r)
    return out


def _bad_float(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, float):
        if math.isnan(value):
            return "NaN"
        if math.isinf(value):
            return "+inf" if value > 0 else "-inf"
    return None


# ---------------------------------------------------------------------------
# 1 -- the normaliser itself
# ---------------------------------------------------------------------------

class NanToNoneTests(unittest.TestCase):

    def test_nan_becomes_none(self):
        self.assertIsNone(checks._nan_to_none(float("nan")))

    def test_numpy_nan_becomes_none(self):
        import numpy as np
        self.assertIsNone(checks._nan_to_none(np.float64("nan")))

    def test_none_stays_none(self):
        self.assertIsNone(checks._nan_to_none(None))

    def test_a_real_amount_is_untouched(self):
        self.assertEqual(checks._nan_to_none(1234.56), 1234.56)

    def test_zero_is_preserved_and_not_confused_with_missing(self):
        """A reported 0.00 is a REPORTED value and must survive as 0.0, not
        collapse into None -- otherwise the fix would destroy real data."""
        result = checks._nan_to_none(0.0)
        self.assertEqual(result, 0.0)
        self.assertIsNotNone(result)

    def test_a_negative_amount_is_untouched(self):
        self.assertEqual(checks._nan_to_none(-99.01), -99.01)

    def test_a_string_is_untouched(self):
        self.assertEqual(checks._nan_to_none("n/a"), "n/a")

    def test_a_list_is_returned_unchanged_not_raised_on(self):
        self.assertEqual(checks._nan_to_none([1, 2]), [1, 2])


class InfinityIsOutOfScopeTests(unittest.TestCase):
    """Infinities are DELIBERATELY not handled by this change. Pinned so a
    future infinity fix is an explicit decision, not an accident."""

    def test_positive_infinity_passes_through(self):
        self.assertEqual(checks._nan_to_none(float("inf")), float("inf"))

    def test_negative_infinity_passes_through(self):
        self.assertEqual(checks._nan_to_none(float("-inf")), float("-inf"))


# ---------------------------------------------------------------------------
# 2 -- the pandas mechanism, documented so the cause stays visible
# ---------------------------------------------------------------------------

class PandasCoercionMechanismTests(unittest.TestCase):

    def test_one_real_value_turns_every_none_into_nan(self):
        frame = pd.DataFrame([{"gp": "A", "liquidation_loss": None},
                              {"gp": "B", "liquidation_loss": 123.45}])
        self.assertEqual(frame["liquidation_loss"].dtype, "float64")
        value = frame.set_index("gp").loc["A"].get("liquidation_loss")
        self.assertTrue(pd.isna(value))
        self.assertIsNotNone(value)          # NaN, not None

    def test_all_none_keeps_object_dtype_and_none_survives(self):
        frame = pd.DataFrame([{"gp": "A", "liquidation_loss": None},
                              {"gp": "B", "liquidation_loss": None}])
        self.assertEqual(frame["liquidation_loss"].dtype, object)
        self.assertIsNone(frame.set_index("gp").loc["A"].get("liquidation_loss"))

    def test_the_consolidated_frame_really_does_go_float64(self):
        """Not a hypothetical: the real consolidation produces the NaN."""
        grouped = checks.consolidate_cycle_postings(_mixed_frame())
        self.assertEqual(grouped["liquidation_loss"].dtype, "float64")
        self.assertTrue(pd.isna(grouped.loc[GP_PLAIN]["liquidation_loss"]))
        self.assertEqual(grouped.loc[GP_LIQUIDATING]["liquidation_loss"], 4321.55)


# ---------------------------------------------------------------------------
# 3 -- THE CRITICAL CASE: mixed column, end to end
# ---------------------------------------------------------------------------

class MixedColumnTests(unittest.TestCase):
    """The shape no fixture had: one loan reporting an amount, one not."""

    def setUp(self):
        self.stat01, self.recs = _run(_mixed_frame(), _mixed_soup())
        self.by_gp = _by_gp(self.recs)

    def test_both_loans_produced_a_recommendation(self):
        self.assertIn(GP_PLAIN, self.by_gp)
        self.assertIn(GP_LIQUIDATING, self.by_gp)

    def test_the_non_liquidating_loan_reports_none_not_nan(self):
        for rec in self.by_gp[GP_PLAIN]:
            value = rec["liquidationLoss"]
            self.assertIsNone(value, "expected None, got %r" % (value,))

    def test_the_non_liquidating_loan_is_not_zero(self):
        """0 is explicitly WRONG: no loss reported is not a loss of zero."""
        for rec in self.by_gp[GP_PLAIN]:
            self.assertNotEqual(rec["liquidationLoss"], 0)
            self.assertNotEqual(rec["liquidationLoss"], 0.0)

    def test_the_real_value_still_round_trips_unchanged(self):
        for rec in self.by_gp[GP_LIQUIDATING]:
            self.assertEqual(rec["liquidationLoss"], 4321.55)

    def test_no_amount_field_is_nan_on_any_recommendation(self):
        for rec in self.recs:
            for field in PAYLOAD_FIELDS:
                self.assertIsNone(_bad_float(rec[field]),
                                  "%s on GP %s" % (field, rec["gp"]))

    def test_correction_data_itself_is_clean_not_just_the_payload(self):
        """The fix must land in correction_data, which has more than one
        reader -- not only in the recommendation payload."""
        for entry in self.stat01.get("correction_data") or []:
            for field in AMOUNT_FIELDS:
                self.assertIsNone(_bad_float(entry.get(field)),
                                  "%s on GP %s" % (field, entry.get("gp")))


class ProceedsAndRecoveriesTests(unittest.TestCase):
    """Production named liquidationLoss only, which means Proceeds was still
    object dtype while Loss had gone float64. Each column goes mixed on its
    own, so each needs its own test."""

    def _mixed_on(self, column, value):
        srf = pd.DataFrame([
            _row(GP_PLAIN, excel_row=HEADER_ROW + 1),
            _row(GP_LIQUIDATING, excel_row=HEADER_ROW + 2, **{column: value}),
        ])
        _stat01, recs = _run(srf, _mixed_soup())
        return _by_gp(recs)

    def test_proceeds_alone_going_mixed(self):
        by_gp = self._mixed_on("Liquidation Proceeds", 5000.0)
        for rec in by_gp[GP_PLAIN]:
            self.assertIsNone(rec["liquidationProceeds"])
        for rec in by_gp[GP_LIQUIDATING]:
            self.assertEqual(rec["liquidationProceeds"], 5000.0)

    def test_subsequent_recoveries_alone_going_mixed(self):
        by_gp = self._mixed_on("Subsequent Recoveries/(Losses)", -250.75)
        for rec in by_gp[GP_PLAIN]:
            self.assertIsNone(rec["subsequentRecoveries"])
        for rec in by_gp[GP_LIQUIDATING]:
            self.assertEqual(rec["subsequentRecoveries"], -250.75)

    def test_loss_going_mixed_leaves_proceeds_as_none_not_nan(self):
        """Exactly the production shape: a loss with no proceeds anywhere."""
        by_gp = self._mixed_on("Liquidation Loss", 4321.55)
        for gp, recs in by_gp.items():
            for rec in recs:
                self.assertIsNone(rec["liquidationProceeds"], "GP %s" % gp)


class UnchangedWhereTheBugNeverBitTests(unittest.TestCase):

    def test_all_none_still_yields_none(self):
        srf = pd.DataFrame([_row(GP_PLAIN, excel_row=HEADER_ROW + 1),
                            _row(GP_LIQUIDATING, excel_row=HEADER_ROW + 2)])
        _stat01, recs = _run(srf, _mixed_soup())
        self.assertTrue(recs)
        for rec in recs:
            for field in PAYLOAD_FIELDS:
                self.assertIsNone(rec[field])

    def test_a_file_without_the_columns_at_all_still_yields_none(self):
        rows = []
        for gp, excel in ((GP_PLAIN, HEADER_ROW + 1), (GP_LIQUIDATING, HEADER_ROW + 2)):
            row = _row(gp, excel_row=excel)
            for column in ("Liquidation Proceeds", "Liquidation Loss",
                           "Subsequent Recoveries/(Losses)"):
                row.pop(column)
            rows.append(row)
        _stat01, recs = _run(pd.DataFrame(rows), _mixed_soup())
        self.assertTrue(recs)
        for rec in recs:
            for field in PAYLOAD_FIELDS:
                self.assertIsNone(rec[field])

    def test_multi_posting_summation_still_totals_the_cycle(self):
        """min_count=1 summation is unchanged: a loss split across two postings
        reports the cycle total, and the OTHER loan still reports None."""
        srf = pd.DataFrame([
            _row(GP_LIQUIDATING, excel_row=HEADER_ROW + 1, instance=1, last="N",
                 **{"Liquidation Loss": 1000.0}),
            _row(GP_LIQUIDATING, excel_row=HEADER_ROW + 2, instance=2, last="Y",
                 **{"Liquidation Loss": 250.25}),
            _row(GP_PLAIN, excel_row=HEADER_ROW + 3),
        ])
        _stat01, recs = _run(srf, _mixed_soup())
        by_gp = _by_gp(recs)
        for rec in by_gp.get(GP_LIQUIDATING, []):
            self.assertEqual(rec["liquidationLoss"], 1250.25)
        for rec in by_gp.get(GP_PLAIN, []):
            self.assertIsNone(rec["liquidationLoss"])


# ---------------------------------------------------------------------------
# 4 -- the serialization boundary no test crossed
# ---------------------------------------------------------------------------

class StrictSerializationTests(unittest.TestCase):
    """allow_nan=False is what Starlette's JSONResponse.render uses. These
    assertions are the ones that would have caught the production failure."""

    def test_the_mixed_column_payload_serializes_strictly(self):
        _stat01, recs = _run(_mixed_frame(), _mixed_soup())
        json.dumps(recs, allow_nan=False)

    def test_liquidation_loss_serializes_as_json_null_never_zero(self):
        _stat01, recs = _run(_mixed_frame(), _mixed_soup())
        plain = _by_gp(recs)[GP_PLAIN][0]
        # Rendered with Starlette's own separators so this asserts the exact
        # bytes JSONResponse would emit, not json.dumps' default spacing.
        encoded = json.dumps({"liquidationLoss": plain["liquidationLoss"]},
                             allow_nan=False, separators=(",", ":"))
        self.assertEqual(encoded, '{"liquidationLoss":null}')
        self.assertIsNone(json.loads(encoded)["liquidationLoss"])

    def test_no_field_of_any_recommendation_is_nan_or_infinite(self):
        _stat01, recs = _run(_mixed_frame(), _mixed_soup())
        offenders = []
        for rec in recs:
            for key, value in rec.items():
                kind = _bad_float(value)
                if kind is not None:
                    offenders.append((rec.get("gp"), key, kind))
        self.assertEqual(offenders, [])

    def test_a_real_jsonresponse_can_be_constructed_from_the_payload(self):
        """JSONResponse renders in __init__, so constructing it IS the test --
        this is the exact call that raised in production."""
        from starlette.responses import JSONResponse
        _stat01, recs = _run(_mixed_frame(), _mixed_soup())
        response = JSONResponse({"recommendations": recs})
        self.assertEqual(response.status_code, 200)


@unittest.skipUnless(REAL, "set FALCONBRIDGE_JULY_SOUP_XLSX / _MEMO_XLSX / "
                           "_JULY_CBT_SRF_XLSX")
class RealEndpointTests(unittest.TestCase):
    """POST /api/audit through the real app, on the real July CB&T fixture."""

    XL = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    @classmethod
    def setUpClass(cls):
        from fastapi.testclient import TestClient
        import main
        name = ("Community Bank & Trust||Community Bank and Trust__"
                + os.path.basename(CBT_XLSX))
        parts = [
            ("soup_file", (os.path.basename(SOUP_XLSX),
                           open(SOUP_XLSX, "rb").read(), cls.XL)),
            ("memo_file", (os.path.basename(MEMO_XLSX),
                           open(MEMO_XLSX, "rb").read(), cls.XL)),
            ("servicer_files", (name, open(CBT_XLSX, "rb").read(), cls.XL)),
        ]
        cls.response = TestClient(main.app).post("/api/audit", files=parts)

    def test_the_endpoint_returns_200_not_400(self):
        self.assertEqual(self.response.status_code, 200)

    def test_the_whole_payload_serializes_strictly(self):
        json.dumps(self.response.json(), allow_nan=False)

    def test_no_recommendation_field_is_nan_or_infinite(self):
        recs = self.response.json().get("recommendations") or []
        self.assertTrue(recs)
        offenders = []
        for rec in recs:
            for key, value in rec.items():
                if _bad_float(value) is not None:
                    offenders.append((rec.get("gp"), key))
        self.assertEqual(offenders, [])

    def test_closed_at_zero_loan_4338609104_is_still_excluded(self):
        """Pins that closed-at-zero suppression was never implicated. The
        production log named this loan next to the NaN, but the NaN was on a
        DIFFERENT loan -- the field is unconditional on every recommendation."""
        recs = self.response.json().get("recommendations") or []
        self.assertEqual([r for r in recs if str(r.get("gp")).strip() == "4338609104"],
                         [])

    def test_the_amount_fields_are_none_across_the_real_file(self):
        recs = self.response.json().get("recommendations") or []
        for rec in recs:
            for field in PAYLOAD_FIELDS:
                self.assertIsNone(rec[field])


if __name__ == "__main__":
    unittest.main()
