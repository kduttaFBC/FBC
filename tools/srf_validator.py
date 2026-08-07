#!/usr/bin/env python3
"""
SRF Workbook Validator  --  fail-closed pre-import gate.

Three layers, run in order. Layer 1 must pass before Layer 2 is meaningful,
and Layer 2 before Layer 3.

    Layer 1  Workbook integrity   schema, formula protection, data types
    Layer 2  Loan integrity       final-posting flags, instance sequence,
                                  balance chaining, principal reconciliation
    Layer 3  Portfolio integrity  loan count, control totals, duplicates

DESIGN RULE: there is no fallback logic. The validator never picks a row,
never resolves an ambiguity, never drops a loan. Any ERROR fails the workbook
and no data is imported. Where an upstream defect makes a downstream check
impossible to evaluate, the check reports UNVERIFIABLE, which is a failure --
not a pass, and not a guess.

Usage:
    python srf_validator.py WORKBOOK.xlsx
        [--expected-loans N] [--expected-balance X] [--tolerance 0.01]
        [--report out.txt] [--diagnostics]

Exit codes:  0 = PASSED (safe to import)   1 = FAILED (do not import)
             2 = could not read the workbook
"""

from __future__ import annotations

__version__ = "2.0.0"

# Template specification this validator asserts against. A template revision
# requires a new identifier, not an edit to this one.
TEMPLATE_SPEC = "TS-2026.06"

import argparse
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.formula import ArrayFormula

# --------------------------------------------------------------------------
# TEMPLATE SPECIFICATION
# Asserted against, never inferred from the file under test. A validator that
# derives its expectations from the file's own majority behaviour cannot
# detect corruption that affects the majority.
# --------------------------------------------------------------------------

# Governance naming: BR = approved business rule, VR = validation rule,
# RP = repair policy. The checks in this module are VR-class; each finding's
# `check` field is its identifier (e.g. VR loan.final_posting_balance).
# VALIDATION RULES -- VR class, frozen at v1.0

SHEET = "Servicer Inputs"
HEADER_ROW = 9
SUBTOTAL_ROW = 8
FIRST_DATA_ROW = 10

# column letter -> expected header text (leading/trailing space is significant
# in this template and is normalised before comparison)
# --------------------------------------------------------------------------
# TEMPLATE REGISTRY
#
# Two template variants are in circulation and their columns do not line up.
# The June "vF" final template deletes T/F, Corrected Interest Rate and the
# whole AD-AJ delinquency block, shifting everything left. Reading a
# Beginning Balance out of a fixed letter is therefore unsafe.
#
# Each variant is asserted against explicitly rather than inferred, so a
# header that changes meaning is still caught. A workbook matching no
# registered variant is rejected -- never guessed at.
# --------------------------------------------------------------------------

TEMPLATES = {
    "submitted-2026.06": {
        "anchors": {
            "A": "SBA GP Number", "B": "Lender Loan Number",
            "C": "Loan Status", "D": "Delinquency Status",
            "H": "Net Interest", "I": "Principal", "J": "Payment Received",
            "O": "Ending Balance", "P": "Beginning Balance",
            "Q": "Payment Receipt Date", "R": "Scheduled Payment",
            "AC": "Loan Payment Instance", "AD": "Total Payments Received",
            "AK": "Last Payment Instance?",
        },
        "roles": {"key": "A", "instance": "AC", "final_flag": "AK",
                  "principal": "I", "begin": "P", "end": "O",
                  "receipt": "Q", "dq_status": "D", "scheduled": "R"},
        "protected": ["AC", "AD", "AE", "AF", "AG", "AH", "AI", "AK",
                      "AL", "AM", "AN", "AP", "AS", "AT", "AU", "AV", "AW"],
        "numeric": ["E", "H", "I", "J", "M", "N", "O", "P", "R",
                    "Z", "AA", "AB", "AC", "AD"],
        "dates": ["K", "L", "Q", "S", "U", "V", "W", "X", "Y"],
        "subtotals": ["H", "I", "J", "O", "P", "R"],
    },
    "final-vF-2026.07": {
        "anchors": {
            "A": "SBA GP Number", "B": "Lender Loan Number",
            "C": "Loan Status", "D": "Delinquency Status",
            "F": "Net Interest", "G": "Principal", "H": "Payment Received",
            "M": "Ending Balance", "N": "Beginning Balance",
            "O": "Payment Receipt Date", "P": "Scheduled Payment",
            "AA": "Loan Payment Instance", "AB": "Last Payment Instance?",
        },
        "roles": {"key": "A", "instance": "AA", "final_flag": "AB",
                  "principal": "G", "begin": "N", "end": "M",
                  "receipt": "O", "dq_status": "D", "scheduled": "P"},
        # AG/AH/AI are a sparse "Unreported Loans" helper block, not per-row
        # derived, so they are not protected. AF is an empty spacer.
        "protected": ["AA", "AB", "AC", "AD", "AE",
                      "AJ", "AK", "AL", "AM", "AN"],
        "numeric": ["E", "F", "G", "H", "K", "L", "M", "N", "P",
                    "X", "Y", "Z", "AA"],
        "dates": ["I", "J", "O", "Q", "S", "T", "U", "V", "W"],
        "subtotals": ["F", "G", "H", "M", "N", "P"],
    },
}

# Retained for srf_repair.py, which restores formulas in the submitted
# template. Repair must resolve its own template before operating on any
# other variant.
PROTECTED_FORMULA_COLS = TEMPLATES["submitted-2026.06"]["protected"]

SUBTOTAL_ROW_OFFSET = 1

# --------------------------------------------------------------------------
# STATUS DOMAINS
#
# Two separate enumerations, because the defect that put 584,023.61 into no
# portfolio bucket was a LOAN STATUS value ('FORCLOSURE', misspelled) sitting
# in a DELINQUENCY STATUS field. One combined list would have permitted it.
#
# Comparison is case-insensitive: the 25,071-row performance dataset contains
# '30 DAYS', '30 Days' and '30 days', plus 'ACTIVE', 'Active' and 'ACTiVE'.
# Case variance is pervasive and harmless; a wrong value is not.
#
# Sources: the template's own Reference Tables sheet, the cached DQ_Lookup
# ladder, and every value observed across five servicers. Where those three
# disagree the union is permitted and the disagreement is reported --
# see reference_domain.drift and dq_status.unresolvable.
# --------------------------------------------------------------------------

PERMITTED_DQ_STATUS = {
    "CURRENT", "DELINQUENT",
    "30 DAYS", "60 DAYS", "90 DAYS", "120 DAYS",
    "150 DAYS", "150+ DAYS", "180 DAYS", "210 DAYS",
}

# Statuses the delinquency ladder can actually resolve. A value may be a
# legitimate Delinquency Status and still have no entry in DQ_Lookup: the
# ladder stops at 150 DAYS, but the dataset contains 180 DAYS and 210 DAYS.
# Such a loan passes the domain check and then fails status movement with
# #N/A, so it is reported separately rather than conflated.
LADDER_RESOLVABLE_DQ_STATUS = {
    "CURRENT", "DELINQUENT",
    "30 DAYS", "60 DAYS", "90 DAYS", "120 DAYS", "150 DAYS",
}

PERMITTED_LOAN_STATUS = {
    "ACTIVE", "DEFERRED", "PAID-IN-FULL", "REO", "FORECLOSURE",
    "BANKRUPTCY", "LIQUIDATION", "SUBSTITUTION-IN", "SUBSTITUTION-OUT",
}

# Where the template declares its own domains, so the spec above can be
# cross-checked against the workbook rather than merely asserted.
REFERENCE_SHEET = "Reference Tables"
REFERENCE_HEADER_ROW = 1
REFERENCE_DOMAIN_COLS = {"loan_status": "D", "dq_status": "G"}

PRIOR_SHEET = "Loan Terms & Prior Period"
PRIOR_KEY_COL = "A"
PRIOR_BALANCE_COL = "S"
PRIOR_HEADER_ROW = 9


# --------------------------------------------------------------------------
# FINDINGS
# --------------------------------------------------------------------------

ERROR = "ERROR"
UNVERIFIABLE = "UNVERIFIABLE"
INFO = "INFO"


@dataclass
class Finding:
    layer: int
    severity: str
    check: str
    scope: str          # "workbook", "loan 9031018509", "row 59"
    expected: str
    found: str

    @property
    def fatal(self) -> bool:
        return self.severity in (ERROR, UNVERIFIABLE)


@dataclass
class Result:
    path: str
    findings: list[Finding] = field(default_factory=list)
    stats: dict = field(default_factory=dict)
    diagnostics: list[str] = field(default_factory=list)
    cols: dict = field(default_factory=dict)

    def add(self, *a, **k) -> None:
        self.findings.append(Finding(*a, **k))

    def fatal_in(self, layer: int) -> bool:
        return any(f.fatal and f.layer == layer for f in self.findings)

    @property
    def passed(self) -> bool:
        return not any(f.fatal for f in self.findings)


# --------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------

def is_formula(v) -> bool:
    if isinstance(v, ArrayFormula):
        return True
    return isinstance(v, str) and v.startswith("=")


def norm_header(v) -> str:
    if v is None:
        return ""
    if is_formula(v):
        return ""          # some header cells are themselves formulas
    return " ".join(str(v).replace("\xa0", " ").split())


def is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def as_number(v):
    """Return float if v is genuinely numeric, else None."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return None


def detect_template(wsf):
    """
    Return (name, spec) for the registered variant whose anchor headers all
    match, or (None, None). No fuzzy matching: every anchor must be exact.
    """
    best = (None, None, -1)
    for name, spec in TEMPLATES.items():
        hits = sum(1 for L, want in spec["anchors"].items()
                   if norm_header(wsf.cell(HEADER_ROW,
                                           column_index_from_string(L)).value)
                   == want)
        if hits == len(spec["anchors"]) and hits > best[2]:
            best = (name, spec, hits)
    return best[0], best[1]


def resolve(spec) -> dict:
    """Column letters -> indices for the matched template."""
    ix = column_index_from_string
    return {
        "roles": {r: ix(L) for r, L in spec["roles"].items()},
        "role_letters": dict(spec["roles"]),
        "protected": [(L, ix(L)) for L in spec["protected"]],
        "numeric": [(L, ix(L)) for L in spec["numeric"]],
        "dates": [(L, ix(L)) for L in spec["dates"]],
        "subtotals": [(L, ix(L)) for L in spec["subtotals"]],
    }


def fmt(v) -> str:
    n = as_number(v)
    if n is not None:
        return f"{n:,.2f}"
    return repr(v)


# --------------------------------------------------------------------------
# LAYER 1 -- WORKBOOK INTEGRITY
# --------------------------------------------------------------------------

def layer1(wsf, wsv, res: Result) -> list[int]:
    """Returns the list of data row numbers. wsf = formula view, wsv = value view."""

    # -- identify which registered template this workbook is ------------
    name, spec = detect_template(wsf)
    if spec is None:
        detail = []
        for tname, tspec in TEMPLATES.items():
            bad = [f"{L}9 is '{norm_header(wsf.cell(HEADER_ROW, column_index_from_string(L)).value)}'"
                   f" not '{want}'"
                   for L, want in tspec["anchors"].items()
                   if norm_header(wsf.cell(HEADER_ROW,
                                           column_index_from_string(L)).value) != want]
            detail.append(f"{tname}: {len(bad)} anchor mismatch(es)"
                          + (f" e.g. {bad[0]}" if bad else ""))
        res.add(1, ERROR, "schema.template", "workbook",
                f"one of the registered templates {list(TEMPLATES)}",
                "; ".join(detail) + ". Column positions cannot be trusted, "
                "so no further check is evaluable.")
        return []
    cols = resolve(spec)
    res.stats["template"] = name
    res.cols = cols

    # -- identify data rows ---------------------------------------------
    key_ix = cols["roles"]["key"]
    rows = [r for r in range(FIRST_DATA_ROW, wsv.max_row + 1)
            if not is_blank(wsv.cell(r, key_ix).value)]
    res.stats["data_rows"] = len(rows)
    if not rows:
        res.add(1, ERROR, "schema.rows", "workbook",
                "at least one data row", "no populated data rows found")
        return []

    # -- gaps inside the data block -------------------------------------
    if rows != list(range(rows[0], rows[-1] + 1)):
        missing = sorted(set(range(rows[0], rows[-1] + 1)) - set(rows))
        res.add(1, ERROR, "schema.contiguity", "workbook",
                "no blank rows inside the data block",
                f"blank row(s) at {missing}")

    # -- protected formula columns --------------------------------------
    for letter, c in cols["protected"]:
        hdr = norm_header(wsf.cell(HEADER_ROW, c).value) or letter
        for r in rows:
            raw = wsf.cell(r, c).value
            if is_formula(raw):
                continue
            if raw is None:
                res.add(1, ERROR, "formula.deleted", f"row {r} / {letter}",
                        f"formula in '{hdr}'", "cell is empty -- formula deleted")
            else:
                res.add(1, ERROR, "formula.overwritten", f"row {r} / {letter}",
                        f"formula in '{hdr}'", f"hardcoded literal {raw!r}")

    # -- workbook must have been recalculated since its last edit ---------
    # openpyxl (and any programmatic edit) writes formulas without cached
    # values. Reading such a file yields None for every formula cell, which
    # would surface as spurious Layer 2 loan defects. Fail here instead.
    cached = 0
    for letter, c in cols["protected"]:
        for r in rows:
            if wsv.cell(r, c).value is not None:
                cached += 1
    if cached == 0:
        res.add(1, ERROR, "workbook.recalculated", "workbook",
                "cached values present in formula columns",
                "no formula in any protected column has a cached value -- "
                "the workbook has been edited programmatically and not "
                "recalculated. Open and recalculate in Excel, then re-validate. "
                "Loan-level results below are not evaluable and are suppressed.")
        return []

    # -- data types ------------------------------------------------------
    for letter, c in cols["numeric"]:
        hdr = norm_header(wsf.cell(HEADER_ROW, c).value) or letter
        for r in rows:
            v = wsv.cell(r, c).value
            if is_blank(v):
                continue
            if as_number(v) is None:
                res.add(1, ERROR, "dtype.numeric", f"row {r} / {letter}",
                        f"number in '{hdr}'", f"{type(v).__name__} {v!r}")

    for letter, c in cols["dates"]:
        hdr = norm_header(wsf.cell(HEADER_ROW, c).value) or letter
        for r in rows:
            v = wsv.cell(r, c).value
            if is_blank(v):
                continue
            if not isinstance(v, (datetime, date)):
                res.add(1, ERROR, "dtype.date", f"row {r} / {letter}",
                        f"date in '{hdr}'", f"{type(v).__name__} {v!r}")

    # -- instance column must be a positive integer ----------------------
    ci = cols["roles"]["instance"]
    for r in rows:
        v = wsv.cell(r, ci).value
        n = as_number(v)
        if n is None:
            continue  # already reported by dtype check
        if n != int(n) or n < 1:
            res.add(1, ERROR, "dtype.instance", f"row {r} / {cols['role_letters']['instance']}",
                    "positive integer", fmt(v))

    # -- final-posting flag domain ---------------------------------------
    cf = cols["roles"]["final_flag"]
    for r in rows:
        v = wsv.cell(r, cf).value
        if is_blank(v):
            res.add(1, ERROR, "flag.blank", f"row {r} / {cols['role_letters']['final_flag']}",
                    "'Y' or 'N'", "blank")
        elif str(v).strip().upper() not in ("Y", "N"):
            res.add(1, ERROR, "flag.domain", f"row {r} / {cols['role_letters']['final_flag']}",
                    "'Y' or 'N'", repr(v))

    # -- status enumerations ---------------------------------------------
    for role, permitted, check in (
            ("dq_status", PERMITTED_DQ_STATUS, "dq_status.domain"),
            ("loan_status", PERMITTED_LOAN_STATUS, "loan_status.domain")):
        if role not in cols["roles"]:
            continue
        c = cols["roles"][role]
        letter = cols["role_letters"][role]
        hdr = norm_header(wsf.cell(HEADER_ROW, c).value) or letter
        seen_case = {}
        for r in rows:
            v = wsv.cell(r, c).value
            if is_blank(v):
                continue
            raw = str(v).strip()
            canon = raw.upper()
            seen_case.setdefault(canon, set()).add(raw)
            if canon not in permitted:
                res.add(1, ERROR, check, f"row {r} / {letter}",
                        f"'{hdr}' one of {sorted(permitted)}",
                        f"{v!r} -- outside the permitted enumeration. A value "
                        f"from the wrong domain or a misspelling is dropped "
                        f"from every portfolio bucket while remaining in the "
                        f"outstanding total.")
        mixed = {k: sorted(v) for k, v in seen_case.items() if len(v) > 1}
        if mixed:
            res.add(1, INFO, f"{check}.case", f"column {letter} ('{hdr}')",
                    "one spelling per status",
                    f"case variants present: {mixed}. Matched case-"
                    f"insensitively, but worth normalising at source.")

    # -- a permitted status the delinquency ladder cannot resolve ---------
    if "dq_status" in cols["roles"]:
        c = cols["roles"]["dq_status"]
        letter = cols["role_letters"]["dq_status"]
        for r in rows:
            v = wsv.cell(r, c).value
            if is_blank(v):
                continue
            canon = str(v).strip().upper()
            if canon in PERMITTED_DQ_STATUS and \
                    canon not in LADDER_RESOLVABLE_DQ_STATUS:
                res.add(1, ERROR, "dq_status.unresolvable",
                        f"row {r} / {letter}",
                        f"a status the delinquency ladder can resolve "
                        f"{sorted(LADDER_RESOLVABLE_DQ_STATUS)}",
                        f"{v!r} is a permitted status but has no entry in the "
                        f"delinquency ladder, which stops at 150 DAYS. Status "
                        f"movement for this loan resolves to #N/A.")

    # -- cross-check the spec against the template's declared domains -----
    if REFERENCE_SHEET in wsf.parent.sheetnames:
        rt = wsf.parent[REFERENCE_SHEET]
        for role, letter in REFERENCE_DOMAIN_COLS.items():
            c = column_index_from_string(letter)
            declared = {str(rt.cell(r, c).value).strip().upper()
                        for r in range(REFERENCE_HEADER_ROW + 1,
                                       REFERENCE_HEADER_ROW + 20)
                        if rt.cell(r, c).value is not None}
            spec = (PERMITTED_DQ_STATUS if role == "dq_status"
                    else PERMITTED_LOAN_STATUS)
            only_spec = sorted(spec - declared)
            only_decl = sorted(declared - spec)
            if only_spec or only_decl:
                res.add(1, INFO, "reference_domain.drift",
                        f"{REFERENCE_SHEET}!{letter} ({role})",
                        "the template's declared domain matches the "
                        "validator's enumeration",
                        f"permitted but not declared: {only_spec}; declared "
                        f"but not permitted: {only_decl}")

    # -- internal control: subtotal row vs recomputed sums ---------------
    for letter, c in cols["subtotals"]:
        hdr = norm_header(wsf.cell(HEADER_ROW, c).value) or letter
        stated = as_number(wsv.cell(SUBTOTAL_ROW, c).value)
        if stated is None:
            continue
        recomputed = sum(as_number(wsv.cell(r, c).value) or 0.0 for r in rows)
        if abs(stated - recomputed) > res.stats["tolerance"]:
            res.add(1, ERROR, "control.subtotal", f"column {letter} ('{hdr}')",
                    f"subtotal {stated:,.2f}",
                    f"sum of rows {recomputed:,.2f} "
                    f"(differs by {recomputed - stated:,.2f})")

    return rows


# --------------------------------------------------------------------------
# LAYER 2 -- LOAN INTEGRITY
# --------------------------------------------------------------------------

def check_prior_period(wsv, prior_ws, rows: list[int], res: Result,
                       cols: dict) -> None:
    """Each loan's first posting must open at the prior period's closing balance."""
    if prior_ws is None:
        res.add(2, INFO, "loan.prior_period_continuity", "workbook",
                f"sheet '{PRIOR_SHEET}' present",
                "sheet not found -- continuity with the prior period was "
                "not verified")
        return

    pk = column_index_from_string(PRIOR_KEY_COL)
    pb = column_index_from_string(PRIOR_BALANCE_COL)
    prior: dict[object, float] = {}
    for r in range(PRIOR_HEADER_ROW + 1, prior_ws.max_row + 1):
        g = prior_ws.cell(r, pk).value
        v = as_number(prior_ws.cell(r, pb).value)
        if g is not None and v is not None:
            prior[g] = v

    ck = cols["roles"]["key"]
    ci = cols["roles"]["instance"]
    cb = cols["roles"]["begin"]
    tol = res.stats["tolerance"]

    for r in rows:
        if as_number(wsv.cell(r, ci).value) != 1:
            continue
        loan = wsv.cell(r, ck).value
        opening = as_number(wsv.cell(r, cb).value)
        if opening is None:
            continue
        if loan not in prior:
            res.add(2, ERROR, "loan.prior_period_missing", f"loan {loan}",
                    f"a record in '{PRIOR_SHEET}'",
                    "loan absent from the prior-period sheet")
            continue
        if abs(prior[loan] - opening) > tol:
            res.add(2, ERROR, "loan.prior_period_continuity",
                    f"loan {loan} (row {r})",
                    f"opening = prior-period closing {prior[loan]:,.2f}",
                    f"opening = {opening:,.2f} "
                    f"(differs by {opening - prior[loan]:,.2f})")


def layer2(wsv, rows: list[int], res: Result, diagnostics: bool,
           cols: dict) -> None:
    R = cols["roles"]
    ck, ci, cf = R["key"], R["instance"], R["final_flag"]
    cp, cb, ce, cq = R["principal"], R["begin"], R["end"], R["receipt"]
    tol = res.stats["tolerance"]

    # A posting separated from the contiguous data block is an orphan. In the
    # June vF workbook one sits at row 114, 47 rows below the data, carrying a
    # loan number and nothing else. COUNTIFS counts it, so it takes the final
    # instance number and the "Y" flag lands on it instead of the real closing
    # posting -- moving 650,128.27 out of the portfolio.
    block_end = rows[0]
    for r in rows:
        if r - block_end > 1:
            break
        block_end = r
    orphan_rows = [r for r in rows if r > block_end]

    loans: dict[object, list[dict]] = defaultdict(list)
    for r in rows:
        loans[wsv.cell(r, ck).value].append({
            "row": r,
            "orphan": r in orphan_rows,
            "instance": as_number(wsv.cell(r, ci).value),
            "flag": (str(wsv.cell(r, cf).value).strip().upper()
                     if not is_blank(wsv.cell(r, cf).value) else None),
            "principal": as_number(wsv.cell(r, cp).value) or 0.0,
            "begin": as_number(wsv.cell(r, cb).value),
            "end": as_number(wsv.cell(r, ce).value),
            "receipt": wsv.cell(r, cq).value,
        })

    res.stats["loan_count"] = len(loans)

    for loan, ps in loans.items():
        scope = f"loan {loan}"

        # -- CONDITION 3: no orphaned or disconnected posting -----------
        orph = [p["row"] for p in ps if p["orphan"]]
        if orph:
            res.add(2, ERROR, "loan.orphan_posting", scope,
                    f"all postings inside the contiguous data block "
                    f"(rows {rows[0]}-{block_end})",
                    f"posting(s) at row(s) {orph}, separated from the block. "
                    f"A disconnected row is still counted by COUNTIFS, so it "
                    f"can take the final instance number and the 'Y' flag "
                    f"away from the real closing posting.")

        # -- CONDITION 1: exactly one final posting ---------------------
        n_final = sum(1 for p in ps if p["flag"] == "Y")
        if n_final != 1:
            where = [p["row"] for p in ps if p["flag"] == "Y"]
            res.add(2, ERROR, "loan.final_posting", scope,
                    "exactly 1 posting flagged 'Y'",
                    f"{n_final} flagged 'Y'"
                    + (f" (rows {where})" if where else "")
                    + f" across {len(ps)} posting(s)")

        # -- CONDITION 2: the final posting carries a valid ending balance
        # "Exactly one Y" is not sufficient. In the vF workbook there IS
        # exactly one Y -- on the orphan row, which has no balance at all.
        # An importer selecting the flagged row gets nothing for the loan.
        flagged = [p for p in ps if p["flag"] == "Y"]
        for p in flagged:
            missing = [n for n, v in (("ending balance", p["end"]),
                                      ("beginning balance", p["begin"]))
                       if v is None]
            if missing:
                res.add(2, ERROR, "loan.final_posting_balance",
                        f"{scope} (row {p['row']})",
                        "the posting flagged 'Y' carries a valid balance",
                        f"{' and '.join(missing)} absent. An importer "
                        f"selecting the flagged posting would take no "
                        f"balance for this loan.")

        # -- instance sequence 1..n, no duplicates ----------------------
        insts = [p["instance"] for p in ps]
        seq_ok = (None not in insts
                  and sorted(insts) == [float(i) for i in range(1, len(ps) + 1)])
        if not seq_ok:
            dupes = [v for v, n in Counter(insts).items() if n > 1]
            res.add(2, ERROR, "loan.instance_sequence", scope,
                    f"instances 1..{len(ps)}",
                    f"instances {[None if i is None else int(i) for i in insts]}"
                    + (f"; duplicated {[int(d) for d in dupes if d is not None]}"
                       if dupes else ""))

        # -- ordering-dependent checks ----------------------------------
        # Without a trustworthy instance sequence there is no defensible
        # posting order, so these are UNVERIFIABLE rather than guessed.
        if not seq_ok:
            res.add(2, UNVERIFIABLE, "loan.balance_chain", scope,
                    "balance chain verified",
                    "cannot establish posting order -- instance sequence invalid")
            res.add(2, UNVERIFIABLE, "loan.principal_reconciliation", scope,
                    "opening - total principal = final ending",
                    "cannot identify opening/final posting -- "
                    "instance sequence invalid")
            if diagnostics:
                _diagnose(loan, ps, res, tol)
            continue

        ordered = sorted(ps, key=lambda p: p["instance"])

        # -- ending = beginning - principal, per posting -----------------
        for p in ordered:
            if p["begin"] is None or p["end"] is None:
                res.add(2, ERROR, "loan.balance_present", f"row {p['row']}",
                        "beginning and ending balance populated",
                        f"beginning={fmt(p['begin'])} ending={fmt(p['end'])}")
                continue
            want = p["begin"] - p["principal"]
            if abs(want - p["end"]) > tol:
                res.add(2, ERROR, "loan.ending_balance",
                        f"{scope} instance {int(p['instance'])} (row {p['row']})",
                        f"ending = {p['begin']:,.2f} - {p['principal']:,.2f} "
                        f"= {want:,.2f}",
                        f"ending = {p['end']:,.2f} "
                        f"(differs by {p['end'] - want:,.2f})")

        # -- beginning chains off prior ending --------------------------
        for prev, cur in zip(ordered, ordered[1:]):
            if prev["end"] is None or cur["begin"] is None:
                continue
            if abs(cur["begin"] - prev["end"]) > tol:
                res.add(2, ERROR, "loan.balance_chain",
                        f"{scope} instance {int(cur['instance'])} "
                        f"(row {cur['row']})",
                        f"beginning = prior ending {prev['end']:,.2f}",
                        f"beginning = {cur['begin']:,.2f} "
                        f"(differs by {cur['begin'] - prev['end']:,.2f})")

        # -- opening - total principal = final ending -------------------
        # Holds regardless of posting count. Catches chain corruption even
        # when every individual posting is internally consistent.
        opening, final_end = ordered[0]["begin"], ordered[-1]["end"]
        if opening is not None and final_end is not None:
            total_prin = sum(p["principal"] for p in ordered)
            want = opening - total_prin
            if abs(want - final_end) > tol:
                res.add(2, ERROR, "loan.principal_reconciliation", scope,
                        f"opening {opening:,.2f} - principal {total_prin:,.2f} "
                        f"= {want:,.2f}",
                        f"final ending {final_end:,.2f} "
                        f"(differs by {final_end - want:,.2f})")

        # -- the 'Y' must sit on the highest instance --------------------
        flagged = [p for p in ps if p["flag"] == "Y"]
        if len(flagged) == 1 and flagged[0]["instance"] != ordered[-1]["instance"]:
            res.add(2, ERROR, "loan.final_posting_position", scope,
                    f"'Y' on instance {int(ordered[-1]['instance'])}",
                    f"'Y' on instance {int(flagged[0]['instance'])} "
                    f"(row {flagged[0]['row']})")


def _diagnose(loan, ps, res: Result, tol: float) -> None:
    """Analyst-facing only. Never influences pass/fail."""
    dated = [p for p in ps if isinstance(p["receipt"], (datetime, date))]
    if len(dated) != len(ps):
        return
    o = sorted(dated, key=lambda p: p["receipt"])
    if o[0]["begin"] is None or o[-1]["end"] is None:
        return
    tp = sum(p["principal"] for p in o)
    want = o[0]["begin"] - tp
    res.diagnostics.append(
        f"loan {loan}: IF postings were ordered by receipt date, opening "
        f"{o[0]['begin']:,.2f} - principal {tp:,.2f} = {want:,.2f}, versus a "
        f"reported final ending of {o[-1]['end']:,.2f} "
        f"(difference {o[-1]['end'] - want:,.2f}). Order is not established; "
        f"this figure is indicative only and is not used for validation."
    )


# --------------------------------------------------------------------------
# LAYER 3 -- PORTFOLIO INTEGRITY
# --------------------------------------------------------------------------

def layer3(wsv, rows: list[int], res: Result,
           expected_loans: int | None, expected_balance: float | None,
           cols: dict) -> None:
    ck = cols["roles"]["key"]
    ci = cols["roles"]["instance"]
    ce = cols["roles"]["end"]
    tol = res.stats["tolerance"]

    loans: dict[object, list[int]] = defaultdict(list)
    for r in rows:
        loans[wsv.cell(r, ck).value].append(r)

    # -- loan identifier hygiene ----------------------------------------
    for loan, rs in loans.items():
        if not isinstance(loan, (int, float)) and not str(loan).strip().isdigit():
            res.add(3, ERROR, "portfolio.loan_id", f"loan {loan!r}",
                    "numeric SBA GP Number", f"{type(loan).__name__} {loan!r}")

    # -- portfolio outstanding balance ----------------------------------
    # The final posting per loan, identified by highest instance. Only
    # computed for loans whose instance sequence is sound; a loan that
    # failed Layer 2 has no defensible final posting, so it is reported
    # as excluded rather than silently estimated.
    total, excluded = 0.0, []
    for loan, rs in loans.items():
        insts = [(as_number(wsv.cell(r, ci).value), r) for r in rs]
        if any(i is None for i, _ in insts) or \
           sorted(i for i, _ in insts) != [float(k) for k in range(1, len(rs) + 1)]:
            excluded.append(loan)
            continue
        last_row = max(insts, key=lambda t: t[0])[1]
        total += as_number(wsv.cell(last_row, ce).value) or 0.0

    res.stats["loans_with_sound_sequence"] = len(loans) - len(excluded)
    res.stats["outstanding_balance_sound_loans"] = total

    if excluded:
        res.add(3, UNVERIFIABLE, "portfolio.balance_coverage", "workbook",
                f"all {len(loans)} loans contribute to the portfolio total",
                f"{len(excluded)} loan(s) excluded -- no defensible final "
                f"posting: {excluded}")

    # -- control totals --------------------------------------------------
    if expected_loans is not None:
        if len(loans) != expected_loans:
            res.add(3, ERROR, "portfolio.loan_count", "workbook",
                    f"{expected_loans} loans",
                    f"{len(loans)} loans "
                    f"(differs by {len(loans) - expected_loans:+d})")
    else:
        res.add(3, INFO, "portfolio.loan_count", "workbook",
                "control total supplied via --expected-loans",
                f"not supplied; observed {len(loans)} loans")

    if expected_balance is not None:
        if excluded:
            res.add(3, UNVERIFIABLE, "portfolio.control_balance", "workbook",
                    f"portfolio balance {expected_balance:,.2f}",
                    "cannot be evaluated -- loans excluded above")
        elif abs(total - expected_balance) > tol:
            res.add(3, ERROR, "portfolio.control_balance", "workbook",
                    f"portfolio balance {expected_balance:,.2f}",
                    f"{total:,.2f} (differs by {total - expected_balance:,.2f})")
    else:
        res.add(3, INFO, "portfolio.control_balance", "workbook",
                "control total supplied via --expected-balance",
                f"not supplied; observed {total:,.2f} across "
                f"{res.stats['loans_with_sound_sequence']} sound loan(s)")


# --------------------------------------------------------------------------
# REPORT
# --------------------------------------------------------------------------

LAYER_NAMES = {1: "WORKBOOK INTEGRITY", 2: "LOAN INTEGRITY",
               3: "PORTFOLIO INTEGRITY"}


def render(res: Result) -> str:
    out: list[str] = []
    w = 74
    out.append("=" * w)
    out.append(f"SRF WORKBOOK VALIDATION REPORT  (validator {__version__})")
    out.append("=" * w)
    out.append(f"Workbook   : {res.path}")
    out.append(f"Sheet      : {SHEET}")
    out.append(f"Template   : {res.stats.get('template', 'unidentified')}")
    out.append(f"Data rows  : {res.stats.get('data_rows', 'n/a')}")
    out.append(f"Loans      : {res.stats.get('loan_count', 'n/a')}")
    out.append(f"Tolerance  : {res.stats['tolerance']:.2f}")
    out.append("")

    # ---- actionable summary: collapse findings to the cells to fix --------
    fatal = [f for f in res.findings if f.fatal]
    if fatal:
        by_row: dict[str, list[str]] = defaultdict(list)
        by_loan: dict[str, list[str]] = defaultdict(list)
        for f in fatal:
            if f.scope.startswith("row "):
                parts = f.scope.split("/")
                row = parts[0].strip()
                col = parts[1].strip() if len(parts) > 1 else "?"
                by_row[row].append(col)
            elif f.scope.startswith("loan "):
                by_loan[f.scope].append(f.check.split(".", 1)[-1])
        out.append("-" * w)
        out.append("SUMMARY -- what to fix")
        out.append("-" * w)
        if by_row:
            out.append(f"  {len(by_row)} row(s) with cell-level defects:")
            for row in sorted(by_row, key=lambda s: int(s.split()[1])):
                cols = sorted(set(by_row[row]),
                              key=lambda c: (len(c), c) if c != "?" else (9, c))
                out.append(f"      {row:<9} -> {', '.join(cols)}")
        if by_loan:
            out.append(f"  {len(by_loan)} loan(s) with integrity defects:")
            for loan in sorted(by_loan):
                out.append(f"      {loan:<22} -> "
                           f"{', '.join(sorted(set(by_loan[loan])))}")
        out.append("")

    for layer in (1, 2, 3):
        fs = [f for f in res.findings if f.layer == layer]
        fatal = [f for f in fs if f.fatal]
        out.append("-" * w)
        status = "PASSED" if not fatal else f"FAILED ({len(fatal)} finding(s))"
        out.append(f"LAYER {layer} -- {LAYER_NAMES[layer]:<34} {status}")
        out.append("-" * w)
        if not fs:
            out.append("  no findings")
        for f in fs:
            out.append(f"  [{f.severity}] {f.check}")
            out.append(f"      scope    : {f.scope}")
            out.append(f"      expected : {f.expected}")
            out.append(f"      found    : {f.found}")
        out.append("")

    if res.diagnostics:
        out.append("-" * w)
        out.append("DIAGNOSTICS -- indicative only, not used for validation")
        out.append("-" * w)
        for d in res.diagnostics:
            out.append(f"  {d}")
        out.append("")

    out.append("=" * w)
    if res.passed:
        out.append("IMPORT STATUS : PASSED")
        out.append("Workbook may be imported.")
    else:
        n = sum(1 for f in res.findings if f.fatal)
        out.append("IMPORT STATUS : FAILED")
        out.append(f"{n} blocking finding(s). NO DATA IMPORTED.")
        out.append("")
        out.append("No row was selected, no ambiguity resolved, and no loan")
        out.append("dropped. Correct the workbook at source and re-validate.")
    out.append("=" * w)
    return "\n".join(out)


# --------------------------------------------------------------------------
# ENTRY POINT
# --------------------------------------------------------------------------

def validate(path: str, tolerance: float = 0.01,
             expected_loans: int | None = None,
             expected_balance: float | None = None,
             diagnostics: bool = False) -> Result:
    res = Result(path=path)
    res.stats["tolerance"] = tolerance

    try:
        wbf = openpyxl.load_workbook(path)                  # formulas
        wbv = openpyxl.load_workbook(path, data_only=True)  # cached values
    except Exception as e:
        res.add(1, ERROR, "workbook.readable", path,
                "a readable .xlsx workbook", f"{type(e).__name__}: {e}")
        return res

    if SHEET not in wbf.sheetnames:
        res.add(1, ERROR, "schema.sheet", path,
                f"sheet '{SHEET}'", f"sheets {wbf.sheetnames}")
        return res

    wsf, wsv = wbf[SHEET], wbv[SHEET]

    rows = layer1(wsf, wsv, res)
    if not rows:
        return res

    layer2(wsv, rows, res, diagnostics, res.cols)
    check_prior_period(wsv, wbv[PRIOR_SHEET] if PRIOR_SHEET in wbv.sheetnames
                       else None, rows, res, res.cols)
    layer3(wsv, rows, res, expected_loans, expected_balance, res.cols)
    return res


def main() -> int:
    ap = argparse.ArgumentParser(description="Fail-closed SRF workbook validator")
    ap.add_argument("workbook")
    ap.add_argument("--tolerance", type=float, default=0.01)
    ap.add_argument("--expected-loans", type=int, default=None)
    ap.add_argument("--expected-balance", type=float, default=None)
    ap.add_argument("--report", default=None, help="also write the report here")
    ap.add_argument("--diagnostics", action="store_true",
                    help="append indicative figures for unverifiable loans")
    ap.add_argument("--version", action="version",
                    version=f"srf_validator {__version__}")
    a = ap.parse_args()

    res = validate(a.workbook, a.tolerance, a.expected_loans,
                   a.expected_balance, a.diagnostics)
    text = render(res)
    print(text)
    if a.report:
        with open(a.report, "w") as fh:
            fh.write(text + "\n")

    if any(f.check == "workbook.readable" or f.check == "schema.sheet"
           for f in res.findings if f.fatal):
        return 2
    return 0 if res.passed else 1


if __name__ == "__main__":
    sys.exit(main())
