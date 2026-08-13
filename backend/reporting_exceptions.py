"""Reporting Exception Register — CLIENT-CONFIRMED (2026-08-12).

A narrowly scoped mechanism for recording that an ALREADY-ISSUED Monthly
Servicer Report carries a value differing from what this application calculates
from the corrected lender source data, where the client has confirmed the
issued report is authoritative for that period and has NOT supplied a general
business rule explaining it.

THIS IS NOT A CALCULATION AND NOT A CORRECTION.
----------------------------------------------
The standard Net Interest / Total Payment Received calculation in checks.py is
untouched and keeps producing the lender-derived value. An exception is applied
only at the REPORT-WRITING layer, after that calculation has run, and only for
the exact (reporting period, deal, SBA GP Number, field) it names. The
system-calculated value is preserved on the record beside the authoritative one
so the difference is always visible and auditable.

Deliberately NOT routed through corrections.py, for four reasons:
  * checks._VERIFIED_OVERRIDE_COLUMNS admits only Loan Status and Delinquency
    Status; carrying a financial field there would widen the correction
    allow-list, which is explicitly forbidden.
  * corrections.generate_reviewed_workbook writes into a copy of the LENDER's
    workbook. The lender data is correct here -- only the report differs.
  * A correction asserts "the lender's value was wrong". This asserts "the
    lender's value is right, and the issued report differs anyway".
  * Correction state is session-scoped with no reporting period; an exception
    must persist as historical evidence and must never reach another period.

AN EXCEPTION IS NEVER A RULE.
-----------------------------
Matching requires an exact match on ALL FOUR keys. There is no pattern, no
wildcard, no status predicate and no inheritance. April's entries are
unreachable from May by construction. No PAID-IN-FULL, SUBSTITUTION-OUT,
payoff-interest or terminal-status rule exists anywhere in this module, and no
SBA GP Number appears in any calculation path -- the register is looked up,
never branched on.

THE STALE-EXCEPTION GUARD.
--------------------------
Before applying, the recorded system_calculated_value is compared against what
the calculation JUST produced. If they differ, the exception does NOT apply and
is surfaced for review. This is what stops an exception silently becoming a
rule: if the lender restates a figure, or a future change alters the
calculation, the exception detaches and is reported rather than masking the
change.
"""

import json
import os

# The ONLY fields the register may carry. CLIENT-CONFIRMED (2026-08-12): the
# initial register is limited to these two. Ending Balance, Liquidation Date,
# Liquidation Loss, Modification Comments, Remittance Date, Note Account, Loan
# Status and Delinquency Status are explicitly excluded, and an entry naming
# any other field is rejected rather than applied.
EXCEPTION_FIELDS = ("Net Interest", "Total Payment Received")

# The four keys that must ALL match. Nothing is inferred from any of them.
EXCEPTION_KEYS = ("reporting_period", "deal", "sba_gp_number", "field")

REGISTER_DIRNAME = "reporting_exceptions"

# Half a cent. Tight enough that a real restatement (which moves a reported
# figure by at least a cent) detaches the exception, loose enough that float
# representation noise on a summed value does not. Deliberately absolute, not
# relative: a cent is a cent whether the balance is 200 or 200,000.
STALE_TOLERANCE = 0.005

ISSUE_EXCEPTION_APPLIED = "reporting_exception_applied"
ISSUE_EXCEPTION_STALE = "reporting_exception_stale"
ISSUE_EXCEPTION_UNMATCHED = "reporting_exception_unmatched"
ISSUE_EXCEPTION_REJECTED = "reporting_exception_rejected"


def register_directory():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), REGISTER_DIRNAME)


def load_register(directory=None):
    """Every exception entry on disk, as a flat list.

    Entries are DATA, not code -- one JSON file per (reporting period, deal).
    A malformed or unreadable file yields nothing rather than raising: a
    missing register must degrade to "no exceptions", never to a failed report.
    """
    directory = directory or register_directory()
    if not os.path.isdir(directory):
        return []
    entries = []
    for name in sorted(os.listdir(directory)):
        if not name.endswith(".json"):
            continue
        try:
            with open(os.path.join(directory, name), "r") as handle:
                payload = json.load(handle)
        except Exception:
            continue
        if isinstance(payload, dict):
            payload = payload.get("exceptions") or []
        for entry in payload if isinstance(payload, list) else []:
            if isinstance(entry, dict):
                entry = dict(entry)
                entry.setdefault("_register_file", name)
                entries.append(entry)
    return entries


def _key(reporting_period, deal, gp, field):
    return (str(reporting_period or "").strip(),
            str(deal or "").strip(),
            str(gp or "").strip(),
            str(field or "").strip())


def _period_label(cycle_month):
    """A Timestamp/date -> "2026-04". None stays None, so a run with no
    resolvable reporting period matches no exception at all."""
    if cycle_month is None:
        return None
    try:
        return cycle_month.strftime("%Y-%m")
    except AttributeError:
        text = str(cycle_month).strip()
        return text[:7] if len(text) >= 7 else None


def index_register(entries, issues=None):
    """{(period, deal, gp, field): entry}, rejecting anything malformed.

    An entry naming a field outside EXCEPTION_FIELDS is REJECTED and surfaced,
    never applied -- the register cannot be used to reach a field the client
    did not confirm."""
    issues = issues if issues is not None else []
    index = {}
    for entry in entries or []:
        missing = [k for k in EXCEPTION_KEYS if not str(entry.get(k) or "").strip()]
        field = str(entry.get("field") or "").strip()
        if missing or field not in EXCEPTION_FIELDS or "system_calculated_value" not in entry:
            issues.append({
                "code": ISSUE_EXCEPTION_REJECTED,
                "field": field or None,
                "SBA GP Number": entry.get("sba_gp_number"),
                "message": (
                    f"A reporting exception in {entry.get('_register_file', 'the register')} was "
                    f"REJECTED and not applied: "
                    + (f"missing key(s) {', '.join(missing)}. " if missing else "")
                    + (f"field {field!r} is outside the permitted set "
                       f"{', '.join(EXCEPTION_FIELDS)}. " if field not in EXCEPTION_FIELDS else "")
                    + ("no system_calculated_value was recorded, so the stale-exception guard "
                       "cannot run. " if "system_calculated_value" not in entry else "")
                    + "Please correct the register entry."
                ),
            })
            continue
        index[_key(entry["reporting_period"], entry["deal"],
                   entry["sba_gp_number"], field)] = entry
    return index


def apply_exceptions(rows, reporting_period, deal, index, servicer_name=None, issues=None):
    """Applies matching exceptions to already-built report rows, in place.

    rows are the report rows for ONE servicer, each carrying the values the
    standard calculation produced. reporting_period is a "YYYY-MM" label; a
    None period or deal matches nothing.

    Every applied exception is surfaced (ISSUE_EXCEPTION_APPLIED) so it can
    never be invisible, per the client's decision that exceptions apply
    automatically but must be prominently reported. Every exception whose
    recorded system value no longer matches is surfaced and NOT applied
    (ISSUE_EXCEPTION_STALE).

    Returns the list of applied entries.
    """
    issues = issues if issues is not None else []
    applied = []
    if not index or reporting_period is None or deal is None:
        return applied

    for row in rows or []:
        gp = str(row.get("SBA GP Number") or "").strip()
        for field in EXCEPTION_FIELDS:
            entry = index.get(_key(reporting_period, deal, gp, field))
            if entry is None:
                continue
            current = row.get(field)
            recorded = entry.get("system_calculated_value")
            try:
                current_value = float(current) if current is not None else None
                recorded_value = float(recorded)
            except (TypeError, ValueError):
                current_value = recorded_value = None

            if current_value is None or recorded_value is None or \
                    abs(current_value - recorded_value) > STALE_TOLERANCE:
                issues.append({
                    "code": ISSUE_EXCEPTION_STALE,
                    "servicer": servicer_name or entry.get("servicer"),
                    "SBA GP Number": gp,
                    "field": field,
                    "reporting_period": reporting_period,
                    "system_calculated_value": current_value,
                    "recorded_system_calculated_value": recorded_value,
                    "message": (
                        f"A recorded reporting exception for GP {gp} ({field}, "
                        f"{reporting_period}) was NOT applied. The value calculated from the "
                        f"lender data is now {current_value!r}, but the exception was recorded "
                        f"against {recorded_value!r}. The source data or the calculation has "
                        f"changed since the exception was confirmed, so the exception no longer "
                        f"describes this loan. The standard calculated value has been reported "
                        f"unchanged. Please re-confirm the exception with the client."
                    ),
                })
                continue

            row[field] = entry.get("authoritative_reported_value")
            applied.append(entry)
            issues.append({
                "code": ISSUE_EXCEPTION_APPLIED,
                "servicer": servicer_name or entry.get("servicer"),
                "SBA GP Number": gp,
                "field": field,
                "reporting_period": reporting_period,
                "deal": deal,
                "system_calculated_value": recorded_value,
                "authoritative_reported_value": entry.get("authoritative_reported_value"),
                "difference": entry.get("difference"),
                "reason": entry.get("reason"),
                "source": entry.get("source"),
                "message": (
                    f"REPORTING EXCEPTION APPLIED. GP {gp} {field} for {reporting_period}: the "
                    f"value calculated from the corrected lender data is "
                    f"{recorded_value:,.2f}; the authoritative reported value is "
                    f"{float(entry.get('authoritative_reported_value') or 0):,.2f} "
                    f"(difference {float(entry.get('difference') or 0):,.2f}). Reason: "
                    f"{entry.get('reason')}. Source: {entry.get('source')}. The standard "
                    f"calculation was NOT changed."
                ),
            })
    return applied


def unmatched_entries(index, reporting_period, deal, rows_by_servicer_gp, issues=None):
    """Register entries for THIS period and deal that matched no loan.

    Surfaced so a stale or mis-keyed entry cannot sit silently in the register
    doing nothing."""
    issues = issues if issues is not None else []
    unmatched = []
    for (period, entry_deal, gp, field), entry in (index or {}).items():
        if period != str(reporting_period or "").strip() or entry_deal != str(deal or "").strip():
            continue
        if gp in rows_by_servicer_gp:
            continue
        unmatched.append(entry)
        issues.append({
            "code": ISSUE_EXCEPTION_UNMATCHED,
            "SBA GP Number": gp,
            "field": field,
            "reporting_period": period,
            "message": (
                f"A recorded reporting exception for GP {gp} ({field}, {period}) matched no "
                f"loan in this report. The loan is not in the reported population. Please check "
                f"the register entry."
            ),
        })
    return unmatched
