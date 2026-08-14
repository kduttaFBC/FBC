# OPEN QUESTION — the current-period Bankruptcy Date branch is gated on LOAN STATUS MISMATCH

**Status: RECORDED, NOT FIXED. Requires client confirmation before any change.**
**Discovered 2026-08-14 on GP 6666709108, Lendistry 2025, July 2026.**
No code was changed in response to this. It is written down so it is not
rediscovered by accident later.

## The rule as the client stated it

CLIENT-CONFIRMED (2026-08-12): BANKRUPTCY is valid for a reporting period when
**that period's own Bankruptcy Date column contains a date**. Prior SOUP
BANKRUPTCY alone is explicitly *not* sufficient and plays no part.

## What the code actually requires

`corrections.build_eligible_corrections`, the current-period Bankruptcy Date
branch, opens with:

```python
if (
    "LOAN STATUS MISMATCH" in reasons
    and "Loan Status" not in _emitted_fields
    and _bankruptcy_dated
    and _balance_outstanding
    and not _finished_loan
    and not _is_bankruptcy_status(c.get("srf_status"))
):
```

`"LOAN STATUS MISMATCH" in reasons` is raised by STAT-01 when the servicer's
reported Loan Status disagrees with the audit's expected status, which is
derived from **prior SOUP**. So the populated current-period Bankruptcy Date is
consulted **only after** a prior-vs-reported disagreement already exists.

The consequence: **a loan whose prior finalized status AGREES with its reported
status produces no recommendation at all, however clearly its current period is
marked as a bankruptcy.**

## The case that exposed it

GP **6666709108**, July 2026, Lendistry 2025 — reported `ACTIVE`, current-period
**Bankruptcy Date 2026-04-06**, Ending Balance **$70,465.59** (live), and a
**Liquidation Date of 05/26/2026** against that live balance.

| Prior finalized period | Prior SOUP status | Reported | LOAN STATUS MISMATCH? | Recommendation |
|---|---|---|---|---|
| May 2026 (wrong pairing) | `ACTIVE` | `ACTIVE` | no | **none — invisible** |
| June 2026 (correct pairing) | `BANKRUPTCY` | `ACTIVE` | yes | `ACTIVE` → `BANKRUPTCY`, Review Required |

Under the correct pairing the loan surfaces, so **this is not currently harming
production**. What it demonstrates is that surfacing it depended on prior SOUP
disagreeing — not on the Bankruptcy Date, which is what the stated rule keys on.

## Why this may still matter

Construct the case where prior and reported agree *and* the period is a
bankruptcy: a loan is `ACTIVE` in the prior finalized SOUP, the borrower files in
the current period, the servicer populates Bankruptcy Date but has **not yet**
changed Loan Status (still `ACTIVE`), and the balance is live.

- Reported `ACTIVE` == prior `ACTIVE` → no `LOAN STATUS MISMATCH`
- Bankruptcy Date populated, balance positive, not terminal, not already BANKRUPTCY
- Gates 1–4 of the client's rule are **all satisfied**
- Yet **no recommendation is produced**, because the branch never runs

That is precisely the "silently invisible loan" condition the branch was added
on 2026-08-12 to end. Note this is the *first* month of a bankruptcy — the month
the servicer is most likely to have updated the date but not the status.

## Not yet measured

Whether any such loan exists in the available data. GP 6666709108 does not
qualify — under the correct pairing its prior status disagrees, which is why it
surfaces. Confirming or excluding the condition needs a sweep across periods for
loans where `bankruptcy_date_populated` is true, the balance is positive, the
reported status is not BANKRUPTCY, and prior == reported. **Not run yet.**

## Options, if the client confirms it should be fixed

1. **Drop the `LOAN STATUS MISMATCH` gate for this branch only.** The other four
   gates already constrain it, and it can only ever propose BANKRUPTCY on a
   positive-balance, non-terminal, non-BANKRUPTCY loan whose current period
   carries the date. Widest fix; needs care that STAT-01 still produces
   `correction_data` for a loan carrying no reasons at all — **it currently does
   not**, so this option is larger than it first appears.
2. **Have STAT-01 raise a distinct reason** (e.g. `CURRENT PERIOD BANKRUPTCY
   DATE`) when the date is populated on a live, non-terminal, non-BANKRUPTCY
   loan, and admit that reason to the branch alongside `LOAN STATUS MISMATCH`.
   Narrower and keeps `reasons` as the single routing vocabulary.
3. **Leave as is** and accept that the branch is a *contradiction* detector
   rather than a bankruptcy detector — i.e. the client's rule is understood to
   apply only where the reported status is already in dispute.

Option 2 looks least invasive, but **all three change what recommendations
exist**, which is a business rule change and out of scope for the Human Review
organization work.

## Related, already-recorded

`_PROPOSABLE_LOAN_STATUS` deliberately excludes the terminal collection
statuses, so the ACTIVE→LIQUIDATION direction produces no recommendation either.
That is a separate, already-flagged item, kept out of scope by the same
reasoning.
