# OPEN QUESTION — the current-period Bankruptcy Date branch is gated on LOAN STATUS MISMATCH

**Status: MEASURED 2026-08-14 — NO MISSED CASE FOUND. NOT FIXED, and no change
is proposed. Requires client confirmation before any change.**
**Discovered 2026-08-14 on GP 6666709108, Lendistry 2025, July 2026.**
No code was changed in response to this, before or after measurement. It is
written down so it is not rediscovered by accident later.

CLIENT-CONFIRMED (2026-08-14), after the sweep recorded below: **the existing rule
currently functions as a LOAN STATUS CONTRADICTION detector**, and the production
Bankruptcy Date logic is to be left alone. What follows documents behaviour; it
does not propose a rule.

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

## MEASURED — 2026-08-14

Swept read-only across every consolidated loan in both periods, with the correct
reporting-period pairing preserved (June audit ← May-latest SOUP; July audit ←
June-inclusive SOUP). Each gate of the branch was evaluated independently per
loan, so the reason a loan was or was not surfaced is recorded rather than
inferred.

**Population: 3,326 loan-periods** — June 1,779 loans across 8 servicers, July
1,547 across the 4 servicers whose source SRFs exist. App-generated CORRECTED
workbooks were deliberately excluded as inputs (they are outbound artifacts;
preflight FORM-03).

| | June (prior = May) | July (prior = June) |
|---|---|---|
| Loans swept | 1,779 | 1,547 |
| Bankruptcy Date populated | **6** | **6** |
| …with a live (nonzero) Ending Balance | **6** | **6** |
| …and reported status ≠ BANKRUPTCY (criteria 1–3) | **0** | **5** |
| …and prior finalized status == reported (the gap condition) | **0** | **0** |
| **Loans missed by Human Review** | **0** | **0** |

**ACROSS ALL 3,326 LOAN-PERIODS, NO ACTUAL EXAMPLE WAS FOUND WHERE THE MISMATCH
GATE CAUSED A BANKRUPTCY DATE CASE TO BE MISSED.**

### June — 6 bankruptcy-dated loans, 0 missed

June is not an empty result; the sweep has signal. All six loans carry a
Bankruptcy Date on a live balance, and all six are excluded for one reason:
**every one already reports `BANKRUPTCY`**, so there is nothing to propose
(`_is_bankruptcy_status(srf_status)` is the gate that excludes them).

| GP | Reported | Prior (May) | Ending Balance | BK Date | LIQ Date |
|---|---|---|---|---|---|
| 5917099105 | BANKRUPTCY | BANKRUPTCY | $10,489.57 | 2026-02-15 | — |
| 6288639102 | BANKRUPTCY | BANKRUPTCY | $55,907.68 | 2026-03-23 | — |
| 6649779105 | BANKRUPTCY | BANKRUPTCY | $7,450.69 | 2026-03-12 | 2026-06-29 |
| 6733679102 | BANKRUPTCY | **ACTIVE** | $14,106.23 | 2026-06-10 | — |
| 6666709108 | BANKRUPTCY | **ACTIVE** | $70,465.59 | 2026-04-06 | — |
| 7126119101 | BANKRUPTCY | BANKRUPTCY | $71,252.97 | 2026-02-28 | — |

`6733679102` and `6666709108` are the important two: the lender reported
`BANKRUPTCY` against a prior `ACTIVE` — it recorded the date **and** updated the
status in the same submission. That is the behaviour the gap theory worries about,
happening correctly, and it is why no recommendation is needed.

### July — 6 bankruptcy-dated loans, 5 non-BANKRUPTCY, all 5 surfaced, 0 missed

All five passed **all eight gates**; the exact reason in each case is *"all gates
pass → recommendation emitted."* Each is Review Required, category `Loan Status`,
never bulk-approvable, and each pairs with a `Standard — Delinquency` companion.
Memo evidence `no-deferral` on all five. No Foreclosure Date or REO Date is
populated anywhere in either period.

| GP | Prior LS | Reported | Prior Bal | Cur Bal | BK Date | LIQ Date | Calculated | Surfaced |
|---|---|---|---|---|---|---|---|---|
| 5917099105 | BANKRUPTCY | ACTIVE | 10,489.57 | 10,489.57 | 2026-02-15 | — | BANKRUPTCY | yes |
| 6288639102 | BANKRUPTCY | ACTIVE | 55,907.68 | 55,907.68 | 2026-03-23 | — | BANKRUPTCY | yes |
| 6649779105 | BANKRUPTCY | ACTIVE | 7,450.69 | 7,450.69 | 2026-03-12 | **06/29/2026** | BANKRUPTCY | yes |
| 6666709108 | BANKRUPTCY | ACTIVE | 70,465.59 | 70,465.59 | 2026-04-06 | **05/26/2026** | BANKRUPTCY | yes |
| 7126119101 | BANKRUPTCY | ACTIVE | 71,252.97 | 71,252.97 | 2026-02-28 | — | BANKRUPTCY | yes |

The sixth, `6733679102`, already reports `BANKRUPTCY` (prior BANKRUPTCY too) and
correctly produces nothing.

### What the five July loans actually are

All five are **historical carry-forward on an already-known bankruptcy**, not
current-period discoveries: every prior finalized status is `BANKRUPTCY` and every
Bankruptcy Date precedes the July cycle (Feb–Apr 2026). What the lender did was
**regress** the status from `BANKRUPTCY` back to `ACTIVE` while the date persisted.

So the branch is catching a lender data regression, which is a contradiction —
confirming option 3 below as a description of current behaviour, not merely as a
fallback position.

**Genuine first-time current-period evidence: 0 in July, 2 in June (`6733679102`,
`6666709108`) — and in both the servicer reported `BANKRUPTCY` itself.** No loan
in either period has a first-time bankruptcy the servicer failed to reflect.

## The remaining theoretical exposure — a ONE-PERIOD DETECTION LAG

The mechanism is real and unchanged; only its incidence is zero. Two structural
reasons make it unlikely rather than merely absent:

1. When a servicer records a Bankruptcy Date it generally updates Loan Status in
   the same submission (June's `6733679102` and `6666709108` both did). That
   creates a mismatch against the prior `ACTIVE` and surfaces the loan anyway.
2. If it does **not** update the status, the mismatch appears the FOLLOWING month
   once SOUP carries `BANKRUPTCY` — which is exactly the five July loans.

**The gate therefore delays detection by one reporting period; it does not
eliminate it.** The residual exposure is the narrow case where a servicer records
the date, leaves the status alone, AND the prior period also happens to agree. No
such loan exists in the available data.

This is recorded as a known, accepted, measured-at-zero exposure. **No production
logic is to change on the strength of it.**

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

4. **Add a non-blocking ADVISORY** (not a recommendation) when the date is
   populated but no mismatch exists — visibility without changing what is
   correctable or what D4 gates. Needs a new display concept: there is no
   existing channel for a loan-level advisory outside `reasons`.

**Current implementation decision: leave production logic unchanged based on the
measured results. Option 3 describes current behavior. Options 1, 2, and 4 remain
available if the client later clarifies that Bankruptcy Date should independently
trigger Human Review.**

Supporting notes: the measurement found nothing missed, and options 1, 2 and 4 all
change what recommendations exist, which is a business rule change. Option 1 is
additionally larger than it looks, since STAT-01 does not emit `correction_data`
for a loan carrying no reasons at all. **No production logic changed.**

## Client clarification still outstanding

These are open questions about INTENT. None is a defect, and none should be acted
on without confirmation.

1. **Is a Bankruptcy Date whose date precedes the current period still
   "current-period evidence"?** All five July loans rely on Feb–Apr dates against
   a July cycle. The 2026-08-12 rule said *"this period's own Bankruptcy Date
   column contains a date"* — the column is populated, but the date is not in the
   period. This is the most consequential ambiguity, and it affects loans that are
   already surfacing today.
2. **Should the rule apply when the reported status is not in dispute?** i.e. is
   the intent a bankruptcy detector or a contradiction detector? Current behaviour
   is the latter.
3. **Is the one-period detection lag acceptable** where a servicer records the
   date but not the status?
4. **How should a `BANKRUPTCY` → `ACTIVE` regression be treated** — as a lender
   data-quality defect worth notifying the servicer about, or purely as a status
   correction? All five July loans are this shape.

## SEPARATE OPEN CLIENT QUESTION — Liquidation Date against a live balance

**Recorded, NOT a rule. No Liquidation business rule is to be derived from this
evidence yet.**

Two loans report a Liquidation Date while still carrying a positive Ending
Balance, and both are in July 2026, Lendistry 2025:

| GP | Liquidation Date | Ending Balance | Reported LS | Prior LS | Also carries |
|---|---|---|---|---|---|
| **6649779105** | 06/29/2026 | **$7,450.69** | ACTIVE | BANKRUPTCY | Bankruptcy Date 2026-03-12 |
| **6666709108** | 05/26/2026 | **$70,465.59** | ACTIVE | BANKRUPTCY | Bankruptcy Date 2026-04-06 |

Both currently surface in Human Review — as `Loan Status` recommendations proposing
`BANKRUPTCY`, carrying
`REVIEW_REASON_BANKRUPTCY_LIQUIDATION_CONTRADICTION` — so the contradiction is
shown to the reviewer and resolved by a human. That is the whole of the present
behaviour and it is working.

What is NOT decided, and must not be inferred from these two loans:

- Whether a Liquidation Date on a live balance should on its own support a
  `LIQUIDATION` Loan Status. It currently does **not**: the application cannot
  generate a terminal collection status at all (`_PROPOSABLE_LOAN_STATUS` is
  `{ACTIVE, DEFERRED, PAID-IN-FULL}`), so `LIQUIDATION` can only ever arrive by
  explicit reviewer selection.
- Whether a Liquidation Date with a nonzero balance is a lender data-quality
  defect (the date is wrong, or premature) rather than a status signal.
- Which of Bankruptcy Date and Liquidation Date should take precedence when both
  are populated in the same period on a live balance. The client's zero-balance
  exception does not apply here, because it is keyed on an Ending Balance of $0.
- Whether Liquidation Proceeds / Liquidation Loss / Subsequent Recoveries should
  participate. They are now carried to the review page as display evidence, and
  are **empty on both of these loans**.

**Two loans is not a population.** Deriving a Liquidation rule from them would be
inventing a business rule from two data points, which is exactly what this section
exists to prevent.

## Related, already-recorded

`_PROPOSABLE_LOAN_STATUS` deliberately excludes the terminal collection
statuses, so the ACTIVE→LIQUIDATION direction produces no recommendation either.
That is a separate, already-flagged item, kept out of scope by the same
reasoning — and it is the reason the two loans above propose `BANKRUPTCY` rather
than `LIQUIDATION`.
