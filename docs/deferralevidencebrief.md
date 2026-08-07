# Deferral Evidence — Implementation Brief (rev. against source)

*Now grounded in the actual modules: `corrections.py`, `report.py`, `CLAUDE.md`, and the delivered `Servicer Memo Summary.xlsx`. The decisive logic — STAT-01 in `checks.py` and the memo reader in `file_readers.py` — I have **not** seen; where this brief reaches into those, it says so and turns the reach into a verification step rather than an assertion. Everything about the memo **data** is measured from the file.*

## Correction to the earlier headline

My first read said this was a wiring job, not a parsing job. Having seen the code, that was half right and the wrong half was the dangerous one. The question isn't parsing-versus-wiring for "the dates." It splits cleanly across three different fields, and they fail in three different ways:

- **Deferral dates** (`Deferral Start / End / Resume Date`, memo cols J/K/L) are columnar — real datetimes in 64 of 68 `Deferred` rows. Reading these is a **wiring** job. This part I had right.
- **Deferral type** (full P&I vs partial) has **no column**. It exists only inside the free-text `Servicer Action` cell, and `corrections.py` gates its entire auto-correction on `deferral_type == "full"` (line 432). So the field that decides whether the tool writes a Loan Status change is a **parsing** problem — with real, live failure modes.
- **Deferral phase** (`start` / `mid` / `resume`) is neither read nor parsed; it's **computed** from those dates against the cycle date, and `corrections.py` gates on `deferral_phase in ("start", "mid")` (line 433). Bad dates silently corrupt it.

The reason this matters: the parsing field and the computed field are exactly the two the correction engine trusts to decide what it writes. The reliable columnar field (dates) is the one it never reads directly. So "the dates are clean" was true and almost irrelevant to the risk.

## Where the fixes actually live (now that I've read the code)

`corrections.py` and `report.py` are mostly *ready* and mostly *not the problem*:

- **`report.py` is a dumb formatter.** `_build_detail_sheet` (line 242) takes `result["records"]`, uses the first record's keys as headers, and renders. There is no narrative parsing anywhere in it. So your worry about a reconciliation workbook that makes a reviewer reconstruct a timeline from a paragraph is not a `report.py` change — it's a question of what fields STAT-01 puts into each record. If STAT-01 emits `deferral_start`, `deferral_end`, `deferral_resume`, and a `deferral_anomaly` flag as discrete keys, `report.py` renders them as columns for free. The prose stays a single reference column.
- **`corrections.py` already has the anti-silent-drop pattern you want.** The cross-domain Delinquency Status case (lines 398–411, `_is_cross_domain_dq` → `review_required` → `reviewFlag`) is exactly the shape: a defect is *named and surfaced to the reviewer* instead of vanishing through an eligibility-layer `continue`. The comment at lines 377–383 states the principle explicitly. The deferral-evidence defect classes should reuse this mechanism verbatim, not invent a new one.
- **`corrections.py` already knows how to be protected** — but only if STAT-01 speaks. `WARNING_REASONS` (lines 62–66) includes `"MEMO DATA QUALITY WARNING"` and `"DEFERRAL DATA QUALITY WARNING"`, and any of them excludes a loan from *both* correction types (line 346). The exclusion machinery exists. What's unverified is whether STAT-01 ever *raises* those reasons for the failures below.

So the real work is upstream, in `checks.py` (STAT-01) and `file_readers.py` (the memo reader and its alias maps) — the two files not in hand. That is the honest scope statement, and it's why step one is verification, not construction.

## The silent-failure gap, stated precisely

`CLAUDE.md` line 36 names the mechanism: `MEMO_LENDER_ALIASES` and `SERVICER_SOUP_KEYS` "handle name spelling variants — a missing mapping silently yields zero SOUP matches." Servicer identity for the memo comes from the upload filename convention `{soupKey}||{displayLabel}`, not from in-sheet labels.

The data shows this failing right now. The pipeline sees 198 content rows. (A caution learned the hard way: measuring the sheet with **openpyxl** surfaces ~233 additional empty/`#N/A` rows, but the production **python-calamine** reader — plus `_rows_to_df`'s blank filter — strips those before the pipeline ever sees them. Measure every baseline through `read_service_memo`, not openpyxl; the two readers disagree on empty and error cells, and that disagreement is itself a trap.) Of those 198 content rows:

- **189 (95%) join** to the origination tape on `(Deal, SBA Loan ID)` after ID normalization.
- **9 do not join at all** — they fail even on an ID-only match that ignores the deal, so the loan ID itself is absent from the population. Eight of these carry `Deal = "#N/A"` — the SOUP-key mapping failing exactly as `CLAUDE.md` warns (Oakley Hills, BCRS, World Class Real Estate, Jkari Hospitality, Chico Fine Art — several `Deferred` or `Modification`). One (AD Transportation, `SOUP2025_1`, `5850799108`) has a valid deal but an ID not in the tape.

Here is the gap that makes a blank Deferral Evidence section meaningless. When the memo lookup fails, STAT-01 produces *no deferral data* for that loan. No deferral data means no `correction_data` deferral fields, which means no candidate in `build_eligible_corrections`, which means nothing for `WARNING_REASONS` to exclude — there is nothing there to exclude. The loan sails through looking **identical to a loan with no deferral**. The exclusion set protects loans it can see; it cannot protect a loan whose deferral evidence never arrived. So:

- **no deferral documented** — a fact — and
- **the memo never joined** — a defect —

are today the same output: silence. `corrections.py`'s exclusion machinery does not close this, because the failure happens before a candidate exists. Only STAT-01 raising `MEMO DATA QUALITY WARNING` for an *expected-but-unresolved* memo can close it, and that requires STAT-01 to know a memo was expected for that loan in the first place.

## Step one: extraction verification (before any display or record-shaping)

The verification is a set of assertions about STAT-01's output over these two sheets. Each maps to a failure class the data proves exists, and each has a baseline number to regression against.

1. **Counts the real population.** Assert the pipeline sees **198 rows, all content, 0 phantom** — measured through the production calamine reader, which already strips the empty/`#N/A` rows. (openpyxl would report 431 total / 233 phantom; that reader mismatch is the trap — assert against `read_service_memo`, not a raw openpyxl read.)
2. **Raises `MEMO DATA QUALITY WARNING` for the 9 unresolved joins** — the 8 `#N/A`-deal rows plus AD Transportation. These must reach the reviewer as *lookup-failed*, never render as "no deferral," and never silently vanish. This is the gap above, made testable.
3. **Handles the 4 `Deferred` rows whose `Deferral Start Date` is the string `"n/a"`.** Deferral asserted, dates missing — a third state (*deferral claimed, dates absent*), not "has evidence" and not "no evidence." It should raise `DEFERRAL DATA QUALITY WARNING`, which correctly excludes it from auto-correction *and* surfaces it.
4. **Raises `DEFERRAL DATA QUALITY WARNING` for the 5 impossible date orderings** (resume/end before start): rows 12, 63, 64 (`SOUP2025_1`, all valid deals) and row 131 (`SOUP2024_1`, end *and* resume before start). All four join, so they reach STAT-01, and their dates feed phase classification. A resume date nine months before the start date can flip `deferral_phase` and cause `corrections.py` to auto-recommend, or auto-suppress, a Loan Status change on garbage input — and, per the delinquency-reset rule below, can fire a reset-to-CURRENT at the wrong cycle. These are almost certainly year-entry typos in the source memos — but the pipeline must flag, not compute through them.
5. **Resolves the 6 non-`Deferred` rows that nonetheless carry a structured deferral start date** — a label/date contradiction. Pick a rule (which field wins) and surface the conflict rather than trusting one field silently.

The acceptance test is one function over the two sheets that tags every content row `phantom | joined | lookup-failed | deferral-dates-missing | date-anomaly | label-conflict` and asserts the baseline **measured through the production reader**: 198 total / 0 phantom / 4 dates-missing / 5 anomaly across 4 rows / 6 conflict (all population-independent), plus — against the **real SOUP** population — **190 joined / 8 lookup-failed**. (These two were 189/9 against the OriginationTapes proxy; the confirmed real-SOUP figures are 190/8 — one memo loan, AD Transportation, exists in SOUP but not the proxy, so it joins. Assert joined/lookup-failed only in the SOUP-backed test.) Confirmed green against the real files on 2026-08-07. Until it's green, the Deferral Evidence section has no defined meaning and shouldn't be built.

## The parsing field with teeth: full vs partial

This is the finding that most changes the work. `corrections.py` auto-recommends `Loan Status → DEFERRED` only for `deferral_type == "full"` (line 432); partial (principal-only / interest-only) is deliberately out of v1 scope. But `deferral_type` is derived from the free-text `Servicer Action` cell, and across the 68 `Deferred` rows that text is genuinely messy — 51 distinct phrasings. A naive keyword classifier lands **29 full-P&I, 18 principal-only, 1 interest-only, and 20 unclassifiable** (30%). Samples of the danger:

- `"3-month deferment of principal paying interest"` — principal deferred, **interest still paid** → this is *partial*, but a matcher that fires on "principal" + "interest" both present reads it as full P&I. A **false-full** here makes the tool auto-write a DEFERRED status change on an out-of-scope loan.
- `"2-month deferment with $7200 interest payments"` — interest being paid → partial, same trap.
- `"additional 6M of the interest only term"` — an I/O term extension; arguably not a deferral event at all. Classifying it as one is wrong outright.
- `"90-Day Deferment"`, `"One-Time 90-Day Deferment"`, `"3 month payment deferral"` — type simply unspecified. These must be *unclassifiable → review*, never defaulted to full.
- Spelling variance a matcher must survive: `P&I`, `P and I`, `PNI`, `P & I`, `principal and interest`, `full P&I`, `principal & interest`.

The safe posture mirrors the rest of the module: classify to `full` only on an explicit, unambiguous full-P&I signal; everything else is `partial` or `unclassified`, and `unclassified` raises `DEFERRAL DATA QUALITY WARNING` so it's excluded from auto-correction and surfaced. The failure to avoid at all costs is a false-full, because that is the one that silently writes.

## The deferral lifecycle and the delinquency reset (client-established rule)

Add this explicitly, because it defines what STAT-01's delinquency ladder must do across the deferral boundary and it is not derivable from the memo data alone:

> **While the loan is in deferral:** use normal delinquency logic.
> **At the Regular Payment Resume Date (deferral exit):** reset Delinquency Status to CURRENT.
> **After resume:** continue normal delinquency logic, starting from CURRENT.

Mapped onto the existing `deferral_phase` field this is exact and needs no new vocabulary:

- `start` / `mid` — *in deferral* — normal delinquency logic runs unchanged. Deferral does **not** suppress the ladder while it's active.
- `resume` — the exit event, triggered by the `Regular Payment Resume Date` (memo col L) — Delinquency Status is forced to CURRENT before the ladder runs for that cycle, and no days-past-due are carried across the boundary.

Two consequences have to be stated so they aren't lost between modules.

**This is a STAT-01 calculation rule, and it does *not* conflict with "resume is out of v1 scope."** `corrections.py` excludes the resume phase from *Loan Status* auto-correction (line 433) — correct, because a loan exiting deferral shouldn't be auto-written to DEFERRED. But the delinquency reset is a *different field* computed in STAT-01, and it feeds two things that are in scope: the v1 Delinquency Status auto-correction (the unambiguous, warning-free path) and the Portfolio Summary aging buckets. If STAT-01 doesn't perform the reset, then at and after resume it carries forward days-past-due it should have zeroed, `calculated_dq` is wrong, and `corrections.py` will faithfully auto-write that wrong value. The reset is therefore a **precondition for an in-scope correction to be correct**, not an optional resume-phase nicety — and anyone reading "resume is out of scope" must not skip it. To be unambiguous: **the reset applies only to Delinquency Status.** Loan Status continues to follow its own business rules and review workflow — "reset to CURRENT" never refers to Loan Status.

**The reset trigger is the exact field the date anomalies corrupt.** The Regular Payment Resume Date is memo col L — the field with the 5 impossible values (resume before start: rows 12 / 63 / 64 / 131). A resume date in the past would fire the reset-to-CURRENT at the wrong cycle, or retroactively. So step one's date-anomaly gate (item 4) is load-bearing for *this* rule specifically: STAT-01 must refuse to apply the reset on a resume date that failed its sanity check, raise `DEFERRAL DATA QUALITY WARNING` instead, and that warning also excludes the loan from auto-correction until a human fixes the date.

Two definitions to pin before coding, since the rule as written is silent on both:

- **Cycle timing — already resolved in code.** On review of `checks.py`, this is not an open assumption: `_lookup_deferral` sets the `resume` phase when `cycle_period == chosen_resume_month` (the cycle *containing* the Regular Payment Resume Date), and `_calculate_expected_status` resets the baseline to CURRENT there (lines ~1123–1130, ~1341–1354). The module comment marks it `CLIENT-CONFIRMED` and states explicitly that Regular Payment Resume Date — not Deferral End Date — is the controlling field. Note this is the *containing-cycle* rule, **not** the "first cycle on or after" I had proposed earlier as a default; the code's client-confirmed choice supersedes that proposal. No further confirmation needed unless the client wants to revisit the confirmed decision.
- **Missing resume date.** A blank or `"n/a"` resume date leaves the reset boundary undefined — treat as `DEFERRAL DATA QUALITY WARNING`, never guess a date.

## Two output paths, one contract

The review payload (`corrections.build_recommendations`) and the reconciliation workbook (`report.py`) are separate code over the same facts; specify the shared field once so they can't drift. Every deferral-evidence record carries an `evidence_status` — one of `documented`, `no-deferral`, `deferral-claimed-dates-missing`, `lookup-failed`, `date-anomaly`, `type-unclassified` — computed once in STAT-01, consumed by both renderers, recomputed by neither.

- `build_recommendations` surfaces `lookup-failed`, `date-anomaly`, and `type-unclassified` through the **existing** `review_required` / `reviewFlag` path (the one built for cross-domain DQ), so a defect reaches the reviewer named rather than dropped.
- `report.py` needs no narrative logic — give its records discrete `deferral_start / end / resume` and `evidence_status` keys and the columnar timeline renders itself, prose demoted to one reference column.

**Reviewer context on every deferral record (eliminates the cross-workbook lookup).** Every deferral-related STAT-01 record — not only the exception rows — must carry, as discrete columns: **Deferral Start Date, Deferral End Date, Regular Payment Resume Date, Prior SOUP Loan Status, Current Servicer Loan Status, and Memo Securitization Action.** Those six together explain, at a glance, *why* a row was flagged — most importantly a label conflict (a `Modification` Securitization Action against a full-P&I window, GP 6673679110's row 49), which is unreadable without seeing the memo label beside the prior/current statuses. STAT-01 is the only place all six coexist for one loan, which is precisely why this belongs here and not in the memo diagnostic (`memo_audit.py` has neither SOUP nor servicer status). The record already carries Current Servicer Loan Status (`srf_status`); the five additions require `_lookup_deferral` to return the chosen memo row's raw `Deferral Start/End/Resume` dates and its `Securitization Action` (today it returns only `deferral_type`/`phase`/`months_since_start`), plus surfacing `soup_status` as its own labeled column rather than only folded into "Expected Loan Status." This is where the reviewer's manual investigation disappears.

There is a dollar consequence that raises the stakes: `report.py`'s Portfolio Summary has a `Deferred` bucket (line 297), and the Original-vs-Audited pair differs only in "which Delinquency Status source feeds the aging buckets." A silently-unrecognized deferral misbuckets its balance into a delinquency band — the same failure mode `corrections.py` already documents costing 584,023.61 unbucketed on one mistyped status (lines 78–81). Silent memo misses aren't just a blank cell; they move money into the wrong bucket.

## Sequence

The trace of GP 6673679110 reorders this away from the original "parse dates, then fix the resume logic." The resume logic is already correct; the problem is upstream — valid deferral evidence is excluded before STAT-01 ever gets to apply the lifecycle. So the sequence now verifies and *protects* existing logic before changing any rule.

1. **Verify memo lookup and alias mapping.** Audit `MEMO_LENDER_ALIASES` / `SERVICER_SOUP_KEYS` in `file_readers.py`: resolve the 8 `#N/A` lenders, and make an *expected-but-unmapped* memo a loud report line rather than a silent zero-match. Stand up the partition/baseline counts **measured through the production calamine reader** (198 total / 0 phantom / 4 dates-missing / 5 anomaly across 4 rows / 6 label-conflict, all population-independent; plus **190 joined / 8 lookup-failed** against the real SOUP population — 189/9 was the OriginationTapes proxy) as a test fixture here — this is the measurement of whether lookup and keying actually work, and it becomes the regression guard in step 6. Do not measure the baseline with openpyxl; it disagrees with calamine on empty/`#N/A` rows (it would show 431/233), and the test must match what the pipeline actually reads.

2. **Investigate why valid deferral rows are excluded by the `Securitization Action == "deferred"` filter (`checks.py:1068`).** GP 6673679110 proved a genuine full-P&I deferral (row 49) is dropped because it's labeled `Modification`. Scope it first: how many other loans carry a complete, internally-consistent deferral window and full-P&I Servicer Action text under a non-`Deferred` action? That count decides whether this is an edge case or a systemic hole.

   > **Proposed investigation outcome (not an approved fix).** If a non-`Deferred` row contains all of: internally-consistent deferral dates, full-P&I Servicer Action text, and overlap with the audit cycle — then classify it as a **data-quality conflict requiring reviewer attention** rather than silently excluding it. **Do not automatically treat it as Deferred without confirming the intended business rule.** The safe move is to expose the conflict for a human, not to assert the deferral.

3. **Raise data-quality warnings instead of silently dropping conflicting rows.** The warning infrastructure already exists (`_lookup_deferral`'s `lender_identity` / `deferral_dates` → `MEMO`/`DEFERRAL DATA QUALITY WARNING` → `corrections.py`'s `WARNING_REASONS` + `review_required` path). Extend it to the two currently-silent cases: the line-1068 label conflict (step 2), and an ambiguous `deferral_type` — `classify_deferral_type` defaults ambiguous text to `full`, the direction that can enable an unwanted auto Loan-Status→DEFERRED write, so surface ambiguity as a warning rather than defaulting through it. Define the shared `evidence_status` contract here — `documented | no-deferral | lookup-failed | date-anomaly | type-unclassified` — computed once in STAT-01 and consumed by `corrections.py`, the review page, and `report.py`, so none of the three invents its own interpretation.

4. **Verify the existing resume reset still works once the data reaches STAT-01.** The reset is implemented and client-confirmed (lines 1341–1354); the task is to confirm it end-to-end for loans that now actually reach it — including a rolling deferral where resume(A) = start(B) — and to add the one guard it lacks: a date-anomaly gate so a resume date *before* its start date raises `DEFERRAL DATA QUALITY WARNING` instead of being computed through.

5. **Only then evaluate whether any business-rule change is still needed.** With lookup verified, exclusions surfaced, and the resume reset confirmed, decide deliberately — with the client — whether a rule actually needs to change (e.g. widening deferral recognition beyond the `Securitization Action` label, or revisiting the confirmed containing-cycle resume timing). Rule changes come last, as confirmed decisions, never as assumptions baked in earlier.

6. **Regression guard.** The partition counts from step 1 become a test fixture. When the next memo drop moves them, the diff is the signal — a jump in `lookup-failed` means an alias map broke; a jump in `label-conflict` means more deferrals are being filed under the wrong action — and you see it instead of shipping it.

## Status of source review

`checks.py`, `file_readers.py`, `corrections.py`, and `report.py` have all now been read; the inferences that earlier stood in for STAT-01 and the memo reader are replaced with line-level behavior throughout. The one remaining gap is *data*, not code: to evaluate a specific loan/cycle numerically (rather than structurally) — e.g. to show the GP 6673679110 Jan–Mar divergence as concrete `calculated_dq` values — the SOUP dataset and the servicer SRF submission for that cycle are needed, since they carry the cycle month, SOUP's last-recorded statuses, and the payment amounts.
