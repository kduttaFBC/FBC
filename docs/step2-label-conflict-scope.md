# Step 2 — Scoping the `Securitization Action == "deferred"` filter (`checks.py:1068`)

*Investigation only. No code changed, no business rule changed, no `Modification` row treated as `Deferred`. Every figure below is measured from the delivered `Servicer Memo Summary.xlsx` and `SOUP_Performance_DatasetALL.xlsx` through the production reader (`read_service_memo`), with `checks._lookup_deferral` called unmodified.*

## Question

Is GP 6673679110 an isolated data-quality conflict, or evidence of a broader hole in the label filter?

## Answer

**Isolated, and — at the cycle actually being audited — with no impact on the result.** Of the six label conflicts, exactly one carries genuine full-P&I deferral evidence under a non-`Deferred` label, and its window is historical: a correctly-labelled `Deferred` row covers the current cycle for the same loan. Meanwhile the filter is actively suppressing two administrative rows *and doing so correctly* — removing it would make those two loans worse, not better.

**No rule change is supported by this evidence.** Confirmed against all 24 cycles, not just the current one — see *Decision evidence* below. The filter stays as written; no production code was changed.

## Method

- Audit cycle per `checks._infer_cycle_month` (SOUP's latest Payment Month + 1). SOUP runs through **May 2026**, so **every servicer's cycle is June 2026**.
- Overlap uses `_lookup_deferral`'s own `still_relevant` test: `start_month <= cycle` and (`resume` unknown or `cycle <= resume`).
- Suppression measured by a **data** experiment, not a code change: a copy of the memo frame with one row's `Securitization Action` relabelled `"Deferred"`, passed to the unmodified `_lookup_deferral`. It measures what the filter withholds without proposing that the filter change.

## Population context

Of 198 content rows: 114 charge-offs (`Charge-off` 107 + `Charge Off` 7), 68 `Deferred`, 11 `Modification`, 3 `Change Memo #N`, 1 `CHANGE IN TERMS AGREEMENT`, 1 blank.

**70 rows carry a parsed Deferral Start Date — 64 `Deferred` and 6 not.** Those 6 are the complete conflict set. Note 8 of the 11 `Modification` rows carry no deferral dates at all, so they are not conflicts in any sense.

## The six conflicts

| # | Row | GP | Securitization Action | Servicer Action | Type | Start / End / Resume | Overlaps Jun-2026? | Suppressed now? |
|---|-----|----|-----------------------|-----------------|------|----------------------|--------------------|-----------------|
| 1 | 45 | 6673679110 | `Modification` | `3-month full principal and interest deferral` | full | 2026-01-01 / 2026-03-01 / 2026-04-01 | **No** | No |
| 2 | 48 | 7135239104 | `Modification` | `One-Time 90-Day Deferment` | full\* | 2025-10-01 / 2025-12-01 / 2026-01-01 | No | No |
| 3 | 94 | 3483589102 | `Modification` | `Monthly payments will be changed to interest only payments` | **partial** | 2025-08-07 / — / 2026-05-07 | No | No |
| 4 | 158 | 1329479107 | `Change Memo #3` | `Change Memo #3` | full\* | 2026-03-01 / 2026-04-01 / — | Yes | **Yes — correctly** |
| 5 | 159 | 6536029107 | `Change Memo #1` | `Change Memo #1` | full\* | 2026-03-05 / 2026-04-05 / — | Yes | **Yes — correctly** |
| 6 | 162 | 5520789102 | `CHANGE IN TERMS AGREEMENT` | `CHANGE IN TERMS AGREEMENT` | full\* | 2026-03-01 / 2026-05-01 / 2026-06-01 | Yes | No — handled by the CIT exception |

\* `full` here is `classify_deferral_type`'s **ambiguous default**, not a positive signal — see "Collateral finding" below.

**Exactly one of six carries explicit full-P&I text** (`principal and interest` / `P&I` / `PNI` / …): row 45. The other five have no P&I signal at all, and row 94 explicitly says interest-only.

### 1 — GP 6673679110, The Nomad Group (the primary case)

Genuine full-P&I deferral filed as `Modification`. Two memo rows:

| Row | Action | Window | Resume |
|---|---|---|---|
| 45 | `Modification` | 2026-01-01 → 2026-03-01 | 2026-04-01 |
| 179 | `Deferred` | 2026-04-01 → 2026-06-01 | 2026-07-01 |

At the June 2026 cycle, row 179 covers the cycle and `_lookup_deferral` returns `phase='mid'`, `deferral_type='full'`. Relabelling row 45 to `Deferred` produces the **identical** result. SOUP's last recorded status for the loan is already `DEFERRED`.

So the mislabel is real, but its consequence is **historical**: it would have mattered for the Jan / Feb / Mar 2026 cycles, exactly as `docs/trace6673679110.md` describes. It does not affect the cycle being audited. This is a data-quality conflict for a human to correct at source, not a live miss.

### 2 — GP 7135239104, POW PRO CONSTRUCTION

Window closed January 2026, five months before the cycle. The loan is `PAID-IN-FULL` in SOUP and carries a later `Charge-off` memo row. Type is unspecified in the text. Nothing suppressed, nothing at stake.

### 3 — GP 3483589102, Shree Austbhuia

`Monthly payments will be changed to interest only payments` — a permanent restructuring, not a temporary deferral, and no Deferral End Date. Correctly excluded on two independent grounds: it isn't labelled `Deferred`, and `classify_deferral_type` reads it as `partial` (out of v1 correction scope regardless). Window closed May 2026.

### 4 and 5 — `Change Memo #3` / `Change Memo #1` — the filter earning its keep

These are the exact case the filter's comment (`checks.py:1060-1066`) describes: administrative rows that **copy a real deferral's dates** but are not deferrals.

- GP 1329479107 also has row 145, `Deferred`, same 2026-03-01 → 2026-04-01 window, **resume 2026-05-01**.
- GP 6536029107 also has row 11, `Deferred`, same 2026-03-05 → 2026-04-05 window, **resume 2026-05-05** (plus a future row 192, Jul–Sep 2026).

Both real deferrals resumed *before* the June cycle, so today both loans correctly return no active deferral. The Change Memo rows have a **blank resume date**, so if the filter admitted them they would be "still relevant" forever, win the most-recent-Report-Date tiebreak, and yield `phase='resume_unverified'` with a `DEFERRAL DATA QUALITY WARNING` — a fabricated warning and a fabricated deferral state on two loans that are fine. Measured, not hypothesised.

**Widening the filter would actively damage these two loans.**

### 6 — GP 5520789102, `CHANGE IN TERMS AGREEMENT`

Already handled by the client-confirmed CIT resume exception, which runs *before* the label filter. It returns `phase='resume'` with `resume_via_change_in_terms_exception=True`. Relabelling to `Deferred` also gives `phase='resume'` — the same answer by a different route. No gap.

## Verdict

| Category | Count | Rows |
|---|---|---|
| Genuine full-P&I evidence under a non-`Deferred` label | **1** | 45 |
| Partial / interest-only restructuring | 1 | 94 |
| Type unspecified, window long closed | 1 | 48 |
| Administrative duplicates the filter exists to reject | 2 | 158, 159 |
| Already handled by the CIT exception | 1 | 162 |

**Missed deferrals at the June 2026 cycle: zero.** The filter suppresses three rows at this cycle; in two cases the suppression prevents a false warning, and in the third the CIT exception already produces the right answer.

---

# Decision evidence — all 24 cycles

One cycle is too thin a basis for a permanent rule decision, so the same counterfactual was swept across **every month from Jan-2025 to Dec-2026** for all six conflicts. This is the evidence the decision rests on; the June-2026 figures above are one column of it.

| Row | GP | Cycles where the filter changes the outcome | Effect of admitting the row | Sole evidence? |
|-----|----|---------------------------------------------|------------------------------|----------------|
| 45 | 6673679110 | 3 (Jan–Mar 2026) | adds `full` deferral, `start`/`mid`/`mid` | No — row 179 covers Apr onward |
| 48 | 7135239104 | 4 (Oct 2025–Jan 2026) | adds `full` deferral (type unspecified in text) | Yes |
| 94 | 3483589102 | 10 (Aug 2025–May 2026) | adds a **`partial`** record, `phase=None` | Yes |
| 158 | 1329479107 | 8 (May–Dec 2026) | **fabricates** `resume_unverified` + DQ warning | No — row 145 |
| 159 | 6536029107 | 3 (Jun, Nov, Dec 2026) | **fabricates** `resume_unverified` + DQ warning | No — rows 11, 192 |
| 162 | 5520789102 | 3 (Mar–May 2026) + 1 changed | adds `full` deferral in-window | Yes, but CIT-governed |

Totals: 20 cycle-instances where admitting a row adds a deferral record, 11 where it fabricates a data-quality warning.

**The 20 do not decompose into 20 hidden deferrals.** Read honestly:

- **10 of them (row 94) are not a full deferral at all.** Relabelled, `_lookup_deferral` returns `{'deferral_type': 'partial', 'phase': None}` — an interest-only restructuring with no end date, which is a permanent change in terms, not a temporary payment freeze. Admitting it would assert a deferral the memo never claims.
- **3 (row 162) are governed by an existing client-confirmed decision.** The CIT exception is deliberately resume-only and explicitly never establishes a *new* deferral (`checks.py:876-885`). Its in-window months are excluded by design, not by oversight. Revisiting that is a client question, not a defect.
- **4 (row 48)** are a genuine but type-unspecified deferment (`One-Time 90-Day Deferment` — `full` only by `classify_deferral_type`'s ambiguous default), on a loan since paid in full and charged off. Window closed January 2026.
- **3 (row 45)** are the one genuine full-P&I case, historical, with a correctly-labelled row covering every later cycle.

Against that, the **11 fabricated warnings are unambiguous harm**, and one is worse than a false positive: at the May-2026 cycle, GP 1329479107 today returns a correct `resume`, and admitting the Change Memo row **degrades a correct answer** into `resume_unverified` plus a warning.

### The narrow rule the brief proposed

> *internally-consistent deferral dates + full-P&I Servicer Action text + overlap with the audit cycle → surface as a conflict.*

Measured against the file, the rows satisfying all three conditions number **one** (row 45) — and at the current cycle, **zero**, because its window closed in March. Even the narrowest widening fires on nothing today.

## Decision — the filter does not change

**Confirmed: no change to `Securitization Action == "deferred"` (`checks.py:1068`). No production code modified.**

The filter changes the lookup outcome on three loans across historical cycles. Detailed review shows these outcomes are explained by memo data quality, administrative rows, or existing business rules rather than a defect in the filter.

1. Each of those three is accounted for: a **memo data defect** (row 45, mislabelled at source), a **type-unspecified deferment on a since-charged-off loan** (row 48), or an **already-confirmed business rule** (row 162, the resume-only CIT exception). None is a coding error.
2. Widening it is measurably harmful — 11 fabricated warnings across 2 loans, including the degradation of a correct answer — and would assert deferrals the memo does not claim (row 94).
3. The one thing genuinely missing was **visibility**, and that is already shipped: MEMO-01 names all six conflicts in the JSON response and the workbook without any rule change. The brief's own proposed outcome — *expose the conflict for a human, do not assert the deferral* — is satisfied.

## Recommendation (for decision, not implemented)

1. **Do not widen the `Securitization Action == "deferred"` filter.** One genuine case in 198 rows, with no current-cycle effect, against two measured cases where widening would inject false warnings. The evidence points the other way.
2. **Surface, don't infer** — already done. MEMO-01 names all six conflicts in the audit output and the workbook, so row 45 reaches a human without any rule change. That was the safe move the brief proposed, and it is sufficient here.
3. **Fix row 45 at source.** The correct resolution is the servicer relabelling that memo row `Deferred`, not the audit guessing.
4. **Re-run this scope each cycle.** The "no impact" finding is specific to June 2026. A conflict whose window overlaps a future cycle would matter; MEMO-01's `label-conflict` count is the tripwire, and step 1's regression guard holds it at 6.

## Collateral finding — evidence for step 3, not acted on here

`classify_deferral_type` returns `full` for `"Change Memo #3"`, `"Change Memo #1"`, `"One-Time 90-Day Deferment"` and `"CHANGE IN TERMS AGREEMENT"` — text with no payment-type signal whatsoever. Four of the six conflicts classify `full` purely by the ambiguous default, which is the direction that *enables* an auto `Loan Status → DEFERRED` write. The same default also inflates the main population: **58 of 68 `Deferred` rows classify `full`**, and the brief's own sampling suggests a substantial share of those are unspecified rather than genuinely full P&I.

Nothing here is at risk today, because these rows are excluded by the label filter and the correction path is review-gated. It is direct evidence for the step-3 proposal to treat unclassifiable text as `unclassified` + warning rather than defaulting to `full`.
