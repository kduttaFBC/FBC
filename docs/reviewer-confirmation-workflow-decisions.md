# Reviewer Confirmation Workflow — architectural decisions

*Confirmed 2026-08-08, before implementation. This records decisions taken at the close of the deferral-evidence phase so the next phase starts from them rather than re-deriving them. It is a decision record, not a design: the design is the next phase's first deliverable.*

## Why this workflow exists

`file_readers.classify_deferral_type` returns `unclassified` when the Servicer Memo's free-text Servicer Action never states whether a deferral is full P&I or principal-only. Measured on the delivered memo, 9 of 68 `Deferred` rows are in that state (`"three months deferment"`, `"90-Day Deferment"`, `"3 month payment deferral"`, `"30-day deferment"`).

Those loans keep the established full-P&I delinquency treatment as a **provisional lenient assumption** — see the governing decision in `deferralevidencebrief.md`, "Interim treatment for unclassified deferrals (2026-08-07)". Full P&I holds the delinquency rank because no payment was due; partial runs normal missed-payment math because interest was still due. With the type unknown, retaining the freeze leans lenient exactly as recomputing would lean strict. The figure is therefore provisional, and only a human can resolve it.

Today there is nowhere for that human to answer. This workflow is that place.

## The structural problem

**Confirming a type is an INPUT to STAT-01, not an output of it — and the pipeline has no path for that.** `/api/audit` computes everything in one pass; `/api/corrections/review` only *writes* approved values and never recomputes. There is no mechanism by which a fact supplied by a reviewer re-enters the calculation.

That is why this is a separate phase. It introduces new business behaviour rather than extending the evidence contract.

## Decisions

### 1. A pre-step, before corrections are generated

Unclassified deferrals are resolved in their own step, ahead of correction generation — not mixed into the corrections list as a second kind of item.

*Why:* corrections are then generated from fully-resolved evidence, so a recommendation never rests on a provisional figure. Mixing a "confirm this type" question into a list of value corrections also conflates two different reviewer actions: approving a computed value versus supplying a missing fact.

*Cost accepted:* an extra round-trip and an extra screen.

### 2. A full reconciliation re-run with overrides, never a partial recalculation

After the reviewer supplies confirmations, re-run the reconciliation with an overrides map rather than recalculating the affected loans in isolation.

*Why:* a targeted recalculation is a second implementation of the delinquency logic, and this codebase has repeatedly and deliberately refused to allow that — `consolidate_cycle_postings` exists because BAL-01/STAT-01/PAY-02/NEG-01 each had their own consolidation and diverged; `_calculate_expected_status` is shared by STAT-01 and NEG-01 for the same reason. A second path would diverge the same way, and the divergence would be invisible because both paths would look plausible.

*Feasible today:* the session already holds what a re-run needs — the original SRF bytes (stashed for correction generation) and the SOUP DataFrame (stashed for the chatbot).

### 3. Session-only storage for v1, designed so durable persistence can replace it

A confirmed type lives in the session for the initial implementation. The interface must be shaped so a durable store can be substituted without redesign.

*Why:* a confirmed deferral type is genuinely a durable fact about a loan and a memo row — next month's audit should not ask again about the same row. But nothing in this application persists between runs today: `SESSION_CACHE` is process-local with a 1-hour TTL and there is no database. Durable confirmation is a new capability, and coupling it to this workflow would make one phase carry two unrelated risks.

*Consequence to design around:* v1 confirmations do not survive the session. That is acceptable for a single reviewer working one audit, and unacceptable as a permanent answer. The seam matters more than the storage.

### 4. A complete audit trail for every confirmation

Every reviewer confirmation is recorded, as an entry type distinct from a cell write.

*Why:* a confirmation changes a reported figure, so it must be as traceable as an approved correction. It is not a cell write, and the Correction Log's schema is built around cell writes — so this needs its own entry shape rather than being forced into that one.

## What the next phase must not do

- **No back door into the classifier.** An override must reach the calculation without becoming a way to write arbitrary values into `classify_deferral_type`. The override is a reviewer-confirmed fact about one loan and one memo row, applied at a defined seam — not a mutation of the classification rules. Designing that seam is the first real design question.
- **No silent re-classification.** A confirmed type must remain visibly a human decision in the output, never presented as though the memo had stated it.
- **No change to the interim treatment for loans that are never confirmed.** Unconfirmed unclassified deferrals keep the provisional lenient assumption and the `DEFERRAL DATA QUALITY WARNING`.

## Prerequisite the design has to address first

Unclassified loans are excluded from the review path entirely today. `corrections.py`'s eligibility loop starts with:

```python
if reasons & WARNING_REASONS:
    continue  # any data-quality warning excludes this loan entirely
```

`DEFERRAL DATA QUALITY WARNING` is in `WARNING_REASONS`, so these loans never become candidates and never reach the review page. **Making them reachable is a prerequisite under any design**, and it must be done without weakening that exclusion — the exclusion is what stops a provisional figure being auto-written.

The pattern to extend, not reinvent: `corrections.py`'s existing `review_required` / `reviewFlag` path (built for the cross-domain Delinquency Status case) already surfaces a defect *as a reviewable item* rather than dropping it through an eligibility-layer `continue`.

## Deliberately left open

- **Terminal-status loans read `evidence_status = "no-deferral"` when the memo was never consulted.** The terminal short-circuit skips `_lookup_deferral` entirely, so the truth is "not evaluated". 4 loans in the June-2026 run are affected (GP 4338609104, 5886179101, 5991379104, 4778059109 — all `LIQUIDATION`, all with memo rows). The honest fix is a seventh vocabulary value, `not-evaluated`. Held open rather than changed, because it alters the published `DEFERRAL_EVIDENCE_STATUSES` contract and deserves discussion on its own terms.
- **`date-anomaly` and `deferral-claimed-dates-missing` have no real-world occurrences** in the current data (evidence across all 8 servicers: 120 `no-deferral`, 8 `documented`, 1 `type-unclassified`). Both are covered by synthetic tests only. Their absence from a report is not evidence of clean data.
