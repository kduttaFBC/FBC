# Brief — restructure the correction workflow to review-then-generate

## Objective

The audit produces **recommendations only**. No corrected workbook is generated
during `/api/audit`. Generation happens after a reviewer submits decisions, starts
from the original uploaded bytes, and writes only approved values.

Prefer this over generating during the audit and rewriting later.

---

## Current flow

`main.py` `/api/audit`, inside the per-servicer loop:

| Line | What happens |
|---|---|
| ~283 | `file_bytes = await uploaded.read()` |
| ~304 | preflight (FORM-01/02/03) — blocks on FAIL |
| ~318 | `read_servicer_file` → DataFrames |
| ~358 | `generate_corrected_workbook(...)` — **writes the workbook here** |
| ~362 | `del file_bytes; gc.collect()` |
| ~365 | corrected workbook appended, base64, to `corrected_workbooks` |
| ~552 | returned in the response |

`corrections.py` decides eligibility (`build_eligible_corrections`), writes cells
(`_apply_one_correction`), and adds the log sheet (`_add_correction_log_sheet`) in
one pass inside `generate_corrected_workbook`.

---

## Target flow

```
/api/audit
    checks.py produces STAT-01 recommendations
    corrections.build_eligible_corrections produces the candidate set
    NO workbook is written
    response carries recommendations + session_id

review page
    reviewer sets, per correction:
        approved value (from that field's domain only)
        decision: approve / reject / return
        note, notify-servicer flag

POST /api/corrections/review
    validates each decision
    generates the corrected workbook from the ORIGINAL bytes
    applies approved values only
    writes the final Correction Log
    returns the workbook
```

---

## Decide this before writing code

**Where do the original bytes live between the two calls?**

`main.py` currently frees them at line ~362 with an explicit `gc.collect()`, so
memory is a live concern. But there is already a precedent for cross-request state:

```python
SESSION_CACHE[session_id] = {"soup_df": soup_df, "created": time.time()}
```

with `_cleanup_old_sessions()` and the `session_id` returned to the client. The
chatbot endpoint already reads from it.

Three options, in order of preference:

1. **Extend `SESSION_CACHE`** to hold `{filename: file_bytes}` per servicer for the
   session. Reuses the existing lifecycle and cleanup. Cost: raw workbook bytes are
   ~2MB per servicer and are held for the session TTL. Confirm that against whatever
   drove the current `del`.
2. **Client re-submits the original file** with the review decisions. No server
   memory cost, no TTL, and the bytes are guaranteed to be the ones the reviewer saw.
   Cost: a second upload, and the server must verify it is the same file — hash it
   during the audit and compare on submission.
3. Persist to disk or object storage keyed by session. Most robust, most new
   infrastructure.

Option 2 is the safest for correctness because it removes any window where the
cached bytes and the reviewed recommendations could drift apart. Option 1 is the
least work. This is a judgement call for the repo owner, not for the implementer.

---

## Correction Log — final columns

Existing (do not reorder or rename):

```
Institution · Source Worksheet · Excel Row Numbers · Rows Updated ·
SBA GP Number · Field Updated · Original Value · Corrected Value ·
Audit Check ID · Eligibility Rule · Timestamp · Skipped Row References ·
Review Status · Review Reason
```

Add:

```
Calculated Value · Reviewer Approved Value · Decision ·
Reviewer Note · Reviewed By · Reviewed On
```

`Original Value` stays as the reported value read from the cells.
`Calculated Value` is what STAT-01 computed.
`Reviewer Approved Value` is what the reviewer authorised — equal to
`Calculated Value` unless they substituted.

Every rejected and returned item is logged with `Rows Updated = 0`.

---

## Must not change

- **`checks.py` calculation logic.** BR-002 v1.0, the reversal branch, the
  delinquency ladder, `consolidate_cycle_postings`. This is a workflow restructure.
- **The eligibility allow-list** in `build_eligible_corrections`. Same rules decide
  what is a candidate; only the timing of writing changes.
- **The cross-domain review flag.** A cross-domain reported value must still be
  flagged `Human review required` and must still never be auto-applied.
- **The approved-value domain restriction.** `Delinquency Status` substitutions come
  only from the delinquency ladder; `Loan Status` from its own list. A substitution
  must not be able to reintroduce a cross-domain value.
- **The original uploaded workbook.** Never mutated. Generation always works on a
  fresh copy of the original bytes.
- **The preflight.** FORM-03 stays a blocking control on ingestion.

---

## Response shape

`/api/audit` loses `corrected_workbooks` and gains recommendations. Keep
`corrections_skipped` — items that were never eligible, or that had no valid rows,
still belong there and are not reviewable.

`POST /api/corrections/review` returns the generated workbook plus a summary of what
was applied, rejected and returned.

Both shapes are consumed by `frontend/`, so the change is not backend-only.

---

## Tests

The suite is 160 tests; 159 passed before this work and the one failure was fixed
separately (`corrections_skipped` propagation in `main.py`).

Expect to touch:

- `test_corrections.py` (21) — `generate_corrected_workbook`'s signature and timing
  change. `CorrectionCountIsLoanAndFieldLevelTests` and `ZeroValidRowsNotAppliedTests`
  are the most exposed.
- `test_main_corrections_visibility.py` — asserts against the `/api/audit` response
  shape, which changes.
- `test_report_filtering.py`, `test_report_portfolio_summary.py` — check whether
  either depends on a corrected workbook existing at audit time.

Do not weaken an assertion to make it pass. If a test now describes the old
workflow, rewrite it to describe the new one and say so.

New coverage needed: a reviewer substituting a value different from the calculated
one; a rejected item leaving the cell untouched; a returned item leaving the cell
untouched; all three appearing in the Correction Log with the right `Rows Updated`.

---

## Context the implementer will need

- **Worked example.** Peoples Bank July 2026 produced four recommendations:
  `1278409107` 30 DAYS→DELINQUENT, `1327409104` DELINQUENT→CURRENT,
  `3883228810` DELINQUENT→CURRENT, and `4017399101` FORCLOSURE→30 DAYS, the last
  flagged for review because `FORCLOSURE` is both from the wrong field and
  misspelled. That loan carries 584,023.61 and is currently absent from every
  portfolio bucket because of the invalid value.
- **`corrections.py` is deliberately isolated** from `checks.py` — it decides what
  may be written, never what is true. Keep that boundary. A bug in one must not be
  able to silently change the other.
- **`window.storage` in the prototype is a scratchpad.** The real record is written
  by the backend. Reviewer identity must come from the authenticated user, not a
  free-text field.
