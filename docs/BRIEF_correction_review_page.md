# Brief — correction review page (frontend)

Companion to `BRIEF_correction_workflow_restructure.md`. **Build that first.** Until
`/api/audit` returns recommendations instead of a generated workbook, this page has
nothing to render.

## Stack

React 18.2 + Vite 5, plain `.jsx`, axios for HTTP, `xlsx` available client-side.
Entry is `src/index.jsx` mounting to `#root`.

No router, no state library, no UI framework, no CSS framework, no frontend test
runner. The prototype was written to match: plain React state, self-contained CSS in
a `<style>` block, no dependencies beyond React.

`correction_review.jsx` therefore ports as code, not as a design reference.

---

## Resolve first — reviewer identity

`package.json` contains no authentication library. The review page records who
approved a change to a servicer's workbook, and an approval that cannot be attributed
is not an audit record — the same standard applied to BR-002, which stays
`NOT APPROVED` until it carries a named approver and a reference.

Three possibilities, and which one is true changes the work:

1. **Auth is upstream** — SSO proxy, VPN, or a gateway injecting a header. Then the
   page reads the identity from a `/api/me` endpoint or a header the backend already
   sees, and the free-text field goes away.
2. **Auth exists in the backend** but the frontend never surfaces it. Same fix,
   different source.
3. **There is no identity.** Then the free-text field is the only option available
   today, and it must be labelled as unverified in the Correction Log — something
   like `Reviewed By (self-declared)` — so nobody later mistakes a typed name for an
   authenticated one.

Do not ship option 3 silently. If it is the reality, say so in the log column.

---

## Port changes to `correction_review.jsx`

The prototype is otherwise usable as-is. Six changes:

| Prototype | Real page |
|---|---|
| Hardcoded `CORRECTIONS` array | `GET` recommendations from the audit response |
| `window.storage` draft persistence | React state; optionally `sessionStorage` for drafts |
| Free-text `Reviewed by` | Authenticated identity — see above |
| Inert Record button | `axios.post('/api/corrections/review', payload)` |
| No session handling | `session_id` from the audit response; expiry countdown |
| No download | Receive the generated workbook and offer it |

Everything else — the delinquency ladder, mono-for-reported / sans-for-computed
typography, four decision states, domain-scoped approved-value dropdown, per
correction notes, the payload panel — carries over unchanged.

### Data shape

The prototype's correction object is the contract it needs from the backend:

```js
{
  id, gp, field, rows, postings, sheet,
  reported, prior, calculated,
  paymentReceived, servicingFee, scheduledPayment,   // BR-002 numerator
  balance, reviewFlag, rule
}
```

`servicingFee` must be present. Payment units are
`(paymentReceived + servicingFee) / scheduledPayment` per BR-002 v1.0; showing
`paymentReceived / scheduledPayment` would put a figure on screen that contradicts
the application's own calculation.

---

## Session expiry is a UI concern

`SESSION_TTL_SECONDS = 3600`, and the session gates generation. A reviewer working a
long queue can exceed it, and the failure would land at submit after all the work.

The page should show remaining time and warn before expiry — not discover it on
`POST`. If the session has expired, say so plainly and state that the files must be
re-uploaded and the audit re-run; do not imply the decisions can be recovered.

---

## Routing

There is no router. Two options:

1. **Conditional rendering** in the existing app — the review page appears as a
   section or step after audit results. No new dependency. Fits a single-session
   workflow where review directly follows audit.
2. **Add `react-router-dom`** and give review its own URL. Only worth it if a
   reviewer needs to return to a review later, or share a link.

Option 1 matches the session model: the review is bound to an in-memory session that
expires in an hour, so a shareable URL would mostly resolve to an expired session.
Prefer option 1 unless review is meant to outlive the session.

---

## What to verify when it runs

- Selecting a flagged correction shows the reported value **off** the ladder, with
  both faults named — wrong field and misspelled.
- Substituting an approved value different from the calculated one shows a second
  ladder marker and records all three values in the payload.
- The approved-value dropdown for `Delinquency Status` contains only ladder values.
  A reviewer cannot reintroduce a cross-domain value.
- Nothing is described as applied before the reviewer acts. States are
  `Recommended` → `Approved and applied` / `Rejected` / `Returned to servicer`.
- Payment evidence shows Payment Received, Servicing Fee, Adjusted Payment Total and
  Payment Units, with the BR-002 formula visible.
- Recording is blocked without a reviewer identity.

No frontend test runner exists. If one is added, these are the cases worth covering
first; if not, they are a manual checklist.

---

## Accessibility floor

The prototype meets it and the port should not regress: visible keyboard focus on
every control, `aria-current` on the selected queue item, labelled form controls,
reduced-motion respected, and a layout that collapses to one column under 940px.
