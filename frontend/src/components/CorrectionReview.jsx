import React, { useMemo, useState } from 'react'
import axios from 'axios'

// Approved-value domains, kept in sync with the backend
// (corrections._PERMITTED_DQ_STATUS / _PERMITTED_LOAN_STATUS). The dropdown a
// reviewer picks a substitute from is scoped to the target field's own domain
// so a substitution can never reintroduce a cross-domain value — the backend
// re-validates regardless (validate_approved_value), this is just the UI floor.
const DQ_LADDER = [
  'CURRENT', 'DELINQUENT', '30 DAYS', '60 DAYS', '90 DAYS',
  '120 DAYS', '150 DAYS', '150+ DAYS', '180 DAYS', '210 DAYS',
]
// The client's authoritative Loan Status domain, from the SRF template's own
// "Reference Tables" sheet — identical across every servicer workbook. A
// reviewer picks from the real field domain, exactly as they do for
// Delinquency Status above; the application itself never GENERATES a terminal
// collection status, and nothing is written without explicit approval.
//
// CLIENT-CONFIRMED (2026-08-13): these nine values are the ONLY way a Loan Status
// is entered. There is no free-text path anywhere on this screen — a Loan Status
// contradiction is resolved by the reviewer SELECTING from this domain, and
// validate_approved_value re-checks the selection against _PERMITTED_LOAN_STATUS
// server-side regardless of what this list says.
const LOAN_STATUS_DOMAIN = [
  'ACTIVE', 'DEFERRED', 'PAID-IN-FULL', 'REO', 'FORECLOSURE',
  'BANKRUPTCY', 'SUBSTITUTION-IN', 'SUBSTITUTION-OUT', 'LIQUIDATION',
]

const DECISIONS = [
  { v: 'approve', label: 'Approve — write to workbook' },
  { v: 'reject', label: 'Reject — keep the reported value' },
  { v: 'return', label: 'Return to servicer — source defect' },
]

// HUMAN REVIEW CLASSIFICATION (CLIENT-CONFIRMED 2026-08-13). Mirrors
// corrections.STANDARD_CATEGORY / REVIEW_CATEGORY_* / REVIEW_REQUIRED_CATEGORIES
// verbatim; a parity test asserts the two lists agree, the same way the existing
// DQ_LADDER and SERVICER_SOUP_KEYS parity tests do.
//
// GROUPING AND DISPLAY ONLY. `standard` and `category` are read off the server
// response and are NEVER recomputed here — this screen must not re-derive a
// classification any more than it re-derives a delinquency figure.
const STANDARD_CATEGORY = 'Standard — Delinquency'
const REVIEW_CATEGORY_LOAN_STATUS = 'Loan Status'
const REVIEW_CATEGORY_REVERSAL = 'Reversal / Transaction'
const REVIEW_CATEGORY_DATA_DOMAIN = 'Data Domain / Exception'
const REVIEW_CATEGORY_TERMINAL = 'Terminal Status Conflict'
const REVIEW_CATEGORY_OTHER = 'Other — Flagged'

const REVIEW_REQUIRED_CATEGORIES = [
  REVIEW_CATEGORY_LOAN_STATUS,
  REVIEW_CATEGORY_REVERSAL,
  REVIEW_CATEGORY_DATA_DOMAIN,
  REVIEW_CATEGORY_TERMINAL,
  REVIEW_CATEGORY_OTHER,
]

// CLIENT-CONFIRMED (2026-08-14): the prefix of
// corrections.REVIEW_REASON_LENDER_REPORTED_LIQUIDATION. Matched rather than
// duplicated in full, so the backend stays the single source of the reason text.
// Every in-scope lender-reported LIQUIDATION carries this ONE reason; the
// contradictions against prior status, balances, dates, payments, bankruptcy /
// foreclosure / REO information, proceeds, losses, recoveries, reversals and memo
// evidence are shown as EVIDENCE on the loan card, never as further decision rules.
const LENDER_LIQUIDATION_REASON = 'LENDER REPORTED LIQUIDATION'

const CATEGORY_BLURB = {
  [REVIEW_CATEGORY_LOAN_STATUS]:
    'Every Loan Status write requires explicit reviewer approval — including one the audit derived without flagging it. Select the correct status from the field’s own nine-value domain.',
  [REVIEW_CATEGORY_REVERSAL]:
    'Transaction history materially affects the recommendation: a negative Payment Received posting this cycle, whether the consolidated net is negative, exactly $0.00, or positive.',
  [REVIEW_CATEGORY_DATA_DOMAIN]:
    'The value the servicer reported is outside its permitted domain — a misspelling, or a value belonging to another field. A source-data defect worth notifying.',
  [REVIEW_CATEGORY_TERMINAL]:
    'The servicer reports a terminal collection status (FORECLOSURE / LIQUIDATION / REO) for this loan. That report is evidence to weigh, not a confirmed fact.',
  [REVIEW_CATEGORY_OTHER]:
    'Flagged for review for a reason this screen has no dedicated section for. Never eligible for bulk approval.',
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

function downloadBase64File(base64, filename) {
  const byteCharacters = atob(base64)
  const byteArray = new Uint8Array(Array.from(byteCharacters, (c) => c.charCodeAt(0)))
  const url = URL.createObjectURL(new Blob([byteArray], { type: XLSX_MIME }))
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

const money = (n) =>
  n === null || n === undefined || n === '' || Number.isNaN(Number(n))
    ? '—'
    : Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })

function dash(value) {
  return value === null || value === undefined || value === '' ? '—' : String(value)
}

function domainFor(field) {
  return field === 'Loan Status' ? LOAN_STATUS_DOMAIN : DQ_LADDER
}

// CLIENT-CONFIRMED (2026-08-15): the required balance decision that accompanies a
// reviewer-approved LIQUIDATION -> ACTIVE reversal.
//
// TWO CHOICES, BOTH SERVER-SUPPLIED, NO FREE-FORM ENTRY. The values come from
// rec.balanceChoices, which the backend derives from the loan's own Beginning
// Balance; this component never computes an amount. The backend re-validates
// every submission against the same two values (validate_approved_value), so
// this is the UI floor, not the control.
//
// NEITHER OPTION IS PRE-SELECTED, and that is deliberate even when the evidence
// agrees: restoring a balance is a financial decision and must be made, not
// accepted by default.
//
// PRIOR MONTH ENDING BALANCE IS EVIDENCE ONLY. It is displayed, and when it
// differs from the Beginning Balance the reviewer is warned -- but it is never
// offered as a third option, because the client's rule restores to the current
// month's Beginning Balance or to nothing.
function BalanceDecision({ rec, decision, onChange }) {
  const choices = rec.balanceChoices || []
  // Compared as strings: the radio's value attribute is a string, and matching
  // on Number() would make a chosen 0.00 indistinguishable from "no choice yet".
  const chosen =
    decision.approved_value === null || decision.approved_value === undefined
      ? ''
      : String(decision.approved_value)
  const restore = choices.find((c) => c.action === 'restore')

  return (
    <div className="balance-decision">
      <div className="balance-evidence">
        <div>
          <strong>Beginning Balance (this month)</strong> {money(rec.beginningBalance)}
        </div>
        {/* Rendered only when SOUP actually carried a prior figure -- an absent
            prior period is shown as absent, never as 0.00. */}
        {rec.priorEndingBalance !== null && rec.priorEndingBalance !== undefined && (
          <div>
            <strong>Prior Month Ending Balance</strong> {money(rec.priorEndingBalance)}{' '}
            <span className="hint">(supporting evidence only)</span>
          </div>
        )}
        <div>
          <strong>Current Ending Balance (lender reported)</strong> {money(rec.reported)}
        </div>
        {rec.liquidationDate && (
          <div>
            <strong>Liquidation Date</strong> {dash(rec.liquidationDate)}
          </div>
        )}
        {rec.liquidationLoss !== null && rec.liquidationLoss !== undefined && (
          <div>
            <strong>Liquidation Loss</strong> {money(rec.liquidationLoss)}
          </div>
        )}
      </div>

      {rec.balanceEvidenceConflict && (
        <div className="balance-warning" role="alert">
          ⚠ The Prior Month Ending Balance and this month's Beginning Balance{' '}
          <strong>differ</strong>. Both are shown above. Neither option is
          pre-selected — choose deliberately.
        </div>
      )}

      <fieldset className="balance-choices">
        <legend>Required Balance Decision</legend>
        {choices.map((choice) => (
          <label key={choice.action} className="balance-choice">
            <input
              type="radio"
              name={`balance-${rec.id}`}
              value={String(choice.value)}
              checked={chosen === String(choice.value)}
              onChange={() =>
                onChange(rec.id, { decision: 'approve', approved_value: choice.value })
              }
            />
            {choice.action === 'restore'
              ? `Restore Ending Balance to ${money(choice.value)}`
              : `Keep Ending Balance at ${money(choice.value)}`}
          </label>
        ))}
      </fieldset>

      {/* The consequences, shown BEFORE the decision is submitted. Silent
          financial change is exactly what this rule exists to prevent. */}
      {restore && chosen === String(restore.value) && (
        <div className="balance-consequences">
          <strong>This correction will also apply:</strong>
          <div>
            Ending Balance {money(rec.reported)} → {money(restore.value)}
          </div>
          <div>Liquidation Date {dash(rec.liquidationDate)} → (blank)</div>
          <div>Liquidation Loss {money(rec.liquidationLoss)} → (blank)</div>
        </div>
      )}
      {chosen === '0' && (
        <div className="balance-consequences">
          <strong>Liquidation Date and Liquidation Loss will be kept.</strong>{' '}
          <span className="hint">
            The loan will report ACTIVE at {money(0)} with its liquidation fields
            intact; the contradiction stays visible for review.
          </span>
        </div>
      )}
    </div>
  )
}

// BR-002 v1.0 payment units = (Payment Received + Servicing Fee) / Scheduled
// Payment. Servicing Fee is part of the numerator — showing Payment Received
// alone would contradict the application's own calculation.
function paymentUnits(rec) {
  // A REVERSAL row is not judged by this ratio. The backend's reversal branch
  // (checks._apply_payment_transition) measures a deficiency of
  // scheduled + |negative amount| in 80%-of-scheduled units and DELIBERATELY
  // excludes the servicing fee — checks.py records why: a small negative
  // payment plus a fee could otherwise sum to a positive figure and bypass the
  // reversal path entirely. Rendering (received + fee) / scheduled here would
  // show the reviewer a number that played no part in the calculated value
  // (e.g. loan 5951029110: −$10.92 + $10.60 fee / $313.16 renders "-0.00
  // units" beside a DELINQUENT result). Reversal rows show the real deficiency
  // arithmetic instead — see reversalEvidence below.
  if (rec.reversalEquivalents != null) return null
  const sched = Number(rec.scheduledPayment)
  if (!sched) return null
  const adjusted = Number(rec.paymentReceived || 0) + Number(rec.servicingFee || 0)
  return adjusted / sched
}

// The reversal arithmetic that DID drive the calculated value, transcribed from
// checks._apply_payment_transition's client-confirmed (2026-08-04, second
// revision) formula. Shown in place of the BR-002 ratio for reversal rows so
// the reviewer sees the figures the result actually came from.
function reversalEvidence(rec) {
  if (rec.reversalEquivalents == null) return null
  const sched = Number(rec.scheduledPayment)
  const received = Number(rec.paymentReceived || 0)
  const n = Number(rec.reversalEquivalents)
  return {
    deficiency: sched + Math.abs(received),
    thresholdUnit: sched * 0.8,
    equivalents: n,
  }
}

// One line describing this cycle's reversal state, for the shared loan card.
// Reads only flags the audit already set; computes nothing.
function reversalSummary(rec) {
  if (rec.reversalEquivalents != null) {
    return `net-negative — ${rec.reversalEquivalents} payment equivalent(s) removed`
  }
  if (rec.offsettingZeroNetReversal) {
    return 'reversed postings offset to a consolidated net of exactly $0.00'
  }
  if (rec.hasNegativePosting) {
    return 'reversed posting(s) present within a positive consolidated net'
  }
  return 'none'
}

// The deferral-evidence contract, computed ONCE in checks.py (STAT-01) and
// only rendered here. This screen must never re-derive what the evidence says
// from the dates or the memo action — that is precisely the drift the shared
// contract exists to prevent. These maps translate the value for a reader;
// they do not decide it.
const EVIDENCE_MEANING = {
  'documented': 'deferral confirmed, dates consistent, type known',
  'no-deferral': 'no deferral applies to this cycle',
  'deferral-claimed-dates-missing': 'a deferral is asserted but no usable start date was written',
  'lookup-failed': 'memo rows exist for this loan but none matched a known Lender for this servicer — NOT evidence that there is no deferral',
  'date-anomaly': 'the matched memo row has impossible dates (end or resume before start)',
  'type-unclassified': 'the memo does not say whether this is full P&I or partial — the delinquency figure is provisional pending confirmation',
}

const EVIDENCE_COLORS = {
  'lookup-failed': 'var(--amber, #BA7517)',
  'date-anomaly': 'var(--amber, #BA7517)',
  'deferral-claimed-dates-missing': 'var(--amber, #BA7517)',
  'type-unclassified': 'var(--amber, #BA7517)',
}

// Shown whenever there is deferral evidence worth reading. "no-deferral" is a
// fact rather than a finding, so it stays out of the way; every other status
// is surfaced, including "documented", since the window and the memo label are
// exactly what a reviewer needs to judge a deferral-driven correction.
function hasDeferralEvidence(rec) {
  return Boolean(rec.evidenceStatus) && rec.evidenceStatus !== 'no-deferral'
}

// CLIENT-CONFIRMED (2026-08-12): the current-period evidence behind a BANKRUPTCY
// Loan Status recommendation. Rendered only when there is something to show, so
// no existing recommendation's evidence cell changes.
//
// The client asked specifically that a Liquidation Date sitting against a
// POSITIVE Ending Balance be surfaced PROMINENTLY as contradictory: it does not
// suppress the recommendation (the zero-balance exception is keyed on a $0
// balance), but the reviewer must not have to infer the conflict from two
// separate cells. That pairing is what `contradiction` marks.
function currentPeriodEvidence(rec) {
  const bankruptcy = rec.bankruptcyDate
  const liquidation = rec.liquidationDate
  const paidInFull = rec.paidInFullDate
  if (!bankruptcy && !liquidation && !paidInFull) return null
  const balance = Number(rec.balance)
  return {
    bankruptcy,
    liquidation,
    paidInFull,
    // Both halves required: a liquidation date alone is not a contradiction, and
    // a positive balance alone certainly isn't.
    contradiction: Boolean(liquidation) && Number.isFinite(balance) && balance > 0,
  }
}

// A recommendation's loan identity. Two recommendations share a loan card when
// this matches — never `gp` alone, since two servicers can carry the same SBA GP
// Number and their recommendations are separate work.
function loanKeyOf(rec) {
  return `${rec.servicer}||${rec.gp}`
}

// CLIENT-CONFIRMED (2026-08-13): grouped loan review. A loan with more than one
// recommendation renders as ONE card with shared evidence and a separate decision
// control per recommendation. The recommendations keep their own ids and their own
// decisions throughout — grouping is a display arrangement and nothing else, and
// the submitted payload is byte-identical to what the flat list produced.
function groupByLoan(recs) {
  const order = []
  const byKey = new Map()
  recs.forEach((r) => {
    const key = loanKeyOf(r)
    if (!byKey.has(key)) {
      byKey.set(key, { key, servicer: r.servicer, gp: r.gp, recs: [] })
      order.push(key)
    }
    byKey.get(key).recs.push(r)
  })
  return order.map((k) => byKey.get(k))
}

// Which section/category a loan card sorts under: the highest-precedence Review
// Required category among its recommendations, matching the backend's own
// precedence order. A card holding ANY Review Required recommendation belongs to
// REVIEW REQUIRED — which is what puts a contradicted loan's Standard delinquency
// companion on screen beside the Loan Status decision it must be weighed against.
function cardCategory(group) {
  for (const cat of REVIEW_REQUIRED_CATEGORIES) {
    if (group.recs.some((r) => !r.standard && r.category === cat)) return cat
  }
  return STANDARD_CATEGORY
}

// ---------------------------------------------------------------------------
// Presentational components. Module scope, props only — deliberately not nested
// inside CorrectionReview, so each piece of markup is defined exactly once.
// ---------------------------------------------------------------------------

function CategoryBadge({ rec }) {
  return (
    <span
      style={{
        display: 'inline-block',
        padding: '1px 6px',
        borderRadius: 10,
        border: '1px solid var(--border, #ccc)',
        fontSize: '0.85em',
        color: rec.standard ? 'inherit' : 'var(--amber, #BA7517)',
      }}
    >
      {rec.standard ? 'Standard' : rec.category}
    </span>
  )
}

// The shared evidence block for a multi-recommendation loan card.
//
// CLIENT-CONFIRMED (2026-08-13): this is the read-only evidence a reviewer needs
// to judge a Loan Status contradiction, gathered in one place for the whole loan.
// Every figure is passed through from the audit — this block computes NOTHING and
// re-derives NOTHING. A field absent from the source reads "—" rather than being
// inferred from anything else.
function LoanEvidenceHeader({ group }) {
  const loanStatusRec = group.recs.find((r) => r.field === 'Loan Status')
  const dqRec = group.recs.find((r) => r.field === 'Delinquency Status')
  const any = group.recs[0]
  const cp = currentPeriodEvidence(any)
  const deferral = group.recs.find((r) => hasDeferralEvidence(r))
  return (
    <tr>
      <td colSpan={10} style={{ background: 'var(--surface-2, rgba(0,0,0,0.03))' }}>
        <div className="field-name">
          Loan {group.gp} · {group.servicer} · {group.recs.length} recommendations
        </div>
        <div className="hint" style={{ marginTop: 4 }}>
          <div>
            <strong>Loan Status</strong> — reported {dash(any.servicerLoanStatus)} · prior
            finalised (SOUP) {dash(any.priorLoanStatus)}
            {loanStatusRec && (
              <> · recommended <strong>{dash(loanStatusRec.calculated)}</strong></>
            )}
          </div>
          <div>
            <strong>Balances</strong> — prior finalised Ending Balance{' '}
            {money(any.priorEndingBalance)} · current Ending Balance {money(any.balance)}
          </div>
          <div>
            <strong>Delinquency</strong> — reported {dash(dqRec ? dqRec.reported : null)} ·
            calculated {dash(dqRec ? dqRec.calculated : null)}
          </div>
          <div>
            <strong>Payment</strong> — Received {money(any.paymentReceived)} · Fee{' '}
            {money(any.servicingFee)} · Scheduled {money(any.scheduledPayment)} ·{' '}
            {any.postings || 1} posting{(any.postings || 1) === 1 ? '' : 's'}
          </div>
          <div>
            <strong>Terminal events</strong> — Liquidation Date {dash(any.liquidationDate)} ·
            Bankruptcy Date {dash(any.bankruptcyDate)} · Paid-In-Full Date{' '}
            {dash(any.paidInFullDate)} · Foreclosure Date {dash(any.foreclosureDate)} · REO
            Date {dash(any.reoDate)}
          </div>
          <div>
            <strong>Liquidation amounts</strong> — Proceeds {money(any.liquidationProceeds)} ·
            Loss {money(any.liquidationLoss)} · Subsequent Recoveries/(Losses){' '}
            {money(any.subsequentRecoveries)}
          </div>
          <div>
            <strong>Reversal</strong> — {reversalSummary(any)}
          </div>
          {deferral && (
            <div>
              <strong>Servicer Memo</strong>{' '}
              <span style={{ color: EVIDENCE_COLORS[deferral.evidenceStatus] || 'inherit' }}>
                {deferral.evidenceStatus}
              </span>
              {EVIDENCE_MEANING[deferral.evidenceStatus] && (
                <> — {EVIDENCE_MEANING[deferral.evidenceStatus]}</>
              )}
              <div>
                Memo action {dash(deferral.memoSecuritizationAction)} · Window{' '}
                {dash(deferral.deferralStart)} → {dash(deferral.deferralEnd)} · Resume{' '}
                {dash(deferral.deferralResume)}
              </div>
            </div>
          )}
          {cp && cp.contradiction && (
            <div style={{ color: 'var(--amber, #BA7517)', fontWeight: 600, marginTop: 4 }}>
              ⚠ CONTRADICTORY: a Liquidation Date of {cp.liquidation} is reported against an
              Ending Balance of {money(any.balance)} (greater than $0), with a reported Loan
              Status of {dash(any.servicerLoanStatus)}. The zero-balance exception does not
              apply — weigh both before approving.
            </div>
          )}
          <div style={{ marginTop: 4 }}>
            Each recommendation below is decided independently and carries its own record.
          </div>
        </div>
      </td>
    </tr>
  )
}

// One recommendation's decision row, plus the per-recommendation evidence rows.
// `grouped` suppresses only the evidence the shared card header already carries,
// so nothing is ever shown twice. The review flag stays per recommendation: it
// describes THAT recommendation, not the loan.
function RecommendationRow({ rec, grouped, decision, onChange }) {
  const d = decision || {}
  const units = paymentUnits(rec)
  const reversal = reversalEvidence(rec)
  const currentPeriod = currentPeriodEvidence(rec)
  return (
    <React.Fragment>
      {/* id is the jump-to-next-undecided anchor. */}
      <tr id={`rec-${rec.id}`}>
        <td>
          <strong>{rec.gp}</strong>
          <div className="hint">
            {rec.servicer}
            {rec.postings ? ` · ${rec.postings} posting${rec.postings === 1 ? '' : 's'}` : ''}
          </div>
        </td>
        <td>
          {rec.field}
          <div className="hint">
            <CategoryBadge rec={rec} />
          </div>
        </td>
        <td style={{ fontFamily: 'monospace' }}>{String(rec.reported)}</td>
        <td style={{ fontFamily: 'monospace' }}>{rec.prior == null ? '—' : String(rec.prior)}</td>
        <td style={{ fontFamily: 'monospace' }}>{String(rec.calculated)}</td>
        <td className="hint">
          {reversal ? (
            <>
              <strong>Reversal</strong> — net Received {money(rec.paymentReceived)} / Sched{' '}
              {money(rec.scheduledPayment)}
              <div>
                Deficiency {money(reversal.deficiency)} / {money(reversal.thresholdUnit)}{' '}
                (80% of sched) ={' '}
                <strong>
                  {reversal.equivalents} payment
                  {reversal.equivalents === 1 ? '' : 's'} removed
                </strong>
              </div>
              <div>Fee {money(rec.servicingFee)} — excluded from this calculation</div>
            </>
          ) : (
            <>
              Received {money(rec.paymentReceived)} + Fee {money(rec.servicingFee)} / Sched{' '}
              {money(rec.scheduledPayment)}
              {units != null && (
                <>
                  {' '}= <strong>{units.toFixed(2)} units</strong>
                </>
              )}
              {/* An OFFSETTING reversal keeps the ratio above: its net really is
                  $0.00 and the ordinary no-payment branch really did produce the
                  calculated value, so unlike a net-negative reversal the figure on
                  screen is the arithmetic that ran. What the $0.00 hides is that
                  reversed postings are inside it — without this line a reviewer
                  cannot tell this apart from a loan that simply paid nothing. */}
              {rec.offsettingZeroNetReversal && (
                <div>
                  <strong>Includes reversed postings</strong> offset to a net of exactly $0.00
                </div>
              )}
              {/* CLIENT-CONFIRMED (2026-08-13): the third reversal case — a reversed
                  posting whose CONSOLIDATED NET IS POSITIVE. Neither flag above can
                  express it, since both are keyed on the net, and without saying so
                  the reviewer cannot tell why an otherwise unflagged recommendation
                  sits in Reversal / Transaction. */}
              {rec.hasNegativePosting && !rec.offsettingZeroNetReversal && (
                <div>
                  <strong>Includes reversed postings</strong> within a positive net of{' '}
                  {money(rec.paymentReceived)}
                </div>
              )}
            </>
          )}
          <div>Balance {money(rec.balance)}</div>
        </td>
        <td>
          <select
            aria-label={`Decision for loan ${rec.gp} ${rec.field}`}
            value={d.decision || ''}
            onChange={(e) => onChange(rec.id, { decision: e.target.value })}
          >
            <option value="">Not yet decided</option>
            {DECISIONS.map((opt) => (
              <option key={opt.v} value={opt.v}>
                {opt.label}
              </option>
            ))}
          </select>
        </td>
        <td>
          {/* The balance decision replaces the value dropdown entirely: there is
              no free-form entry and no status domain here, only the two
              server-supplied amounts. Rendered whenever the item requires a
              balance decision, not only after "approve" is chosen, because
              choosing an amount IS the approval. */}
          {rec.requiresBalanceDecision ? (
            <BalanceDecision rec={rec} decision={d} onChange={onChange} />
          ) : d.decision === 'approve' ? (
            <select
              aria-label={`Approved value for loan ${rec.gp} ${rec.field}`}
              value={d.approved_value || rec.calculated || ''}
              onChange={(e) => onChange(rec.id, { approved_value: e.target.value })}
            >
              {/* CLIENT-CONFIRMED (2026-08-14): a mandatory lender-reported
                  LIQUIDATION review carries NO engine recommendation when the
                  audit has no proposable status. Nothing is defaulted -- not
                  ACTIVE, not LIQUIDATION -- so the reviewer must choose. This adds
                  an empty placeholder only; the domain below is unchanged and is
                  still exactly the nine permitted values. `|| ''` keeps the select
                  controlled, since value={null} would silently make it
                  uncontrolled and drop the reviewer's choice. */}
              {!rec.calculated && <option value="">— Select Loan Status —</option>}
              {domainFor(rec.field).map((v) => (
                <option key={v} value={v}>
                  {v}
                </option>
              ))}
            </select>
          ) : (
            <span className="hint">—</span>
          )}
        </td>
        <td style={{ textAlign: 'center' }}>
          <input
            type="checkbox"
            aria-label={`Notify servicer for loan ${rec.gp} ${rec.field}`}
            checked={!!d.notify}
            onChange={(e) => onChange(rec.id, { notify: e.target.checked })}
          />
        </td>
        <td>
          <input
            type="text"
            aria-label={`Reviewer note for loan ${rec.gp} ${rec.field}`}
            value={d.note || ''}
            placeholder="optional"
            onChange={(e) => onChange(rec.id, { note: e.target.value })}
            style={{ width: 140 }}
          />
        </td>
      </tr>
      {!grouped && currentPeriod && (
        <tr>
          <td colSpan={10} className="hint">
            <strong>Current-period evidence:</strong> Bankruptcy Date{' '}
            {dash(currentPeriod.bankruptcy)} · Paid-In-Full Date{' '}
            {dash(currentPeriod.paidInFull)} · Liquidation Date{' '}
            {dash(currentPeriod.liquidation)} · Ending Balance {money(rec.balance)}
            {currentPeriod.contradiction && (
              <div style={{ color: 'var(--amber, #BA7517)', fontWeight: 600, marginTop: 4 }}>
                ⚠ CONTRADICTORY: a Liquidation Date of {currentPeriod.liquidation} is reported
                against an Ending Balance of {money(rec.balance)} (greater than $0), with a
                reported Loan Status of {dash(rec.servicerLoanStatus)}. The zero-balance
                exception does not apply — weigh both before approving.
              </div>
            )}
          </td>
        </tr>
      )}
      {!grouped && hasDeferralEvidence(rec) && (
        <tr>
          <td colSpan={10} className="hint">
            <strong>Deferral evidence:</strong>{' '}
            <span style={{ color: EVIDENCE_COLORS[rec.evidenceStatus] || 'inherit' }}>
              {rec.evidenceStatus}
            </span>
            {EVIDENCE_MEANING[rec.evidenceStatus] && (
              <> — {EVIDENCE_MEANING[rec.evidenceStatus]}</>
            )}
            <div>
              Memo action {dash(rec.memoSecuritizationAction)} · Window{' '}
              {dash(rec.deferralStart)} → {dash(rec.deferralEnd)} · Resume{' '}
              {dash(rec.deferralResume)} · Prior (SOUP) Loan Status{' '}
              {dash(rec.priorLoanStatus)} · Servicer Loan Status{' '}
              {dash(rec.servicerLoanStatus)}
            </div>
          </td>
        </tr>
      )}
      {/* CLIENT-CONFIRMED (2026-08-14): the reviewer instruction for a mandatory
          lender-reported LIQUIDATION review. Stated plainly and above the reason
          text, because this item asks the reviewer to make a decision rather than
          to confirm one the audit already made. */}
      {rec.field === 'Loan Status'
        && (rec.reviewFlag || '').includes(LENDER_LIQUIDATION_REASON) && (
        <tr>
          <td colSpan={10} style={{ fontWeight: 600 }}>
            The lender reported this loan as LIQUIDATION. The prior finalized record
            and current servicing evidence are shown below. Please review the
            evidence and select the Loan Status you believe is appropriate.
            {!rec.calculated && (
              <span style={{ fontWeight: 400 }}>
                {' '}Falcon Bridge has no recommended value for this loan — you must
                choose one of the nine permitted statuses.
              </span>
            )}
          </td>
        </tr>
      )}
      {rec.reviewFlag && (
        <tr>
          <td colSpan={10} style={{ color: 'var(--amber, #BA7517)' }}>
            ⚠ {rec.reviewFlag}
          </td>
        </tr>
      )}
    </React.Fragment>
  )
}

// The one and only recommendation table in this file. Rendered once per section
// and once per Review Required category; there is no other table of
// recommendations anywhere on the screen.
function ReviewTable({ cards, decisionOf, onChange }) {
  return (
    <table className="detail-table" style={{ width: '100%' }}>
      <thead>
        <tr>
          <th>Loan</th>
          <th>Field / Category</th>
          <th>Reported</th>
          <th>Prior</th>
          <th>Calculated</th>
          <th>Payment evidence (BR-002)</th>
          <th>Decision</th>
          <th>Approved value</th>
          <th>Notify</th>
          <th>Note</th>
        </tr>
      </thead>
      <tbody>
        {cards.map((group) => (
          <React.Fragment key={group.key}>
            {group.recs.length > 1 && <LoanEvidenceHeader group={group} />}
            {group.recs.map((rec) => (
              <RecommendationRow
                key={rec.id}
                rec={rec}
                grouped={group.recs.length > 1}
                decision={decisionOf(rec.id)}
                onChange={onChange}
              />
            ))}
          </React.Fragment>
        ))}
      </tbody>
    </table>
  )
}

export default function CorrectionReview({ recommendations, sessionId }) {
  const recs = recommendations || []
  const [decisions, setDecisions] = useState({}) // id -> {decision, approved_value, note, notify}
  const [reviewer, setReviewer] = useState('')
  const [submitting, setSubmitting] = useState(false)
  const [error, setError] = useState('')
  const [result, setResult] = useState(null)
  const [collapsed, setCollapsed] = useState({}) // category -> bool

  const decisionOf = (id) => decisions[id] || {}
  const setDecision = (id, patch) =>
    setDecisions((cur) => ({ ...cur, [id]: { ...cur[id], ...patch } }))

  const decidedCount = useMemo(
    () => recs.filter((r) => decisionOf(r.id).decision).length,
    [recs, decisions]
  )
  const pendingCount = recs.length - decidedCount

  // Classification counts, read straight off the server response. `standard` is
  // never derived here.
  const standardRecs = useMemo(() => recs.filter((r) => r.standard), [recs])
  const reviewRecs = useMemo(() => recs.filter((r) => !r.standard), [recs])

  // Loan cards, partitioned into the two sections. A card goes to REVIEW REQUIRED
  // if it holds ANY Review Required recommendation.
  const { standardCards, reviewCardsByCategory } = useMemo(() => {
    const std = []
    const byCat = {}
    REVIEW_REQUIRED_CATEGORIES.forEach((c) => {
      byCat[c] = []
    })
    groupByLoan(recs).forEach((group) => {
      const cat = cardCategory(group)
      if (cat === STANDARD_CATEGORY) std.push(group)
      else byCat[cat].push(group)
    })
    return { standardCards: std, reviewCardsByCategory: byCat }
  }, [recs])

  // CLIENT-CONFIRMED (2026-08-13): THE COMPANION HOLD. A Standard recommendation
  // whose loan still has an UNDECIDED Loan Status recommendation is held out of
  // the bulk control until that Loan Status is decided.
  //
  // Why: approving a loan's delinquency figure in bulk while its Loan Status is
  // contradicted is exactly the split judgement the grouped loan card exists to
  // prevent. The item stays CLASSIFIED Standard — the 77/50 counts are unchanged
  // by this — stays listed, and stays individually approvable. Only the bulk sweep
  // declines it, and it becomes eligible the moment the Loan Status is decided
  // either way.
  //
  // Keyed on Loan Status specifically, per the client's wording, which is also the
  // fully general condition here rather than a narrower one: a Standard
  // recommendation is always a Delinquency Status recommendation, and ids are
  // (servicer, gp, field), so the only companion a loan can have is a Loan Status
  // recommendation — every one of which is Review Required.
  // Companion hold. A Standard correction is not swept while the SAME loan still
  // has an undecided Loan Status decision -- or, since 2026-08-15, an undecided
  // REQUIRED BALANCE decision. Both are judgements about the same loan, and a
  // bulk approval must not run ahead of either.
  const loanStatusUndecided = (rec) =>
    recs.some(
      (other) =>
        loanKeyOf(other) === loanKeyOf(rec) &&
        (other.field === 'Loan Status' || other.requiresBalanceDecision) &&
        !decisionOf(other.id).decision
    )

  const bulkEligible = useMemo(
    () => standardRecs.filter((r) => !decisionOf(r.id).decision && !loanStatusUndecided(r)),
    [recs, decisions]
  )
  const bulkHeld = useMemo(
    () => standardRecs.filter((r) => !decisionOf(r.id).decision && loanStatusUndecided(r)),
    [recs, decisions]
  )

  // CLIENT-CONFIRMED (2026-08-14): an "approve" on a recommendation that carries NO
  // engine value -- a mandatory lender-reported LIQUIDATION with no proposable
  // status -- is INCOMPLETE until the reviewer picks one of the nine permitted
  // statuses.
  //
  // This is a MALFORMED-ENTRY guard, not a review-completeness gate. D4 still does
  // not gate submission on pending items, and a partial review is still legitimate.
  // The server already fails closed on a blank approved value
  // (validate_approved_value -> "Approved value is blank." -> approved_not_applied
  // -> D4 blocks final reporting), so this only stops a submission that provably
  // could not be applied, and says which loans need a choice.
  const approvalsMissingValue = useMemo(
    () => recs.filter((r) => {
      const d = decisionOf(r.id)
      // A balance decision of 0.00 is a REAL answer, so the falsy check below
      // would wrongly read it as "no value chosen". Balance items are judged on
      // whether an amount was chosen at all.
      if (r.requiresBalanceDecision) {
        return (
          d.decision === 'approve' &&
          (d.approved_value === null || d.approved_value === undefined)
        )
      }
      return d.decision === 'approve' && !d.approved_value && !r.calculated
    }),
    [recs, decisions]
  )

  // CLIENT-CONFIRMED (2026-08-15): a reviewer-approved LIQUIDATION -> ACTIVE is
  // not fully resolved until its balance decision is made. The server enforces
  // this for real -- the balance item is its own candidate, so D4's pending count
  // already blocks reporting readiness -- and this is the matching UI floor, so
  // the reviewer is stopped here rather than discovering it after submitting.
  const balanceDecisionsPending = useMemo(
    () => recs.filter((r) => r.requiresBalanceDecision && !decisionOf(r.id).decision),
    [recs, decisions]
  )

  // Populates the EXISTING client-side decision state for the eligible Standard
  // recommendations and does nothing else. It does not submit, does not call any
  // endpoint, does not overwrite a decision already recorded, and never touches a
  // Review Required recommendation — bulkEligible is filtered on `standard`, which
  // the server computed. The reviewer still types their name and still presses
  // Record, and every pre-filled decision stays individually editable until they do.
  const approveAllStandard = () =>
    setDecisions((cur) => {
      const out = { ...cur }
      bulkEligible.forEach((r) => {
        if (!(out[r.id] && out[r.id].decision)) {
          out[r.id] = { ...out[r.id], decision: 'approve' }
        }
      })
      return out
    })

  const jumpToNextUndecided = () => {
    const next = recs.find((r) => !decisionOf(r.id).decision)
    if (!next) return
    const el = document.getElementById(`rec-${next.id}`)
    if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center' })
  }

  const rejectAllRemaining = () =>
    setDecisions((cur) => {
      const out = { ...cur }
      recs.forEach((r) => {
        if (!(out[r.id] && out[r.id].decision)) {
          out[r.id] = { ...out[r.id], decision: 'reject' }
        }
      })
      return out
    })

  // D4 (CLIENT-CONFIRMED 2026-08-13): submission is deliberately NOT gated on
  // pendingCount. Corrected workbooks and Correction Logs are PER-DECISION
  // artifacts and are explicitly not gated, so a reviewer must be able to record
  // partial review work and take the workbooks for it. D4 gates FINAL REPORTING,
  // not the ability to record a completed decision.
  //
  // Reviewer name is still required -- an unattributed sign-off is not an audit
  // record -- and an empty submission still records nothing, so at least one
  // decision is required. Both mirror the server's own guards; neither is the
  // authority. review_complete and reporting_ready are computed SERVER-SIDE and
  // read off the response, never derived here.
  const blocked = !reviewer.trim() || decidedCount === 0
  // Zero recommendations is 0 of 0 -- complete and reporting-ready -- so the
  // reviewer must still be able to submit and get the reports.
  const nothingToReview = recs.length === 0

  const submit = async () => {
    setSubmitting(true)
    setError('')
    try {
      const payload = {
        session_id: sessionId,
        reviewed_by: reviewer.trim(),
        reviewed_on: new Date().toISOString().replace(/\.\d{3}Z$/, 'Z'),
        decisions: recs
          .filter((r) => decisionOf(r.id).decision)
          .map((r) => {
            const d = decisionOf(r.id)
            const out = { id: r.id, decision: d.decision }
            // Only send an approved value for an approval; default to the
            // calculated value when the reviewer didn't substitute.
            if (d.decision === 'approve') out.approved_value = d.approved_value || r.calculated
            if (d.note) out.reviewer_note = d.note
            if (d.notify) out.notify_servicer = true
            return out
          }),
      }
      const apiUrl = import.meta.env.VITE_API_URL || ''
      const { data } = await axios.post(`${apiUrl}/api/corrections/review`, payload)
      setResult(data)
    } catch (err) {
      const status = err.response?.status
      if (status === 410) {
        setError(
          err.response?.data?.error ||
            'This review session has expired. Please re-upload the files and re-run the audit.'
        )
      } else {
        setError(err.response?.data?.error || 'Could not record the corrections. Please try again.')
      }
    } finally {
      setSubmitting(false)
    }
  }

  const decidedIn = (cards) =>
    cards.reduce((n, g) => n + g.recs.filter((r) => decisionOf(r.id).decision).length, 0)
  const countIn = (cards) => cards.reduce((n, g) => n + g.recs.length, 0)

  // D4 (CLIENT-CONFIRMED 2026-08-13): a clean month must still be able to record
  // its review and produce its reports. This component previously returned null
  // when there was nothing to review, so the reviewer-name field and the submit
  // button never rendered at all -- there was no path from a clean audit to a
  // Monthly Servicer Report or a Final Audited Portfolio Summary. 0 of 0 is the
  // most complete review possible, and it is reporting-ready.
  // NOTE: no early return here any more. `nothingToReview` drives the sections and
  // the bulk controls instead, so the submit affordance is always present.

  return (
    <div className="card" style={{ marginBottom: 22 }}>
      <h3 style={{ marginTop: 0 }}>Human Review</h3>
      <p style={{ color: 'var(--text-secondary)', marginTop: 0 }}>
        These are <strong>recommendations</strong> — nothing is written until you approve it. For each
        one, approve (optionally substituting a value from the field&apos;s own list), reject, or
        return it to the servicer. The corrected workbook is generated from your decisions, from the
        original uploaded file; the original is never modified.
      </p>

      {nothingToReview && (
        <div className="alert" style={{ marginBottom: 12 }}>
          <strong>Nothing to review — this audit produced no recommendations.</strong>
          <div className="hint" style={{ marginTop: 4 }}>
            That is a complete review (0 of 0) and is ready for final reporting. Enter
            your name and record it to generate the Monthly Servicer Report and the
            Final Audited Portfolio Summary.
          </div>
        </div>
      )}

      {/* Workload summary + global affordances. Every figure is counted from the
          classified response — nothing here is hard-coded. A gate tedious enough
          to route around is not a control: at 127 recommendations and rising,
          "reject all remaining" and "jump to next undecided" are what make
          finishing a review feasible. */}
      {!nothingToReview && (
        <div className="map-row" style={{ gridTemplateColumns: '1.6fr 1fr', marginBottom: 12 }}>
          <div className="field-name">
            {recs.length} Recommendation{recs.length === 1 ? '' : 's'}
            <div className="hint" style={{ marginTop: 2 }}>
              {standardRecs.length} Standard Correction{standardRecs.length === 1 ? '' : 's'} ·{' '}
              {reviewRecs.length} Review Required · {decidedCount} decided · {pendingCount}{' '}
              pending
            </div>
            <div className="hint" style={{ marginTop: 2 }}>
              {pendingCount > 0
                ? 'Final reporting is blocked until every recommendation reaches a reportable disposition.'
                : 'All recommendations decided. Reporting readiness is confirmed by the server on submission.'}
            </div>
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <button className="btn" disabled={pendingCount === 0} onClick={jumpToNextUndecided}>
              Jump to next undecided
            </button>
            <button className="btn" disabled={pendingCount === 0} onClick={rejectAllRemaining}>
              Reject all remaining ({pendingCount})
            </button>
          </div>
        </div>
      )}

      {standardRecs.length > 0 && (
        <div style={{ marginBottom: 18 }}>
          <h4 style={{ marginBottom: 4 }}>
            STANDARD CORRECTIONS · {standardRecs.length} ·{' '}
            {standardRecs.filter((r) => decisionOf(r.id).decision).length} decided
          </h4>
          <p className="hint" style={{ marginTop: 0 }}>
            Deterministic Delinquency Status corrections. The existing business rule already
            determines the value: STAT-01 calculated it from missed-payment history, no
            data-quality warning applies, and there is no reversal, negative posting, invalid
            reported value, or terminal-status conflict. Every Loan Status recommendation is
            Review Required and is never included here.
          </p>
          <div
            style={{
              display: 'flex',
              gap: 10,
              alignItems: 'center',
              flexWrap: 'wrap',
              marginBottom: 8,
            }}
          >
            <button
              className="btn btn-teal"
              disabled={bulkEligible.length === 0}
              onClick={approveAllStandard}
            >
              Approve All Standard Corrections ({bulkEligible.length})
            </button>
            <span className="hint">
              Pre-fills the decision for {bulkEligible.length} undecided Standard
              recommendation{bulkEligible.length === 1 ? '' : 's'}. Nothing is submitted or
              written until you enter your name and press Record — and every one stays
              individually editable.
            </span>
          </div>
          {bulkHeld.length > 0 && (
            <div className="alert" style={{ marginBottom: 8 }}>
              <strong>
                {bulkHeld.length} Standard recommendation{bulkHeld.length === 1 ? '' : 's'} held
                back from bulk approval.
              </strong>
              <div className="hint" style={{ marginTop: 4 }}>
                {bulkHeld.map((r) => `loan ${r.gp} (${r.servicer})`).join(', ')} — this loan also
                has an undecided <strong>Loan Status</strong> recommendation. Decide the Loan
                Status on its loan card under Review Required and this becomes eligible, or
                approve it individually there. It is still counted as a Standard correction.
              </div>
            </div>
          )}
          {standardCards.length > 0 ? (
            <ReviewTable
              cards={standardCards}
              decisionOf={decisionOf}
              onChange={setDecision}
            />
          ) : (
            <p className="hint">
              Every Standard correction belongs to a loan that also needs individual review — see
              the loan cards under Review Required.
            </p>
          )}
        </div>
      )}

      {reviewRecs.length > 0 && (
        <div>
          <h4 style={{ marginBottom: 4 }}>
            REVIEW REQUIRED · {reviewRecs.length} ·{' '}
            {reviewRecs.filter((r) => decisionOf(r.id).decision).length} decided
          </h4>
          <p className="hint" style={{ marginTop: 0 }}>
            Each of these needs individual judgement and none can be bulk-approved.
          </p>
          {REVIEW_REQUIRED_CATEGORIES.map((cat) => {
            const cards = reviewCardsByCategory[cat] || []
            // An empty category is not rendered at all — the sections come from the
            // data, so a month with no reversals simply has no Reversal section.
            if (cards.length === 0) return null
            const isCollapsed = !!collapsed[cat]
            return (
              <div key={cat} style={{ marginBottom: 14 }}>
                <button
                  className="btn"
                  aria-expanded={!isCollapsed}
                  onClick={() => setCollapsed((c) => ({ ...c, [cat]: !c[cat] }))}
                  style={{ marginBottom: 6 }}
                >
                  {isCollapsed ? '▶' : '▼'} {cat} · {countIn(cards)} · {decidedIn(cards)} decided
                </button>
                {!isCollapsed && (
                  <>
                    {CATEGORY_BLURB[cat] && (
                      <p className="hint" style={{ marginTop: 0 }}>
                        {CATEGORY_BLURB[cat]}
                      </p>
                    )}
                    <ReviewTable cards={cards} decisionOf={decisionOf} onChange={setDecision} />
                  </>
                )}
              </div>
            )
          })}
        </div>
      )}

      <div style={{ marginTop: 16, display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
        <label>
          Reviewed by{' '}
          <input
            type="text"
            value={reviewer}
            placeholder="your name"
            onChange={(e) => setReviewer(e.target.value)}
          />
        </label>
        {/* No authentication exists — the recorded name is self-declared and is
            written to the Correction Log labelled "(unverified)". */}
        <span className="hint">Recorded as self-declared (unverified) — no authentication.</span>
        <button
          className="btn btn-teal"
          onClick={submit}
          disabled={(blocked && !nothingToReview) || !reviewer.trim() || submitting
            || approvalsMissingValue.length > 0}
        >
          {submitting
            ? 'Recording…'
            : nothingToReview
              ? 'Record review (nothing to review)'
              : `Record ${decidedCount} decision${decidedCount === 1 ? '' : 's'}`}
        </button>
        {!submitting && !reviewer.trim() && (
          <span className="hint">A reviewer name is required.</span>
        )}
        {!submitting && reviewer.trim() && !nothingToReview && decidedCount === 0 && (
          <span className="hint">Decide at least one recommendation.</span>
        )}
        {!submitting && approvalsMissingValue.length > 0 && (
          <span className="hint" style={{ color: 'var(--amber, #BA7517)', fontWeight: 600 }}>
            Select a value for{' '}
            {approvalsMissingValue.map((r) => `loan ${r.gp} ${r.field}`).join(', ')}. An
            approval with no value cannot be applied, so final reporting would stay
            blocked.
          </span>
        )}
        {/* The balance decision is REQUIRED, so it is called out separately from
            the generic missing-value hint: an undecided one leaves the liquidation
            reversal unresolved and blocks final reporting. */}
        {!submitting && balanceDecisionsPending.length > 0 && (
          <span className="hint" style={{ color: 'var(--amber, #BA7517)', fontWeight: 600 }}>
            A required balance decision is outstanding for{' '}
            {balanceDecisionsPending.map((r) => `loan ${r.gp}`).join(', ')}. Choose
            whether to restore the Ending Balance or keep it at 0.00 — until then the
            liquidation reversal is not fully resolved and final reporting stays
            blocked.
          </span>
        )}
        {/* D4: pending items do NOT block submission -- they block FINAL
            REPORTING. Said plainly here so the reviewer knows a partial
            submission is legitimate and what it will and will not produce. */}
        {!submitting && reviewer.trim() && decidedCount > 0 && pendingCount > 0 && (
          <span className="hint">
            <strong>
              {decidedCount} of {recs.length} reviewed, {pendingCount} pending. Final
              reporting is blocked.
            </strong>{' '}
            Recording now still produces corrected workbooks and Correction Logs for
            the decisions you have made.
          </span>
        )}
      </div>

      {error && (
        <div style={{ marginTop: 12, color: 'var(--red, #A32D2D)' }}>{error}</div>
      )}

      {result && (
        <div style={{ marginTop: 16 }}>
          <h4 style={{ marginBottom: 6 }}>
            Recorded: {result.summary.applied} applied · {result.summary.rejected} rejected ·{' '}
            {result.summary.returned} returned
          </h4>
          {(result.corrected_workbooks || []).map((cw) => (
            <div className="map-row" style={{ gridTemplateColumns: '1.4fr 1fr' }} key={cw.filename}>
              <div className="field-name">
                {cw.servicer}
                <div className="hint" style={{ marginTop: 2 }}>
                  {cw.applied_count} applied · {cw.rejected_count} rejected · {cw.returned_count} returned
                </div>
              </div>
              <button className="btn btn-teal" onClick={() => downloadBase64File(cw.file_b64, cw.filename)}>
                Download Corrected File
              </button>
            </div>
          ))}
          {/* The FINAL AUDITED Portfolio Summary — the third layer, available
              only after Manual Review. Built from the same verified values as
              the corrected workbooks above: original lender data + the
              reviewer's approved corrections + closed-at-zero suppression.
              Carries one row per uploaded servicer, including servicers with
              no corrections. Distinct from the Original Portfolio Summary
              (what the lender reported, on the results screen) and the
              Audited Portfolio Summary (what the audit calculated, downloaded
              from the results screen before review). */}
          {/* Monthly Servicer Reports — one workbook per SOUP deal, 2024 then
              2025, never combined. Consolidation of the same reviewer-approved
              data the corrected workbooks carry; no figure is recalculated
              here. Ordered between the corrected workbooks and the Final
              Audited Portfolio Summary, matching the client's stated
              post-review sequence. */}
          {(result.monthly_servicer_reports || []).map((report) => (
            <div
              key={report.deal || report.filename}
              className="map-row"
              style={{ gridTemplateColumns: '1.4fr 1fr', marginTop: 10 }}
            >
              <div className="field-name">
                {report.note_series ? `${report.note_series} ` : ''}Monthly Servicer Report
                <div className="hint" style={{ marginTop: 2 }}>
                  {report.loan_count} loans · {(report.servicers || []).join(', ')}
                </div>
              </div>
              <button
                className="btn btn-teal"
                onClick={() =>
                  downloadBase64File(
                    report.file_b64,
                    report.filename || 'Monthly_Servicer_Report.xlsx'
                  )
                }
              >
                Download Monthly Servicer Report
              </button>
            </div>
          ))}
          {/* D4 — review completeness and reporting readiness, BOTH computed
              server-side and only read here. They can legitimately disagree: a
              review can be 100% decided and still not reportable, when a
              recommendation was returned to the servicer or an approval could not
              be applied. Corrected workbooks above are returned either way. */}
          {result.reporting_readiness && !result.reporting_readiness.reporting_ready && (
            <div className="alert alert-error" style={{ marginTop: 10 }}>
              <strong>Final reporting is blocked.</strong>
              <div className="hint" style={{ marginTop: 4 }}>
                {result.review_completeness && (
                  <>
                    {result.review_completeness.approved_applied +
                      result.review_completeness.rejected +
                      result.review_completeness.returned +
                      result.review_completeness.approved_not_applied}{' '}
                    of {result.review_completeness.total} reviewed,{' '}
                    {result.review_completeness.pending} pending.{' '}
                  </>
                )}
                No Monthly Servicer Report and no Final Audited Portfolio Summary was
                produced. Your corrected workbooks and Correction Logs above are
                complete for the decisions you recorded.
              </div>
              {(result.reporting_readiness.blocking_reasons || []).includes('returned') && (
                <div className="hint" style={{ marginTop: 4 }}>
                  {result.review_completeness.returned} recommendation(s) were{' '}
                  <strong>returned to the servicer</strong> — reviewed, but the value is
                  awaiting a resubmission, so it cannot be reported as final.
                </div>
              )}
              {(result.reporting_readiness.blocking_reasons || []).includes('approved_not_applied') && (
                <div className="hint" style={{ marginTop: 4 }}>
                  {result.review_completeness.approved_not_applied} approval(s) could{' '}
                  <strong>not be applied</strong> (value refused, or no verifiable row to
                  write to). The review is decided but the workbook does not contain the
                  approved value. This cannot be cleared from this page — the source file
                  needs correcting and the audit re-running.
                </div>
              )}
              {(result.reporting_readiness.blocking_reasons || []).includes('pending') && (
                <div className="hint" style={{ marginTop: 4 }}>
                  Decide the remaining {result.review_completeness.pending} recommendation(s)
                  and record again to produce the reports.
                </div>
              )}
            </div>
          )}
          {result.reporting_readiness && result.reporting_readiness.reporting_ready &&
           result.review_completeness && (
            <div className="hint" style={{ marginTop: 10 }}>
              Review complete and reporting ready — {result.review_completeness.total} of{' '}
              {result.review_completeness.total} at a reportable disposition (
              {result.review_completeness.approved_applied} applied,{' '}
              {result.review_completeness.rejected} rejected).
            </div>
          )}
          {/* D3 — per-deal blocked state. A SOUP deal whose expected servicers
              did not all reach the report produces NO workbook, and says so
              here as a prominent panel rather than a hint. Rendered per deal:
              a blocked SOUP 2024 never suppresses a complete SOUP 2025, so this
              panel can sit beside a successful download above. The corrected
              workbooks are unaffected and remain downloadable. */}
          {(result.deals || []).filter((d) => d.status === 'blocked').map((d) => (
            <div
              key={`blocked-${d.deal || 'unassigned'}`}
              className="alert alert-error"
              style={{ marginTop: 10 }}
            >
              <strong>
                {d.deal || 'Unassigned deal'} — no Monthly Servicer Report produced
              </strong>
              <div className="hint" style={{ marginTop: 4 }}>
                Expected {(d.servicers_expected || []).length} servicer(s):{' '}
                {(d.servicers_expected || []).join(', ') || '—'}. Reported{' '}
                {(d.servicers_reported || []).length}:{' '}
                {(d.servicers_reported || []).join(', ') || 'none'}.
              </div>
              {(d.failures || []).map((f, i) => (
                <div key={i} className="hint" style={{ marginTop: 2 }}>
                  {f.servicer} — failed at stage <strong>{f.stage}</strong>:{' '}
                  {f.error_type}: {f.error_detail}
                  {f.filename ? ` (${f.filename})` : ''}
                </div>
              ))}
              <div className="hint" style={{ marginTop: 4 }}>
                A partial report is not produced for a blocked deal. Fix the file
                and re-run the audit; other deals are unaffected and your
                corrected workbooks above are unchanged.
              </div>
            </div>
          ))}
          {(result.monthly_servicer_report_validation_issues || []).length > 0 && (
            <div style={{ marginTop: 10 }}>
              <strong>Monthly Servicer Report — needs your attention</strong>
              {result.monthly_servicer_report_validation_issues.map((issue, i) => (
                <div key={i} className="hint">
                  {issue.field ? `${issue.field}: ` : ''}
                  {issue.message}
                </div>
              ))}
            </div>
          )}
          {result.portfolio_summary_final_b64 && (
            <div className="map-row" style={{ gridTemplateColumns: '1.4fr 1fr', marginTop: 10 }}>
              <div className="field-name">
                Final Audited Portfolio Summary
                <div className="hint" style={{ marginTop: 2 }}>
                  All servicers, built from your verified review decisions
                </div>
              </div>
              <button
                className="btn btn-teal"
                onClick={() =>
                  downloadBase64File(
                    result.portfolio_summary_final_b64,
                    result.portfolio_summary_final_filename ||
                      'Portfolio_Summary_Final_Audited.xlsx'
                  )
                }
              >
                Download Final Audited Portfolio Summary
              </button>
            </div>
          )}
          {(result.corrections_not_applied || []).length > 0 && (
            <div style={{ marginTop: 10 }}>
              <strong>Not applied</strong>
              {result.corrections_not_applied.map((item, i) => (
                <div key={i} className="hint">
                  {item.servicer}
                  {item['SBA GP Number'] ? ` · Loan ${item['SBA GP Number']}` : ''}
                  {item.field ? ` · ${item.field}` : ''}: {item.reason}
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  )
}
