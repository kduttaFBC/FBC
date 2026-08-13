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
const LOAN_STATUS_DOMAIN = [
  'ACTIVE', 'DEFERRED', 'PAID-IN-FULL', 'REO', 'FORECLOSURE',
  'BANKRUPTCY', 'SUBSTITUTION-IN', 'SUBSTITUTION-OUT', 'LIQUIDATION',
]

const DECISIONS = [
  { v: 'approve', label: 'Approve — write to workbook' },
  { v: 'reject', label: 'Reject — keep the reported value' },
  { v: 'return', label: 'Return to servicer — source defect' },
]

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

function domainFor(field) {
  return field === 'Loan Status' ? LOAN_STATUS_DOMAIN : DQ_LADDER
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

function dash(value) {
  return value === null || value === undefined || value === '' ? '—' : String(value)
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

export default function CorrectionReview({ recommendations, sessionId }) {
  const recs = recommendations || []
  const [decisions, setDecisions] = useState({}) // id -> {decision, approved_value, note, notify}
  const [reviewer, setReviewer] = useState('')
  const [submitting, setSubmitting] = useState(false)
  const [error, setError] = useState('')
  const [result, setResult] = useState(null)

  const decisionOf = (id) => decisions[id] || {}
  const setDecision = (id, patch) =>
    setDecisions((cur) => ({ ...cur, [id]: { ...cur[id], ...patch } }))

  const decidedCount = useMemo(
    () => recs.filter((r) => decisionOf(r.id).decision).length,
    [recs, decisions]
  )
  const pendingCount = recs.length - decidedCount
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

  // D4 (CLIENT-CONFIRMED 2026-08-13): a clean month must still be able to record
  // its review and produce its reports. This component previously returned null
  // when there was nothing to review, so the reviewer-name field and the submit
  // button never rendered at all -- there was no path from a clean audit to a
  // Monthly Servicer Report or a Final Audited Portfolio Summary. 0 of 0 is the
  // most complete review possible, and it is reporting-ready.
  // NOTE: no early return here any more. `nothingToReview` drives the table and
  // the bulk controls instead, so the submit affordance is always present.

  return (
    <div className="card" style={{ marginBottom: 22 }}>
      <h3 style={{ marginTop: 0 }}>Review corrections</h3>
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

      {/* D4 progress + bulk affordances. A gate tedious enough to route around is
          not a control: at 121 recommendations and rising, "reject all remaining"
          and "jump to next undecided" are what make finishing a review feasible. */}
      {!nothingToReview && (
        <div className="map-row" style={{ gridTemplateColumns: '1.6fr 1fr', marginBottom: 12 }}>
          <div className="field-name">
            {decidedCount} of {recs.length} reviewed
            {pendingCount > 0 ? ` · ${pendingCount} pending` : ' · none pending'}
            <div className="hint" style={{ marginTop: 2 }}>
              {pendingCount > 0
                ? 'Final reporting is blocked until every recommendation reaches a reportable disposition.'
                : 'All recommendations decided. Reporting readiness is confirmed by the server on submission.'}
            </div>
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <button
              className="btn"
              disabled={pendingCount === 0}
              onClick={() => {
                const next = recs.find((r) => !decisionOf(r.id).decision)
                if (!next) return
                const el = document.getElementById(`rec-${next.id}`)
                if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center' })
              }}
            >
              Jump to next undecided
            </button>
            <button
              className="btn"
              disabled={pendingCount === 0}
              onClick={() =>
                setDecisions((cur) => {
                  const out = { ...cur }
                  recs.forEach((r) => {
                    if (!(out[r.id] && out[r.id].decision)) {
                      out[r.id] = { ...out[r.id], decision: 'reject' }
                    }
                  })
                  return out
                })
              }
            >
              Reject all remaining ({pendingCount})
            </button>
          </div>
        </div>
      )}

      <table className="detail-table" style={{ width: '100%' }}>
        <thead>
          <tr>
            <th>Loan</th>
            <th>Field</th>
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
          {recs.map((r) => {
            const d = decisionOf(r.id)
            const units = paymentUnits(r)
            const reversal = reversalEvidence(r)
            const currentPeriod = currentPeriodEvidence(r)
            return (
              <React.Fragment key={r.id}>
                {/* id is the jump-to-next-undecided anchor. */}
                <tr id={`rec-${r.id}`}>
                  <td>
                    <strong>{r.gp}</strong>
                    <div className="hint">{r.servicer}{r.postings ? ` · ${r.postings} posting${r.postings === 1 ? '' : 's'}` : ''}</div>
                  </td>
                  <td>{r.field}</td>
                  <td style={{ fontFamily: 'monospace' }}>{String(r.reported)}</td>
                  <td style={{ fontFamily: 'monospace' }}>{r.prior == null ? '—' : String(r.prior)}</td>
                  <td style={{ fontFamily: 'monospace' }}>{String(r.calculated)}</td>
                  <td className="hint">
                    {reversal ? (
                      <>
                        <strong>Reversal</strong> — net Received {money(r.paymentReceived)} / Sched{' '}
                        {money(r.scheduledPayment)}
                        <div>
                          Deficiency {money(reversal.deficiency)} / {money(reversal.thresholdUnit)}{' '}
                          (80% of sched) ={' '}
                          <strong>
                            {reversal.equivalents} payment
                            {reversal.equivalents === 1 ? '' : 's'} removed
                          </strong>
                        </div>
                        <div>Fee {money(r.servicingFee)} — excluded from this calculation</div>
                      </>
                    ) : (
                      <>
                        Received {money(r.paymentReceived)} + Fee {money(r.servicingFee)} / Sched{' '}
                        {money(r.scheduledPayment)}
                        {units != null && (
                          <>
                            {' '}= <strong>{units.toFixed(2)} units</strong>
                          </>
                        )}
                        {/* An OFFSETTING reversal keeps the ratio above: its net
                            really is $0.00 and the ordinary no-payment branch
                            really did produce the calculated value, so unlike a
                            net-negative reversal the figure on screen is the
                            arithmetic that ran. What the $0.00 hides is that
                            reversed postings are inside it — without this line a
                            reviewer cannot tell this apart from a loan that
                            simply paid nothing. */}
                        {r.offsettingZeroNetReversal && (
                          <div>
                            <strong>Includes reversed postings</strong> offset to a net of exactly
                            $0.00
                          </div>
                        )}
                      </>
                    )}
                    <div>Balance {money(r.balance)}</div>
                  </td>
                  <td>
                    <select
                      aria-label={`Decision for loan ${r.gp}`}
                      value={d.decision || ''}
                      onChange={(e) => setDecision(r.id, { decision: e.target.value })}
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
                    {d.decision === 'approve' ? (
                      <select
                        aria-label={`Approved value for loan ${r.gp}`}
                        value={d.approved_value || r.calculated}
                        onChange={(e) => setDecision(r.id, { approved_value: e.target.value })}
                      >
                        {domainFor(r.field).map((v) => (
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
                      aria-label={`Notify servicer for loan ${r.gp}`}
                      checked={!!d.notify}
                      onChange={(e) => setDecision(r.id, { notify: e.target.checked })}
                    />
                  </td>
                  <td>
                    <input
                      type="text"
                      aria-label={`Reviewer note for loan ${r.gp}`}
                      value={d.note || ''}
                      placeholder="optional"
                      onChange={(e) => setDecision(r.id, { note: e.target.value })}
                      style={{ width: 140 }}
                    />
                  </td>
                </tr>
                {currentPeriod && (
                  <tr>
                    <td colSpan={10} className="hint">
                      <strong>Current-period evidence:</strong>{' '}
                      Bankruptcy Date {dash(currentPeriod.bankruptcy)} · Paid-In-Full Date{' '}
                      {dash(currentPeriod.paidInFull)} · Liquidation Date{' '}
                      {dash(currentPeriod.liquidation)} · Ending Balance {money(r.balance)}
                      {currentPeriod.contradiction && (
                        <div style={{ color: 'var(--amber, #BA7517)', fontWeight: 600, marginTop: 4 }}>
                          ⚠ CONTRADICTORY: a Liquidation Date of {currentPeriod.liquidation} is
                          reported against an Ending Balance of {money(r.balance)} (greater than
                          $0), with a reported Loan Status of {dash(r.servicerLoanStatus)}. The
                          zero-balance exception does not apply — weigh both before approving.
                        </div>
                      )}
                    </td>
                  </tr>
                )}
                {hasDeferralEvidence(r) && (
                  <tr>
                    <td colSpan={10} className="hint">
                      <strong>Deferral evidence:</strong>{' '}
                      <span style={{ color: EVIDENCE_COLORS[r.evidenceStatus] || 'inherit' }}>
                        {r.evidenceStatus}
                      </span>
                      {EVIDENCE_MEANING[r.evidenceStatus] && (
                        <> — {EVIDENCE_MEANING[r.evidenceStatus]}</>
                      )}
                      <div>
                        Memo action {dash(r.memoSecuritizationAction)} · Window{' '}
                        {dash(r.deferralStart)} → {dash(r.deferralEnd)} · Resume{' '}
                        {dash(r.deferralResume)} · Prior (SOUP) Loan Status{' '}
                        {dash(r.priorLoanStatus)} · Servicer Loan Status{' '}
                        {dash(r.servicerLoanStatus)}
                      </div>
                    </td>
                  </tr>
                )}
                {r.reviewFlag && (
                  <tr>
                    <td colSpan={10} style={{ color: 'var(--amber, #BA7517)' }}>
                      ⚠ {r.reviewFlag}
                    </td>
                  </tr>
                )}
              </React.Fragment>
            )
          })}
        </tbody>
      </table>

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
          disabled={(blocked && !nothingToReview) || !reviewer.trim() || submitting}
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
