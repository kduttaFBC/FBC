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
const LOAN_STATUS_DOMAIN = ['ACTIVE', 'DEFERRED']

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
  const sched = Number(rec.scheduledPayment)
  if (!sched) return null
  const adjusted = Number(rec.paymentReceived || 0) + Number(rec.servicingFee || 0)
  return adjusted / sched
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
  const blocked = !reviewer.trim() || decidedCount === 0

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

  if (!recs.length) return null

  return (
    <div className="card" style={{ marginBottom: 22 }}>
      <h3 style={{ marginTop: 0 }}>Review corrections</h3>
      <p style={{ color: 'var(--text-secondary)', marginTop: 0 }}>
        These are <strong>recommendations</strong> — nothing is written until you approve it. For each
        one, approve (optionally substituting a value from the field&apos;s own list), reject, or
        return it to the servicer. The corrected workbook is generated from your decisions, from the
        original uploaded file; the original is never modified.
      </p>

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
            return (
              <React.Fragment key={r.id}>
                <tr>
                  <td>
                    <strong>{r.gp}</strong>
                    <div className="hint">{r.servicer}{r.postings ? ` · ${r.postings} posting${r.postings === 1 ? '' : 's'}` : ''}</div>
                  </td>
                  <td>{r.field}</td>
                  <td style={{ fontFamily: 'monospace' }}>{String(r.reported)}</td>
                  <td style={{ fontFamily: 'monospace' }}>{r.prior == null ? '—' : String(r.prior)}</td>
                  <td style={{ fontFamily: 'monospace' }}>{String(r.calculated)}</td>
                  <td className="hint">
                    Received {money(r.paymentReceived)} + Fee {money(r.servicingFee)} / Sched{' '}
                    {money(r.scheduledPayment)}
                    {units != null && (
                      <>
                        {' '}= <strong>{units.toFixed(2)} units</strong>
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
        <button className="btn btn-teal" onClick={submit} disabled={blocked || submitting}>
          {submitting ? 'Recording…' : `Record ${decidedCount} decision${decidedCount === 1 ? '' : 's'}`}
        </button>
        {blocked && !submitting && (
          <span className="hint">
            {reviewer.trim() ? 'Decide at least one recommendation.' : 'A reviewer name is required.'}
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
