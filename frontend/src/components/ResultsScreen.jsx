import React, { useState } from 'react'
import CorrectionReview from './CorrectionReview.jsx'

function DetailTable({ records }) {
  if (!records || !records.length) {
    return <div className="no-issues">✓ No issues found for this check.</div>
  }
  const headers = Object.keys(records[0])
  return (
    <table className="detail-table">
      <thead>
        <tr>{headers.map((h) => <th key={h}>{h}</th>)}</tr>
      </thead>
      <tbody>
        {records.map((rec, i) => (
          <tr key={i}>
            {headers.map((h) => <td key={h}>{String(rec[h])}</td>)}
          </tr>
        ))}
      </tbody>
    </table>
  )
}

function downloadBase64File(base64, filename, mimeType) {
  const byteCharacters = atob(base64)
  const byteNumbers = Array.from(byteCharacters, c => c.charCodeAt(0))
  const byteArray = new Uint8Array(byteNumbers)
  const blob = new Blob([byteArray], { type: mimeType })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

// CLIENT-CONFIRMED (2026-08-11): checks that are computed and returned in full
// by the API, but are NOT rendered as a row in the on-screen Check results
// table because a dedicated client-facing artifact already serves their
// purpose.
//
// PORT-01 is the Original Portfolio Summary. It has its own downloadable
// workbook (see the Download Original Portfolio Summary button above), so a
// duplicate on-screen row showing "INFO / 0 issues" per servicer -- eight rows
// on the June set -- adds nothing a reader can act on.
//
// DISPLAY ONLY. PORT-01 stays in all_results and in the /api/audit response
// unchanged, because main.py builds the Original Portfolio Summary workbook
// FROM that record, and because this app's stated contract is that every
// check's complete result is always in the JSON regardless of what any
// client-facing view chooses to show (see main.py _filter_report_results).
// Hiding a row here must never be read as the check not having run.
export const CHECKS_HIDDEN_FROM_RESULTS_TABLE = ['PORT-01']

// Exported as a pure function of the results list so the hiding rule can be
// asserted directly rather than inferred from rendered output.
export function visibleResults(results) {
  return (results || []).filter(
    (r) => !CHECKS_HIDDEN_FROM_RESULTS_TABLE.includes(String(r.id || '').split('::')[0])
  )
}

export default function ResultsScreen({ audit, onNewAudit, onExit }) {
  const { results, summary, report_id } = audit
  const recommendations = audit.recommendations || []
  const correctionsSkipped = audit.corrections_skipped || []
  const [expanded, setExpanded] = useState(null)
  const timestamp = new Date().toLocaleString()

  const toggle = (id) => setExpanded((cur) => (cur === id ? null : id))

  const handleDownload = () => {
    downloadBase64File(audit.report_b64, audit.report_filename || 'Audit_Intelligence_Report.xlsx', XLSX_MIME)
  }

  // CLIENT-CONFIRMED (2026-08-04, revised 2026-08-11): TWO SEPARATE
  // downloadable workbooks, both independent of the detailed reconciliation
  // report above — Original uses the servicer's own reported Delinquency
  // Status for the aging buckets, Audited uses STAT-01's reconciled status.
  // Each is only present when the backend actually produced one (the b64
  // field is null if no servicer record was available, e.g. every uploaded
  // file failed to parse before that point).
  //
  // ORDER IS THE POINT: Original is the PROMINENT top-level download and is
  // rendered first, because it is the lender-submitted starting position the
  // client reads before any audit result. Audited is retained immediately
  // after it as a secondary download. Renaming or reordering these buttons is
  // a client-facing decision, not a cosmetic one.
  const hasPortfolioSummaryOriginal = Boolean(audit.portfolio_summary_original_b64)
  const handleDownloadPortfolioSummaryOriginal = () => {
    downloadBase64File(
      audit.portfolio_summary_original_b64,
      audit.portfolio_summary_original_filename || 'Portfolio_Summary_Original.xlsx',
      XLSX_MIME
    )
  }

  const hasPortfolioSummaryAudited = Boolean(audit.portfolio_summary_audited_b64)
  const handleDownloadPortfolioSummaryAudited = () => {
    downloadBase64File(
      audit.portfolio_summary_audited_b64,
      audit.portfolio_summary_audited_filename || 'Portfolio_Summary_Audited.xlsx',
      XLSX_MIME
    )
  }

  return (
    <>
      {/* Summary card */}
      <div className="card" style={{ marginBottom: 22 }}>
        <div className="summary-head">
          <div>
            <h2>Audit complete</h2>
            <div className="ts">{timestamp}</div>
          </div>
        
            <a className="btn btn-teal" onClick={handleDownload} target="_blank" rel="noreferrer">
            Download Excel Report
          </a>
          {hasPortfolioSummaryOriginal && (
            <a
              className="btn btn-teal"
              onClick={handleDownloadPortfolioSummaryOriginal}
              target="_blank"
              rel="noreferrer"
              style={{ marginLeft: 10 }}
            >
              Download Original Portfolio Summary
            </a>
          )}
          {hasPortfolioSummaryAudited && (
            <a
              className="btn btn-teal"
              onClick={handleDownloadPortfolioSummaryAudited}
              target="_blank"
              rel="noreferrer"
              style={{ marginLeft: 10 }}
            >
              Download Audited Portfolio Summary
            </a>
          )}
        </div>

        <div className="scorecards">
          <div className="scorecard blue">
            <div className="num">{summary.total_records}</div>
            <div className="label">Total records checked</div>
          </div>
          <div className="scorecard red">
            <div className="num">{summary.total_issues}</div>
            <div className="label">Total issues found</div>
          </div>
          <div className="scorecard green">
            <div className="num">{summary.checks_passed}</div>
            <div className="label">Checks passed</div>
          </div>
          <div className="scorecard amber">
            <div className="num">{summary.checks_flagged}</div>
            <div className="label">Checks flagged</div>
          </div>
        </div>
      </div>

      {/* Check results table */}
      <div className="card" style={{ marginBottom: 22 }}>
        <h3 style={{ marginTop: 0 }}>Check results</h3>
        <table className="results-table">
          <thead>
            <tr>
              <th></th>
              <th>Check</th>
              <th>Category</th>
              <th>Status</th>
              <th>Issues Found</th>
              <th>Details</th>
            </tr>
          </thead>
          <tbody>
            {visibleResults(results).map((r) => (
              <React.Fragment key={r.id}>
                <tr className="clickable" onClick={() => toggle(r.id)}>
                  <td className="caret">{expanded === r.id ? '▾' : '▸'}</td>
                  <td><strong>{r.id}</strong> — {r.name}</td>
                  <td>{r.category}</td>
                  <td><span className={`pill ${r.status}`}>{r.status}</span></td>
                  <td>{r.issues_found}</td>
                  <td>{r.details}</td>
                </tr>
                {expanded === r.id && (
                  <tr className="expand-row">
                    <td colSpan={6}>
                      <div className="expand-inner">
                        <DetailTable records={r.records} />
                      </div>
                    </td>
                  </tr>
                )}
              </React.Fragment>
            ))}
          </tbody>
        </table>
      </div>

      {/* Correction review — the audit produces recommendations only; the
          reviewer decides each one and the corrected workbook is generated
          from those decisions (POST /api/corrections/review). Absent entirely
          when there is nothing to review. */}
      <CorrectionReview recommendations={recommendations} sessionId={audit.session_id} />

      {correctionsSkipped.length > 0 && (
        <div className="card" style={{ marginBottom: 22 }}>
          <h3 style={{ marginTop: 0 }}>Corrections Not Applied</h3>
          {correctionsSkipped.map((item, index) => {
            const gp = item['SBA GP Number']
            const detailParts = []
            if (gp) detailParts.push(`Loan ${gp}`)
            if (item.field) detailParts.push(item.field)
            const detail = detailParts.join(' — ')
            return (
              <div key={`${item.servicer}-${index}`} style={{ marginBottom: 10 }}>
                <strong>{item.servicer}</strong>
                {detail && <span style={{ marginLeft: 6, color: 'var(--text-secondary)' }}>({detail})</span>}
                <div className="hint">{item.reason}</div>
              </div>
            )
          })}
        </div>
      )}

      <div className="btn-row" style={{ marginBottom: 8 }}>
        <button className="btn btn-primary" onClick={onNewAudit}>Upload New Files</button>
        <button className="btn btn-secondary" onClick={onExit}>Exit</button>
      </div>
    </>
  )
}
