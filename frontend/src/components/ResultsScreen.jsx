import React, { useState } from 'react'

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

export default function ResultsScreen({ audit, onNewAudit, onExit }) {
  const { results, summary, report_id } = audit
  const correctedWorkbooks = audit.corrected_workbooks || []
  const correctionsSkipped = audit.corrections_skipped || []
  const [expanded, setExpanded] = useState(null)
  const timestamp = new Date().toLocaleString()

  const toggle = (id) => setExpanded((cur) => (cur === id ? null : id))

  const handleDownload = () => {
    downloadBase64File(audit.report_b64, audit.report_filename || 'Audit_Intelligence_Report.xlsx', XLSX_MIME)
  }

  // CLIENT-CONFIRMED (2026-08-04, revised): TWO SEPARATE downloadable
  // workbooks, both independent of the detailed reconciliation report
  // above — Original uses the servicer's own reported Delinquency Status
  // for the aging buckets, Audited uses STAT-01's reconciled status. Each
  // is only present when the backend actually produced one (the b64
  // field is null if no servicer record was available, e.g. every
  // uploaded file failed to parse before that point).
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
            {results.map((r) => (
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

      {/* Corrected servicer workbooks — only shown when the audit engine
          found at least one high-confidence, auto-correctable field for a
          servicer's file. Absent entirely when there's nothing eligible,
          rather than showing an empty or misleading section. */}
      {correctedWorkbooks.length > 0 && (
        <div className="card" style={{ marginBottom: 22 }}>
          <h3 style={{ marginTop: 0 }}>Corrected servicer workbooks</h3>
          <p style={{ color: 'var(--text-secondary)', marginTop: 0 }}>
            A copy of each servicer file below with only its high-confidence corrections applied,
            plus a Correction Log tab documenting every change. The original uploaded file is never modified.
          </p>
          {correctedWorkbooks.map((cw) => (
            <div className="map-row" style={{ gridTemplateColumns: '1.4fr 1fr' }} key={cw.filename}>
              <div className="field-name">
                {cw.servicer}
                <div className="hint" style={{ marginTop: 2 }}>
                  {cw.correction_count} correction{cw.correction_count === 1 ? '' : 's'} applied
                </div>
              </div>
              <a
                className="btn btn-teal"
                onClick={() => downloadBase64File(cw.file_b64, cw.filename, XLSX_MIME)}
              >
                Download Corrected File
              </a>
            </div>
          ))}
        </div>
      )}

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
