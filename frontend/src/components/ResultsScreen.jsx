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

export default function ResultsScreen({ audit, onNewAudit, onExit }) {
  const { results, summary, report_id } = audit
  const [expanded, setExpanded] = useState(null)
  const timestamp = new Date().toLocaleString()

  const toggle = (id) => setExpanded((cur) => (cur === id ? null : id))

const handleDownload = () => {
  console.log('report_b64 exists:', !!audit.report_b64, 'length:', audit.report_b64?.length)
  const byteCharacters = atob(audit.report_b64)
  const byteNumbers = Array.from(byteCharacters, c => c.charCodeAt(0))
  const byteArray = new Uint8Array(byteNumbers)
  const blob = new Blob([byteArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = audit.report_filename || 'FalconBridge_Audit.xlsx'
  a.click()
  URL.revokeObjectURL(url)
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

      <div className="btn-row" style={{ marginBottom: 8 }}>
        <button className="btn btn-primary" onClick={onNewAudit}>Upload New Files</button>
        <button className="btn btn-secondary" onClick={onExit}>Exit</button>
      </div>
    </>
  )
}
