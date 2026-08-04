import React from 'react'
import { SERVICER_SLOTS, SERVICER_SLOT_LABELS } from '../App.jsx'

function FileRow({ label, state, institution }) {
  return (
    <div className="map-row" style={{ gridTemplateColumns: '1.4fr 1fr' }}>
      <div className="field-name">
        {label}
        {institution && <div className="hint" style={{ marginTop: 2 }}>{institution}</div>}
      </div>
      <div style={{ color: state?.valid ? 'var(--teal)' : 'var(--text-secondary)' }}>
        {state?.valid ? `✓ ${state.fileName} — ${state.rowCount} rows` : '— not provided —'}
      </div>
    </div>
  )
}

export default function ReadyScreen({ soup, memo, servicerFiles, onBack, onRun, loading, error }) {
  return (
    <div className="content">
      <div className="card">
        <h2 style={{ marginTop: 0 }}>Ready to run</h2>
        <p style={{ color: 'var(--text-secondary)', marginTop: 0 }}>
          Column mapping is on hold for this build — review the files below and run when ready.
        </p>

        <div className="mapper-section">
          <h3>Source Files</h3>
          <FileRow label="Primary Dataset" state={soup} />
          <FileRow
            label={`Supporting Dataset${memo?.reportingMonth ? ` (${memo.reportingMonth})` : ''}`}
            state={memo}
          />
        </div>

        <div className="mapper-section">
          <h3>Comparison Datasets</h3>
          {SERVICER_SLOTS.map((label, idx) => (
            <FileRow
              key={label}
              label={SERVICER_SLOT_LABELS[idx]}
              institution={servicerFiles[idx]?.valid ? servicerFiles[idx].servicerLabel : null}
              state={servicerFiles[idx]}
            />
          ))}
        </div>

        {error && (
          <div className="status-bar red" style={{ margin: '8px 0' }}>{error}</div>
        )}

        <div className="mapper-actions">
          <button className="link-btn" onClick={onBack} disabled={loading}>← Back</button>
          <button className="btn btn-primary" disabled={loading} onClick={onRun}>
            {loading ? (<><span className="spinner" />Running…</>) : 'Run Audit →'}
          </button>
        </div>
      </div>
    </div>
  )
}
