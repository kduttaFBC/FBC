import React, { useState, useMemo } from 'react'
import { CUSTOMER_FIELDS, BANK_FIELDS } from '../App.jsx'

const norm = (s) => String(s).toLowerCase().replace(/[\s_\-]+/g, '')

// Best-guess match for a required field within a file's headers.
function autoMatch(field, headers) {
  const nf = norm(field)
  // 1. exact (case/space-insensitive)
  let hit = headers.find((h) => norm(h) === nf)
  if (hit) return hit
  // 2. header contains the field or vice versa
  hit = headers.find((h) => norm(h).includes(nf) || nf.includes(norm(h)))
  if (hit) return hit
  // 3. token overlap — every word of the field appears in the header
  const tokens = field.toLowerCase().split(/\s+/)
  hit = headers.find((h) => {
    const hl = h.toLowerCase()
    return tokens.every((t) => hl.includes(t))
  })
  return hit || ''
}

function buildInitial(fields, headers) {
  const map = {}
  fields.forEach((f) => { map[f] = autoMatch(f, headers) })
  return map
}

function MappingSection({ title, fields, headers, mapping, onChange }) {
  return (
    <div className="mapper-section">
      <h3>{title}</h3>
      {fields.map((field) => {
        const value = mapping[field] || ''
        return (
          <div className="map-row" key={field}>
            <div className="field-name">{field}</div>
            <div className="arrow">→</div>
            <select
              className={value ? '' : 'unmatched'}
              value={value}
              onChange={(e) => onChange(field, e.target.value)}
            >
              <option value="">— Select column —</option>
              {headers.map((h) => (
                <option key={h} value={h}>{h}</option>
              ))}
            </select>
          </div>
        )
      })}
    </div>
  )
}

export default function ColumnMapper({ customer, bank, onBack, onConfirm, loading, error }) {
  const [custMap, setCustMap] = useState(() => buildInitial(CUSTOMER_FIELDS, customer.headers))
  const [bankMap, setBankMap] = useState(() => buildInitial(BANK_FIELDS, bank.headers))

  const allMatched = useMemo(() => {
    const custOk = CUSTOMER_FIELDS.every((f) => custMap[f])
    const bankOk = BANK_FIELDS.every((f) => bankMap[f])
    return custOk && bankOk
  }, [custMap, bankMap])

  const handleConfirm = () => {
    if (!allMatched || loading) return
    onConfirm({ customer: custMap, bank: bankMap })
  }

  return (
    <div className="content">
      <div className="card">
        <h2 style={{ marginTop: 0 }}>Confirm your column mapping</h2>
        <p style={{ color: 'var(--text-secondary)', marginTop: 0 }}>
          We&apos;ve matched your file&apos;s columns to our required fields. Review and
          correct any mismatches.
        </p>

        <MappingSection
          title="Customer Master"
          fields={CUSTOMER_FIELDS}
          headers={customer.headers}
          mapping={custMap}
          onChange={(f, v) => setCustMap((m) => ({ ...m, [f]: v }))}
        />

        <MappingSection
          title="Banking Transactions"
          fields={BANK_FIELDS}
          headers={bank.headers}
          mapping={bankMap}
          onChange={(f, v) => setBankMap((m) => ({ ...m, [f]: v }))}
        />

        {error && (
          <div className="status-bar red" style={{ margin: '8px 0' }}>{error}</div>
        )}

        <div className="mapper-actions">
          <button className="link-btn" onClick={onBack} disabled={loading}>← Back</button>
          <button
            className="btn btn-primary"
            disabled={!allMatched || loading}
            onClick={handleConfirm}
          >
            {loading ? (<><span className="spinner" />Running audit…</>) : 'Confirm & Run Audit →'}
          </button>
        </div>
      </div>
    </div>
  )
}
