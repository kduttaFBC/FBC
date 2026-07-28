import React, { useState, useRef } from 'react'
import * as XLSX from 'xlsx'
import { SOUP_FIELDS, SERVICER_FIELDS, SERVICER_SLOTS, MEMO_FIELDS } from '../App.jsx'

const MAX_HEADER_SCAN_ROWS = 15

const norm = (s) => String(s).toLowerCase().replace(/[\s_\-]+/g, '')

function countMatches(requiredFields, headerRow) {
  const headers = (headerRow || []).map((h) => norm(h))
  let matches = 0
  requiredFields.forEach((f) => {
    if (headers.includes(norm(f))) matches += 1
  })
  return matches
}

// Reads just the first N rows of a sheet cheaply using the sheet's declared
// range, instead of converting the whole sheet (which can be 25,000+ rows)
// to JSON just to look at the header. Keeps large files from freezing the tab.
function readTopRows(sheet, maxRows) {
  if (!sheet['!ref']) return []
  const fullRange = XLSX.utils.decode_range(sheet['!ref'])
  const limitedRange = {
    s: { r: fullRange.s.r, c: fullRange.s.c },
    e: { r: Math.min(fullRange.e.r, fullRange.s.r + maxRows - 1), c: fullRange.e.c },
  }
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    range: limitedRange,
  })
}

// Scans every sheet of a workbook for the row that best matches the required
// fields (real servicer files bury headers a few rows down, under titles and
// summary rows), rather than assuming headers sit on row 1. Only reads the
// first MAX_HEADER_SCAN_ROWS of each sheet — cheap even for large files.
function findBestHeaderRow(workbook, requiredFields) {
  let best = { matches: -1, sheetName: null, rowIndex: -1, headerRow: null, totalRows: 0 }

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet['!ref']) return
    const fullRange = XLSX.utils.decode_range(sheet['!ref'])
    const totalRows = fullRange.e.r - fullRange.s.r + 1
    const rows = readTopRows(sheet, MAX_HEADER_SCAN_ROWS)
    rows.forEach((row, i) => {
      const matches = countMatches(requiredFields, row)
      if (matches > best.matches) {
        best = { matches, sheetName, rowIndex: i, headerRow: row, totalRows }
      }
    })
  })

  return best
}

// Counts rows that actually have data, checking only the first column
// cell-by-cell (cheap object lookups, no JSON materialization) rather than
// trusting the sheet's declared '!ref' range — Excel often pads that range
// far beyond the real data, which made large files report bogus row counts.
// Stops early after a long run of empty rows so a padded range (e.g. Excel
// reporting 1,048,576 rows) doesn't force scanning the whole thing.
function countActualDataRows(sheet, startRow, endRow, col) {
  const EARLY_EXIT_EMPTY_STREAK = 500
  let count = 0
  let consecutiveEmpty = 0
  for (let r = startRow; r <= endRow; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: col })
    const cell = sheet[addr]
    const hasValue = cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== ''
    if (hasValue) {
      count++
      consecutiveEmpty = 0
    } else {
      consecutiveEmpty++
      if (count > 0 && consecutiveEmpty >= EARLY_EXIT_EMPTY_STREAK) break
    }
  }
  return count
}

function parseWorkbookForFields(file, requiredFields, monthColumn) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      // yield to the browser once before the heavy parse so the "reading…"
      // state actually paints, then do the parse itself off the click handler
      setTimeout(() => {
        try {
          const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true })
          const best = findBestHeaderRow(wb, requiredFields)

          // require at least half the fields to show up somewhere, or we've
          // likely picked the wrong sheet/row entirely
          if (best.matches < Math.ceil(requiredFields.length / 2)) {
            resolve({ headers: [], rowCount: 0, sheetName: null, reportingMonth: null })
            return
          }

          const headers = (best.headerRow || []).map((h) => String(h).trim()).filter((h) => h !== '')
          const sheet = wb.Sheets[best.sheetName]
          const fullRange = XLSX.utils.decode_range(sheet['!ref'])
          // Check for data starting at the column where the header row's
          // first real (non-blank) cell actually is, not blindly the
          // sheet's absolute leftmost column — some files (e.g. the
          // Servicer Memo Summary) leave column A entirely blank with real
          // data starting in column B onward, which would otherwise make
          // every row look empty and wrongly report 0 rows found.
          const firstDataColOffset = (best.headerRow || []).findIndex(
            (h) => h !== undefined && h !== null && String(h).trim() !== ''
          )
          const rowCount = countActualDataRows(
            sheet,
            fullRange.s.r + best.rowIndex + 1,
            fullRange.e.r,
            fullRange.s.c + Math.max(0, firstDataColOffset),
          )
          const reportingMonth = monthColumn
            ? extractLatestMonth(sheet, best.rowIndex, headers, monthColumn)
            : null

          resolve({ headers, rowCount, sheetName: best.sheetName, reportingMonth })
        } catch (err) {
          reject(err)
        }
      }, 0)
    }
    reader.onerror = () => reject(new Error('Could not read the file.'))
    reader.readAsArrayBuffer(file)
  })
}

function findMissing(requiredFields, headers) {
  const lower = headers.map((h) => h.toLowerCase().trim())
  return requiredFields.filter((f) => !lower.includes(f.toLowerCase()))
}

// Scans a single column across ALL of a sheet's data rows (not just the
// header-scan window) for the latest parseable date, used to surface a
// "reporting month" for source files like the Servicer Memo Summary that
// carry one row per entry rather than one value per file. Never hardcodes a
// month — it's derived fresh from whatever file was actually uploaded, so a
// new month's file just produces a new answer automatically. Returns a
// formatted "Month YYYY" string, or null if no parseable date is found
// (callers should show "Not available" rather than guessing).
function extractLatestMonth(sheet, headerRowIndex, headers, columnName) {
  const colIdx = headers.findIndex((h) => norm(h) === norm(columnName))
  if (colIdx === -1 || !sheet['!ref']) return null

  const fullRange = XLSX.utils.decode_range(sheet['!ref'])
  const range = {
    s: { r: fullRange.s.r + headerRowIndex + 1, c: fullRange.s.c + colIdx },
    e: { r: fullRange.e.r, c: fullRange.s.c + colIdx },
  }
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', range, raw: false })

  let latest = null
  rows.forEach((row) => {
    const raw = row[0]
    if (raw === '' || raw == null) return
    const parsed = raw instanceof Date ? raw : new Date(raw)
    if (!isNaN(parsed) && (!latest || parsed > latest)) latest = parsed
  })
  if (!latest) return null
  return latest.toLocaleDateString('en-US', { month: 'long', year: 'numeric' })
}

function UploadCard({ title, subtitle, roleLabel, requiredFields, state, onFile, compact, reading, showValidationLabel }) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef(null)

  const handleFiles = (files) => {
    if (files && files.length) onFile(files[0])
  }

  return (
    <div className={`card upload-card ${compact ? 'upload-card-compact' : ''}`}>
      {roleLabel && <div className="source-role">{roleLabel}</div>}
      <h3>{title}</h3>
      {subtitle && <div className="hint">{subtitle}</div>}

      {!compact && (
        <div className="required-cols">
          <strong>Expected columns</strong>
          {requiredFields.join(', ')}
        </div>
      )}

      {reading && (
        <div className="file-state" style={{ background: 'var(--blue-light)', border: '1px solid #BFD8F2' }}>
          <div className="fname"><span className="spinner" style={{ borderTopColor: 'var(--blue)', borderColor: 'rgba(24,95,165,0.3)' }} />Reading file…</div>
        </div>
      )}

      {!state && !reading && (
        <div
          className={`dropzone ${compact ? 'dropzone-compact' : ''} ${dragging ? 'dragging' : ''}`}
          onClick={() => inputRef.current?.click()}
          onDragOver={(e) => { e.preventDefault(); setDragging(true) }}
          onDragLeave={() => setDragging(false)}
          onDrop={(e) => {
            e.preventDefault()
            setDragging(false)
            handleFiles(e.dataTransfer.files)
          }}
        >
          <div className="big">⬆</div>
          <div>{compact ? 'Click or drop file' : 'Drag & drop your file here, or click to browse'}</div>
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls"
            style={{ display: 'none' }}
            onChange={(e) => handleFiles(e.target.files)}
          />
        </div>
      )}

      {state && state.valid && (
        <div className="file-state success">
          <div className="fname">✓ {state.fileName}</div>
          <div className="meta">
            {state.rowCount} data row(s) detected{state.sheetName ? ` in "${state.sheetName}"` : ''} — looks good.
          </div>
          {state.reportingMonth && (
            <div className="meta">Reporting month: <strong>{state.reportingMonth}</strong></div>
          )}
          {showValidationLabel && <div className="meta">Validation status: <strong>Passed</strong></div>}
          <button className="link-btn" onClick={() => onFile(null)}>Replace file</button>
        </div>
      )}

      {state && !state.valid && (
        <div className="file-state error">
          <div className="fname">✗ {state.fileName || 'Upload failed'}</div>
          <div className="meta">{state.message}</div>
          {showValidationLabel && <div className="meta">Validation status: <strong>Failed</strong></div>}
          <button className="link-btn" onClick={() => onFile(null)}>Replace file</button>
        </div>
      )}
    </div>
  )
}

export default function UploadScreen({ onContinue }) {
  const [soup, setSoup] = useState(null)
  const [soupReading, setSoupReading] = useState(false)
  const [memo, setMemo] = useState(null)
  const [memoReading, setMemoReading] = useState(false)
  const [servicerFiles, setServicerFiles] = useState(Array(SERVICER_SLOTS.length).fill(null))
  const [servicerReading, setServicerReading] = useState(Array(SERVICER_SLOTS.length).fill(false))

  const validateAndSet = (requiredFields, file, applyState, setReading, monthColumn) => async (fileArg) => {
    if (!fileArg) { applyState(null); return }

    const name = fileArg.name.toLowerCase()
    if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
      applyState({
        valid: false,
        fileName: fileArg.name,
        message: 'Wrong file type. Please upload an Excel file (.xlsx or .xls).',
      })
      return
    }

    setReading?.(true)
    try {
      const { headers, rowCount, sheetName, reportingMonth } = await parseWorkbookForFields(fileArg, requiredFields, monthColumn)
      if (!headers.length || rowCount < 1) {
        applyState({
          valid: false,
          fileName: fileArg.name,
          message: 'Could not find the expected columns in this file. Please check it is the correct export.',
        })
        return
      }
      const missing = findMissing(requiredFields, headers)
      if (missing.length) {
        applyState({
          valid: false,
          fileName: fileArg.name,
          message: `Missing expected columns: ${missing.join(', ')}`,
        })
        return
      }
      applyState({ valid: true, fileName: fileArg.name, headers, rowCount, sheetName, reportingMonth, file: fileArg })
    } catch (err) {
      applyState({
        valid: false,
        fileName: fileArg.name,
        message: 'We could not read this file. Please re-export and try again.',
      })
    } finally {
      setReading?.(false)
    }
  }

  const handleSoupFile = validateAndSet(SOUP_FIELDS, null, setSoup, setSoupReading)

  const handleMemoFile = validateAndSet(MEMO_FIELDS, null, setMemo, setMemoReading, 'Report Date')

  const handleServicerFile = (idx) => validateAndSet(SERVICER_FIELDS, null, (state) => {
    setServicerFiles((cur) => {
      const next = [...cur]
      next[idx] = state
      return next
    })
  }, (isReading) => {
    setServicerReading((cur) => {
      const next = [...cur]
      next[idx] = isReading
      return next
    })
  })

  const servicerValidCount = servicerFiles.filter((s) => s?.valid).length
  const servicerErrorCount = servicerFiles.filter((s) => s && !s.valid).length
  const soupValid = soup?.valid
  const soupError = soup && !soup.valid
  const memoError = memo && !memo.valid

  // Servicer Memo Summary is optional (the backend degrades gracefully
  // without it), so it never blocks continuing on its own — but if one was
  // uploaded and failed validation, that's surfaced and must be resolved
  // rather than silently ignored (per "do not silently process the file").
  const canContinue = soupValid && servicerFiles.some((s) => s?.valid) && !memoError

  let status = { cls: 'gray', text: 'Upload the SOUP dataset and at least one servicer file to continue' }
  if (soupError || servicerErrorCount > 0 || memoError) {
    status = { cls: 'red', text: 'Fix the file errors above before continuing' }
  } else if (canContinue) {
    const memoNote = memo?.valid ? ', Servicer Memo Summary validated' : ''
    status = {
      cls: 'green',
      text: `Ready — SOUP dataset validated${memoNote}, ${servicerValidCount} of ${SERVICER_SLOTS.length} servicer file(s) uploaded`,
    }
  } else if (soupValid && servicerValidCount === 0) {
    status = { cls: 'amber', text: 'Waiting for at least one servicer file' }
  } else if (servicerValidCount > 0 && !soupValid) {
    status = { cls: 'amber', text: 'Waiting for the SOUP Performance Dataset' }
  }

  const handleContinue = () => {
    if (!canContinue) return
    onContinue(
      { valid: true, fileName: soup.fileName, file: soup.file, headers: soup.headers, rowCount: soup.rowCount, sheetName: soup.sheetName },
      memo?.valid
        ? { valid: true, fileName: memo.fileName, file: memo.file, headers: memo.headers, rowCount: memo.rowCount, sheetName: memo.sheetName, reportingMonth: memo.reportingMonth }
        : null,
      servicerFiles,
    )
  }

  return (
    <>
      <div className="hero">
        <h1>Data audit &amp; compliance engine</h1>
        <p>Upload your files and get a full audit report in seconds</p>
      </div>

      <div className="content">
        {/* Source files: SOUP is the primary monthly performance data;
            Servicer Memo Summary is a supporting source the audit engine
            cross-references it against (currently for deferral info). Grid
            is auto-fit so a future third source file can simply be added
            here without any layout rework. */}
        <div className="source-files-section">
          <div className="source-files-row">
            <UploadCard
              title="SOUP Master File"
              roleLabel="Primary source"
              subtitle="Primary monthly loan performance data (~25,000 records)"
              requiredFields={SOUP_FIELDS}
              state={soup}
              onFile={handleSoupFile}
              reading={soupReading}
            />
            <UploadCard
              title="Servicer Memo Summary"
              roleLabel="Supporting source"
              subtitle="Supporting monthly servicing and deferral information"
              requiredFields={MEMO_FIELDS}
              state={memo}
              onFile={handleMemoFile}
              reading={memoReading}
              showValidationLabel
            />
          </div>
        </div>

        <div className="servicer-panel-header">
          <h3>Servicer Files</h3>
          <div className="hint">Poppy Bank + up to 7 additional servicer SRF files (same template)</div>
        </div>
        <div className="servicer-grid">
          {SERVICER_SLOTS.map((label, idx) => (
            <UploadCard
              key={label}
              title={label}
              requiredFields={SERVICER_FIELDS}
              state={servicerFiles[idx]}
              onFile={handleServicerFile(idx)}
              compact
              reading={servicerReading[idx]}
            />
          ))}
        </div>

        <div className={`status-bar ${status.cls}`}>{status.text}</div>

        <div className="btn-row">
          <button className="btn btn-primary" disabled={!canContinue} onClick={handleContinue}>
            Continue →
          </button>
        </div>
      </div>
    </>
  )
}
