import React, { useState } from 'react'
import axios from 'axios'
import Navbar from './components/Navbar.jsx'
import UploadScreen from './components/UploadScreen.jsx'
import ReadyScreen from './components/ReadyScreen.jsx'
import ResultsScreen from './components/ResultsScreen.jsx'
import Chatbot from './components/Chatbot.jsx'
import SessionEnded from './components/SessionEnded.jsx'

// The standard fields we look for in each file type. The parser scans a
// file's sheets/rows to find where these headers actually live, rather than
// assuming row 1 — real servicer files bury the header row a few rows down.

// Left side — SOUP Performance Dataset (25,000-record loan performance file)
export const SOUP_FIELDS = [
  'SBA GP Number', 'Servicer', 'Payment Month', 'Loan Status',
  'Delinquency Status', 'Beginning Actual Balance', 'Ending Actual Balance',
]

// Right side — servicer SRF files (Poppy Bank + 7 additional, same template)
export const SERVICER_FIELDS = [
  'SBA GP Number', 'Lender Loan Number', 'Loan Status',
  'Delinquency Status', 'Beginning Balance', 'Ending Balance', 'Payment Received',
]

// Second monthly source file, alongside SOUP — supporting servicing/deferral
// information the audit cross-references against SOUP's loan performance
// data (see file_readers.read_service_memo / checks._lookup_deferral on the
// backend). Workbook structure is fixed; only the loan-level data changes
// each month, same as SOUP and the servicer files.
export const MEMO_FIELDS = [
  'Report Date', 'Lender', 'SBA Loan ID', 'Lender Loan ID', 'Borrower Name',
  'Servicer Action', 'Deferral Start Date', 'Deferral End Date',
  'Regular Payment Resume Date',
]

// Right-side slot labels. First is Poppy Bank (the file we already have);
// the rest are placeholders until each servicer's file is confirmed.
// Both Lendistry and Millennium Bank submit two separate SBA note-series
// files (2024 FBC-1 and 2025 FBC-1 pools) that each run as their own
// independent reconciliation — NOT merged into one. "2024"/"2025" here
// refers to the note series/pool, not the calendar year of the activity
// data itself (both pools can contain the same reporting month's activity).
// Both note-series files for a servicer still need to match against that
// single servicer's name in the SOUP dataset, so SERVICER_SOUP_KEYS below
// maps each display label back to the actual SOUP servicer name.
export const SERVICER_SLOTS = [
  'Poppy Bank',
  'Millennium 2025',
  'Millennium 2024',
  'Mission Valley Bank',
  'Peoples Bank',
  'Lendistry 2025',
  'Lendistry 2024',
  'Community Bank and Trust',
  // Added 2026-08-04: 4 new banks, one file each (no note-series split
  // confirmed for any of them, unlike Millennium/Lendistry above). These
  // are the exact names as given — NOT yet verified against the actual
  // spelling used in SOUP's "Servicer" column. Community Bank and Trust
  // above needed a SERVICER_SOUP_KEYS mapping because SOUP spells it with
  // "&" not "and" — if any of these 4 have a similar mismatch, their audit
  // will silently find zero matching SOUP rows (every history-dependent
  // check skips every loan) until a SERVICER_SOUP_KEYS entry is added
  // below. Verify each of these against a real SOUP export before relying
  // on a real audit for these banks.
  'Shoreham',
  'Cogent',
  'FinWise',
  'X-Caliber',
]

// Maps a display label to the servicer name that actually appears in the
// SOUP dataset's "Servicer" column, for any label that doesn't match SOUP
// verbatim (e.g. note-series-specific labels all map back to the servicer's
// plain name). Labels not listed here are assumed to match SOUP directly.
export const SERVICER_SOUP_KEYS = {
  'Lendistry 2025': 'Lendistry',
  'Lendistry 2024': 'Lendistry',
  'Millennium 2025': 'Millennium Bank',
  'Millennium 2024': 'Millennium Bank',
  // SOUP's Servicer column spells this with "&", not "and" — without this
  // mapping, filter_soup_by_servicer finds zero matching SOUP rows for this
  // servicer, which silently makes every history-dependent check (BAL-01,
  // STAT-01, PAY-01...) skip every loan instead of actually checking it.
  'Community Bank and Trust': 'Community Bank & Trust',
}

// Neutral on-screen labels for each comparison-file slot, shown in the UI in
// place of the real servicer names. Purely cosmetic: SERVICER_SLOTS (and the
// soupKey/backend-matching derived from it above) is completely unchanged —
// this array only controls what's rendered as each upload card's title.
export const SERVICER_SLOT_LABELS = SERVICER_SLOTS.map((_, i) => `Comparison Dataset ${i + 1}`)

export default function App() {
  // screens: 'upload' | 'ready' | 'results' | 'ended'
  const [screen, setScreen] = useState('upload')

  // validated files (File objects) and their detected headers/rows
  const [soup, setSoup] = useState(null) // { file, headers, rowCount, sheetName }
  const [memo, setMemo] = useState(null) // { file, headers, rowCount, sheetName } — optional
  const [servicerFiles, setServicerFiles] = useState(Array(SERVICER_SLOTS.length).fill(null))

  const [audit, setAudit] = useState(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  const goToReady = (soupFile, memoFile, servicers) => {
    setSoup(soupFile)
    setMemo(memoFile)
    setServicerFiles(servicers)
    setScreen('ready')
  }

  const runAudit = async () => {
    setLoading(true)
    setError('')
    try {
      const form = new FormData()
      form.append('soup_file', soup.file)
      // Servicer Memo Summary is a separate source document — sent as its
      // own field, never merged into soup_file. It's optional: the backend
      // (checks.run_servicer_reconciliation) degrades gracefully to
      // pre-deferral-aware behavior when it's not provided.
      if (memo && memo.valid) {
        form.append('memo_file', memo.file)
      }
      servicerFiles.forEach((s) => {
        // The file's servicer identity comes from what was detected/confirmed
        // for THIS file (set in UploadScreen) — never from its slot position.
        // SERVICER_SLOTS[i] must not be consulted here; a file dropped into
        // any slot is tagged with whatever institution was actually confirmed
        // for it.
        if (s && s.valid && s.servicerLabel) {
          const label = s.servicerLabel
          const soupKey = SERVICER_SOUP_KEYS[label] || label
          // Filename convention: "{soupKey}||{displayLabel}__{originalFilename}"
          // soupKey is what filter_soup_by_servicer matches against SOUP's
          // Servicer column; displayLabel is what shows up in the report
          // tabs, summary, and chatbot. They're equal for most servicers —
          // only diverge for things like Lendistry's two yearly files.
          form.append('servicer_files', s.file, `${soupKey}||${label}__${s.file.name}`)
        }
      })

      const apiUrl = import.meta.env.VITE_API_URL || ''
      const { data } = await axios.post(`${apiUrl}/api/audit`, form, {
        headers: { 'Content-Type': 'multipart/form-data' },
      })
      setAudit(data)
      setScreen('results')
    } catch (err) {
      const msg =
        err.response?.data?.error ||
        'The audit could not be completed. Please check your files and try again. If the issue continues, review the server error details.'
      setError(msg)
    } finally {
      setLoading(false)
    }
  }

  const resetAll = () => {
    setSoup(null)
    setMemo(null)
    setServicerFiles(Array(SERVICER_SLOTS.length).fill(null))
    setAudit(null)
    setError('')
    setScreen('upload')
  }

  if (screen === 'ended') {
    return <SessionEnded onStartNew={resetAll} />
  }

  return (
    <>
      <Navbar />

      {screen === 'upload' && <UploadScreen onContinue={goToReady} />}

      {screen === 'ready' && (
        <ReadyScreen
          soup={soup}
          memo={memo}
          servicerFiles={servicerFiles}
          onBack={() => setScreen('upload')}
          onRun={runAudit}
          loading={loading}
          error={error}
        />
      )}

      {screen === 'results' && audit && (
        <div className="content">
          <ResultsScreen audit={audit} onNewAudit={resetAll} onExit={() => setScreen('ended')} />
          <Chatbot auditContext={audit.chat_context} summary={audit.summary} results={audit.results} sessionId={audit.session_id} />
        </div>
      )}
    </>
  )
}
