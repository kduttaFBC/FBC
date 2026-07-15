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

// Right-side slot labels. First is Poppy Bank (the file we already have);
// the rest are placeholders until each servicer's file is confirmed.
// Lendistry provides two separate yearly files (2024, 2025) that each run
// as their own independent reconciliation — NOT merged into one. Both
// still need to match against the single "Lendistry" servicer key in the
// SOUP dataset, so SERVICER_SOUP_KEYS below maps each display label back
// to the actual SOUP servicer name for lookup purposes.
export const SERVICER_SLOTS = [
  'Poppy Bank',
  'Millennium Bank',
  'Mission Valley Bank',
  'Peoples Bank',
  'Lendistry 2025',
  'Lendistry 2024',
  'Community Bank and Trust',
  'Servicer File 8',
]

// Maps a display label to the servicer name that actually appears in the
// SOUP dataset's "Servicer" column, for any label that doesn't match SOUP
// verbatim (e.g. year-specific Lendistry labels all map back to plain
// "Lendistry"). Labels not listed here are assumed to match SOUP directly.
export const SERVICER_SOUP_KEYS = {
  'Lendistry 2025': 'Lendistry',
  'Lendistry 2024': 'Lendistry',
}

export default function App() {
  // screens: 'upload' | 'ready' | 'results' | 'ended'
  const [screen, setScreen] = useState('upload')

  // validated files (File objects) and their detected headers/rows
  const [soup, setSoup] = useState(null) // { file, headers, rowCount, sheetName }
  const [servicerFiles, setServicerFiles] = useState(Array(SERVICER_SLOTS.length).fill(null))

  const [audit, setAudit] = useState(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  const goToReady = (soupFile, servicers) => {
    setSoup(soupFile)
    setServicerFiles(servicers)
    setScreen('ready')
  }

  const runAudit = async () => {
    setLoading(true)
    setError('')
    try {
      const form = new FormData()
      form.append('soup_file', soup.file)
      servicerFiles.forEach((s, i) => {
        if (s && s.valid) {
          const label = SERVICER_SLOTS[i]
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
        'The audit backend has not been updated for the new files yet. Upload flow and validation are ready for UAT — audit logic is next.'
      setError(msg)
    } finally {
      setLoading(false)
    }
  }

  const resetAll = () => {
    setSoup(null)
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
