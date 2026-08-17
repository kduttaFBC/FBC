/**
 * Submitted Information vs. FBC Mismatch Report download — CLIENT-CONFIRMED
 * (2026-08-17).
 *
 * The completion screen offers the download when, and only when, the server
 * says it generated one. This is a PRESENTATION change: the frontend reads
 * three fields the backend already returns and decides nothing.
 *
 * WHAT MATTERS HERE
 *   * `generated` is the gate. Never inferred from the presence of bytes.
 *   * A blocked run shows a concise reason and NO download control -- a button
 *     that cannot produce a file is worse than its absence.
 *   * The download is FIRST among the population-level artifacts, before the
 *     Monthly Servicer Reports and the Final Audited Portfolio Summary.
 *   * Rendering or clicking it triggers no API call. The bytes arrived with the
 *     review response; there is no second endpoint and there must never be one.
 */
import React from 'react'
import { describe, it, expect, vi, beforeEach } from 'vitest'
import { render, screen, fireEvent, waitFor } from '@testing-library/react'

vi.mock('axios', () => ({
  default: { post: vi.fn(() => Promise.resolve({ data: {} })) },
}))
import axios from 'axios'

import CorrectionReview from './CorrectionReview.jsx'

const MISMATCH_LABEL = 'Download Submitted Information vs. FBC Mismatch Report'
const MSR_LABEL = 'Download Monthly Servicer Report'
const SUMMARY_LABEL = 'Download Final Audited Portfolio Summary'
const CORRECTED_LABEL = 'Download Corrected File'

// "MISMATCH" / "MSR" / "SUMMARY" as base64, so a decoded blob is traceable to
// the field it came from.
const B64_MISMATCH = btoa('MISMATCH')
const B64_MSR = btoa('MSR')
const B64_SUMMARY = btoa('SUMMARY')
const B64_CORRECTED = btoa('CORRECTED')

const FILENAME = 'Submitted_vs_FBC_Mismatch_Report_June_2026.xlsx'

function rec(overrides = {}) {
  return {
    id: 'rec-1',
    servicer: 'Lendistry 2025',
    gp: '1000000001',
    field: 'Delinquency Status',
    standard: true,
    category: 'Standard — Delinquency',
    reported: '60 DAYS',
    calculated: '30 DAYS',
    requiresBalanceDecision: false,
    ...overrides,
  }
}

/** A complete, reporting-ready review response. Overrides are shallow-merged so
 *  a test can change one field without restating the rest. */
function response(overrides = {}) {
  return {
    summary: { applied: 1, rejected: 0, returned: 0, not_applied: 0 },
    corrected_workbooks: [{
      servicer: 'Lendistry 2025', filename: 'Lendistry_CORRECTED.xlsx',
      file_b64: B64_CORRECTED, applied_count: 1, rejected_count: 0, returned_count: 0,
    }],
    monthly_servicer_reports: [{
      deal: 'SOUP2025_1', note_series: '2025 FBC-1', loan_count: 1436,
      servicers: ['Lendistry 2025'], filename: 'MSR_June_2026.xlsx', file_b64: B64_MSR,
    }],
    monthly_servicer_report_validation_issues: [],
    status: 'complete',
    deals: [{ deal: 'SOUP2025_1', status: 'complete', servicers_expected: ['Lendistry 2025'],
              servicers_reported: ['Lendistry 2025'], failures: [], report_available: true }],
    review_completeness: {
      total: 1, approved_applied: 1, approved_not_applied: 0, rejected: 0,
      returned: 0, pending: 0, unknown: 0, review_complete: true,
    },
    reporting_readiness: {
      reportable: 1, non_reportable: 0, reporting_ready: true,
      blocking_reasons: [], blocking_items: [],
    },
    portfolio_summary_final_b64: B64_SUMMARY,
    portfolio_summary_final_filename: 'Portfolio_Summary_Final_Audited_June_2026.xlsx',
    mismatch_report_b64: B64_MISMATCH,
    mismatch_report_filename: FILENAME,
    mismatch_report_status: {
      generated: true, blocked_reason: null, mismatch_count: 3, provenance_gap_count: 0,
    },
    corrections_not_applied: [],
    unknown_ids: [],
    ...overrides,
  }
}

/** Renders, records one decision, submits, and waits for the completion screen.
 *  The real submit path -- the only way `result` is ever set. */
async function submitWith(data, recs = [rec()]) {
  axios.post.mockResolvedValueOnce({ data })
  render(<CorrectionReview recommendations={recs} sessionId="s1" />)
  fireEvent.change(screen.getByPlaceholderText(/name/i), { target: { value: 'K. Dutta' } })
  fireEvent.click(screen.getByRole('button', { name: /Approve All Recommendations/ }))
  fireEvent.click(screen.getByRole('button', { name: /^Record 1 decision$/ }))
  await waitFor(() => expect(screen.getByText(/^Recorded:/)).toBeTruthy())
}

// Captured downloads: [{ filename, bytes }]
let downloads = []

beforeEach(() => {
  axios.post.mockClear()
  downloads = []
  global.URL.createObjectURL = vi.fn(() => 'blob:stub')
  global.URL.revokeObjectURL = vi.fn()
  // The component builds an <a>, sets download/href and clicks it. Intercept
  // the click so the filename it actually chose is observable.
  vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function () {
    downloads.push({ filename: this.download, href: this.href })
  })
})

describe('the download appears when the server generated a report', () => {
  it('renders the control with the exact approved label', async () => {
    await submitWith(response())
    expect(screen.getByRole('button', { name: MISMATCH_LABEL })).toBeTruthy()
  })

  it('summarises how many mismatches the workbook holds', async () => {
    await submitWith(response())
    expect(screen.getByText(/3 mismatches across SOUP 2024 FBC1 and SOUP 2025 FBC1/)).toBeTruthy()
  })

  it('states a clean month explicitly rather than showing a bare zero', async () => {
    await submitWith(response({
      mismatch_report_status: {
        generated: true, blocked_reason: null, mismatch_count: 0, provenance_gap_count: 0,
      },
    }))
    expect(screen.getByRole('button', { name: MISMATCH_LABEL })).toBeTruthy()
    expect(screen.getByText(/No mismatches identified/)).toBeTruthy()
  })

  it('sits BEFORE the Monthly Servicer Report and the Final Audited Summary', async () => {
    await submitWith(response())
    const buttons = screen.getAllByRole('button')
      .map((b) => b.textContent.trim())
      .filter((t) => [MISMATCH_LABEL, MSR_LABEL, SUMMARY_LABEL].includes(t))
    expect(buttons).toEqual([MISMATCH_LABEL, MSR_LABEL, SUMMARY_LABEL])
  })
})

describe('it downloads the returned bytes under the returned filename', () => {
  it('uses mismatch_report_filename verbatim', async () => {
    await submitWith(response())
    fireEvent.click(screen.getByRole('button', { name: MISMATCH_LABEL }))
    expect(downloads).toHaveLength(1)
    expect(downloads[0].filename).toBe(FILENAME)
  })

  it('decodes mismatch_report_b64, not another artifact’s bytes', async () => {
    let blobText = null
    const realBlob = global.Blob
    global.Blob = class extends realBlob {
      constructor(parts, opts) {
        super(parts, opts)
        blobText = new TextDecoder().decode(parts[0])
      }
    }
    try {
      await submitWith(response())
      fireEvent.click(screen.getByRole('button', { name: MISMATCH_LABEL }))
      expect(blobText).toBe('MISMATCH')
    } finally {
      global.Blob = realBlob
    }
  })

  it('falls back to a sane filename when the server sends none', async () => {
    await submitWith(response({ mismatch_report_filename: null }))
    fireEvent.click(screen.getByRole('button', { name: MISMATCH_LABEL }))
    expect(downloads[0].filename).toBe('Submitted_vs_FBC_Mismatch_Report.xlsx')
  })
})

describe('it is hidden when the server did not generate one', () => {
  it('renders no control when generated is false', async () => {
    await submitWith(response({
      mismatch_report_b64: null, mismatch_report_filename: null,
      mismatch_report_status: {
        generated: false,
        blocked_reason: 'Human Review is not yet complete: 2 recommendation(s) still pending.',
        mismatch_count: 0, provenance_gap_count: 0,
      },
    }))
    expect(screen.queryByRole('button', { name: MISMATCH_LABEL })).toBeNull()
  })

  it('never renders a control from bytes alone when generated is false', async () => {
    // Bytes present, generated false -- `generated` is the gate, and a stale or
    // partial payload must not be able to offer a download the server disowned.
    await submitWith(response({
      mismatch_report_status: {
        generated: false, blocked_reason: 'Blocked.', mismatch_count: 0,
        provenance_gap_count: 0,
      },
    }))
    expect(screen.queryByRole('button', { name: MISMATCH_LABEL })).toBeNull()
  })

  it('renders nothing at all for a server that returned no status', async () => {
    const data = response()
    delete data.mismatch_report_status
    delete data.mismatch_report_b64
    delete data.mismatch_report_filename
    await submitWith(data)
    expect(screen.queryByRole('button', { name: MISMATCH_LABEL })).toBeNull()
    expect(screen.queryByText(/Mismatch Report not generated/)).toBeNull()
    // ... and every other download is unaffected.
    expect(screen.getByRole('button', { name: MSR_LABEL })).toBeTruthy()
    expect(screen.getByRole('button', { name: SUMMARY_LABEL })).toBeTruthy()
  })
})

describe('a blocked run explains itself safely', () => {
  const BLOCKED = {
    mismatch_report_b64: null,
    mismatch_report_filename: null,
    mismatch_report_status: {
      generated: false,
      blocked_reason: 'The mismatch report was not generated because Lendistry 2025 ' +
        'could not be read or consolidated. An incomplete reconciliation would ' +
        'understate the differences.',
      mismatch_count: 0, provenance_gap_count: 0,
    },
  }

  it('shows the server’s reason and no download control', async () => {
    await submitWith(response(BLOCKED))
    expect(screen.getByText(/Mismatch Report not generated/)).toBeTruthy()
    expect(screen.getByText(/could not be read or consolidated/)).toBeTruthy()
    expect(screen.queryByRole('button', { name: MISMATCH_LABEL })).toBeNull()
  })

  it('shows the reason verbatim rather than rewording it', async () => {
    await submitWith(response(BLOCKED))
    expect(screen.getByText(
      new RegExp(BLOCKED.mismatch_report_status.blocked_reason.slice(0, 60)))).toBeTruthy()
  })

  it('exposes no internal diagnostic detail', async () => {
    await submitWith(response(BLOCKED))
    const text = document.body.textContent
    for (const leak of ['provenance_gap', 'provenance_gaps', 'unassigned',
                        'mismatch_report_b64', 'build_mismatch_rows', 'Traceback',
                        'RuntimeError']) {
      expect(text).not.toContain(leak)
    }
  })

  it('says nothing when generation was blocked without a reason', async () => {
    await submitWith(response({
      mismatch_report_b64: null,
      mismatch_report_status: {
        generated: false, blocked_reason: null, mismatch_count: 0, provenance_gap_count: 0,
      },
    }))
    expect(screen.queryByText(/Mismatch Report not generated/)).toBeNull()
  })
})

describe('every existing download is preserved', () => {
  it('keeps the corrected workbook, MSR and Final Audited Summary controls', async () => {
    await submitWith(response())
    expect(screen.getByRole('button', { name: CORRECTED_LABEL })).toBeTruthy()
    expect(screen.getByRole('button', { name: MSR_LABEL })).toBeTruthy()
    expect(screen.getByRole('button', { name: SUMMARY_LABEL })).toBeTruthy()
  })

  it('leaves their filenames and bytes untouched', async () => {
    await submitWith(response())
    fireEvent.click(screen.getByRole('button', { name: MSR_LABEL }))
    fireEvent.click(screen.getByRole('button', { name: SUMMARY_LABEL }))
    expect(downloads.map((d) => d.filename)).toEqual([
      'MSR_June_2026.xlsx', 'Portfolio_Summary_Final_Audited_June_2026.xlsx',
    ])
  })

  it('still shows them when the mismatch report is blocked', async () => {
    await submitWith(response({
      mismatch_report_b64: null,
      mismatch_report_status: {
        generated: false, blocked_reason: 'Blocked.', mismatch_count: 0,
        provenance_gap_count: 0,
      },
    }))
    expect(screen.getByRole('button', { name: MSR_LABEL })).toBeTruthy()
    expect(screen.getByRole('button', { name: SUMMARY_LABEL })).toBeTruthy()
  })

  it('appears in the complete-but-not-reporting-ready state, where the others do not', async () => {
    // The asymmetry the backend gate creates, shown end to end: one returned
    // recommendation withholds the MSR and the Final Audited Summary while the
    // mismatch report is still produced.
    await submitWith(response({
      monthly_servicer_reports: [],
      portfolio_summary_final_b64: null,
      portfolio_summary_final_filename: null,
      review_completeness: {
        total: 1, approved_applied: 0, approved_not_applied: 0, rejected: 0,
        returned: 1, pending: 0, unknown: 0, review_complete: true,
      },
      reporting_readiness: {
        reportable: 0, non_reportable: 1, reporting_ready: false,
        blocking_reasons: ['returned'], blocking_items: [],
      },
    }))
    expect(screen.getByRole('button', { name: MISMATCH_LABEL })).toBeTruthy()
    expect(screen.queryByRole('button', { name: MSR_LABEL })).toBeNull()
    expect(screen.queryByRole('button', { name: SUMMARY_LABEL })).toBeNull()
    expect(screen.getByText(/Final reporting is blocked/)).toBeTruthy()
  })
})

describe('no API request is triggered by the download', () => {
  it('rendering the completion screen makes exactly the one review call', async () => {
    await submitWith(response())
    expect(axios.post).toHaveBeenCalledTimes(1)
    expect(axios.post.mock.calls[0][0]).toMatch(/\/api\/corrections\/review$/)
  })

  it('clicking the download makes no further call', async () => {
    await submitWith(response())
    axios.post.mockClear()
    fireEvent.click(screen.getByRole('button', { name: MISMATCH_LABEL }))
    expect(axios.post).not.toHaveBeenCalled()
    expect(downloads).toHaveLength(1)
  })

  it('reaches the network by no other route either', async () => {
    // The bytes arrive with the review response; there is no second endpoint
    // and adding one would change a backend API contract this commit must not
    // touch. axios is already mocked, so fetch and XHR are what remain.
    const fetchSpy = vi.fn()
    const xhrSpy = vi.fn()
    global.fetch = fetchSpy
    vi.spyOn(XMLHttpRequest.prototype, 'open').mockImplementation(xhrSpy)
    await submitWith(response())
    fireEvent.click(screen.getByRole('button', { name: MISMATCH_LABEL }))
    expect(fetchSpy).not.toHaveBeenCalled()
    expect(xhrSpy).not.toHaveBeenCalled()
    expect(downloads).toHaveLength(1)
  })
})
