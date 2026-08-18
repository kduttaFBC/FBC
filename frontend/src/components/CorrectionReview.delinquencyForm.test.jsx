/**
 * Delinquency Information Request Form download — CLIENT-CONFIRMED (2026-08-18).
 *
 * The completion screen offers the download when, and only when, the server says
 * it generated one. PRESENTATION ONLY: the frontend reads four fields the
 * backend already returns and decides nothing.
 *
 * WHAT MATTERS HERE
 *   * `generated` is the gate, AND the bytes must be present. A payload
 *     asserting one without the other offers no download.
 *   * A blocked run shows a concise reason and NO download control -- a button
 *     that cannot produce a file is worse than its absence.
 *   * The control sits AFTER the mismatch report and BEFORE the Monthly Servicer
 *     Reports.
 *   * Rendering or clicking triggers no API call. The bytes arrived with the
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

const FORM_LABEL = 'Download Delinquency Information Request Form'
const MISMATCH_LABEL = 'Download Submitted Information vs. FBC Mismatch Report'
const MSR_LABEL = 'Download Monthly Servicer Report'
const SUMMARY_LABEL = 'Download Final Audited Portfolio Summary'
const CORRECTED_LABEL = 'Download Corrected File'

const B64_FORM = btoa('DELINQUENCYFORM')
const B64_MISMATCH = btoa('MISMATCH')
const B64_MSR = btoa('MSR')
const B64_SUMMARY = btoa('SUMMARY')
const B64_CORRECTED = btoa('CORRECTED')

const FILENAME = 'Delinquency_Information_Request_Form_July_2026.xlsx'

function rec(overrides = {}) {
  return {
    id: 'rec-1', servicer: 'Lendistry 2025', gp: '1000000001',
    field: 'Delinquency Status', standard: true,
    category: 'Standard — Delinquency', reported: '60 DAYS', calculated: '30 DAYS',
    requiresBalanceDecision: false, ...overrides,
  }
}

/** A complete, reporting-ready review response. */
function response(overrides = {}) {
  return {
    summary: { applied: 1, rejected: 0, returned: 0, not_applied: 0 },
    corrected_workbooks: [{
      servicer: 'Lendistry 2025', filename: 'Lendistry_CORRECTED.xlsx',
      file_b64: B64_CORRECTED, applied_count: 1, rejected_count: 0, returned_count: 0,
    }],
    monthly_servicer_reports: [{
      deal: 'SOUP2025_1', note_series: '2025 FBC-1', loan_count: 1436,
      servicers: ['Lendistry 2025'], filename: 'MSR_July_2026.xlsx', file_b64: B64_MSR,
    }],
    monthly_servicer_report_validation_issues: [],
    status: 'complete',
    deals: [{ deal: 'SOUP2025_1', status: 'complete',
              servicers_expected: ['Lendistry 2025'],
              servicers_reported: ['Lendistry 2025'], failures: [],
              report_available: true }],
    review_completeness: {
      total: 1, approved_applied: 1, approved_not_applied: 0, rejected: 0,
      returned: 0, pending: 0, unknown: 0, review_complete: true,
    },
    reporting_readiness: {
      reportable: 1, non_reportable: 0, reporting_ready: true,
      blocking_reasons: [], blocking_items: [],
    },
    portfolio_summary_final_b64: B64_SUMMARY,
    portfolio_summary_final_filename: 'Portfolio_Summary_Final_Audited_July_2026.xlsx',
    mismatch_report_b64: B64_MISMATCH,
    mismatch_report_filename: 'Submitted_vs_FBC_Mismatch_Report_July_2026.xlsx',
    mismatch_report_status: {
      generated: true, blocked_reason: null, mismatch_count: 3, provenance_gap_count: 0,
    },
    delinquency_request_form_b64: B64_FORM,
    delinquency_request_form_filename: FILENAME,
    delinquency_request_form_status: {
      generated: true, blocked_reason: null, loan_count: 104, lender_count: 8,
    },
    corrections_not_applied: [],
    unknown_ids: [],
    ...overrides,
  }
}

async function submitWith(data, recs = [rec()]) {
  axios.post.mockResolvedValueOnce({ data })
  render(<CorrectionReview recommendations={recs} sessionId="s1" />)
  fireEvent.change(screen.getByPlaceholderText(/name/i), { target: { value: 'K. Dutta' } })
  fireEvent.click(screen.getByRole('button', { name: /Approve All Recommendations/ }))
  fireEvent.click(screen.getByRole('button', { name: /^Record 1 decision$/ }))
  await waitFor(() => expect(screen.getByText(/^Recorded:/)).toBeTruthy())
}

let downloads = []

beforeEach(() => {
  axios.post.mockClear()
  downloads = []
  global.URL.createObjectURL = vi.fn(() => 'blob:stub')
  global.URL.revokeObjectURL = vi.fn()
  vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function () {
    downloads.push({ filename: this.download, href: this.href })
  })
})

describe('the download appears when the server generated a form', () => {
  it('renders the control with the exact approved label', async () => {
    await submitWith(response())
    expect(screen.getByRole('button', { name: FORM_LABEL })).toBeTruthy()
  })

  it('summarises the loans and lender worksheets', async () => {
    await submitWith(response())
    expect(screen.getByText(/104 loans across 8 lender worksheets/)).toBeTruthy()
  })

  it('states a clean month rather than showing a bare zero', async () => {
    await submitWith(response({
      delinquency_request_form_status: {
        generated: true, blocked_reason: null, loan_count: 0, lender_count: 8,
      },
    }))
    expect(screen.getByRole('button', { name: FORM_LABEL })).toBeTruthy()
    expect(screen.getByText(/No loans requiring additional information/)).toBeTruthy()
  })

  it('sits AFTER the mismatch report and BEFORE the Monthly Servicer Report', async () => {
    await submitWith(response())
    const order = screen.getAllByRole('button')
      .map((b) => b.textContent.trim())
      .filter((t) => [CORRECTED_LABEL, MISMATCH_LABEL, FORM_LABEL, MSR_LABEL,
                      SUMMARY_LABEL].includes(t))
    expect(order).toEqual([CORRECTED_LABEL, MISMATCH_LABEL, FORM_LABEL, MSR_LABEL,
                           SUMMARY_LABEL])
  })
})

describe('it downloads the returned bytes under the returned filename', () => {
  it('uses delinquency_request_form_filename verbatim', async () => {
    await submitWith(response())
    fireEvent.click(screen.getByRole('button', { name: FORM_LABEL }))
    expect(downloads).toHaveLength(1)
    expect(downloads[0].filename).toBe(FILENAME)
  })

  it('decodes delinquency_request_form_b64, not another artifact’s bytes', async () => {
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
      fireEvent.click(screen.getByRole('button', { name: FORM_LABEL }))
      expect(blobText).toBe('DELINQUENCYFORM')
    } finally {
      global.Blob = realBlob
    }
  })

  it('falls back to a sane filename when the server sends none', async () => {
    await submitWith(response({ delinquency_request_form_filename: null }))
    fireEvent.click(screen.getByRole('button', { name: FORM_LABEL }))
    expect(downloads[0].filename).toBe('Delinquency_Information_Request_Form.xlsx')
  })
})

describe('it is hidden when the server did not generate one', () => {
  it('renders no control when generated is false', async () => {
    await submitWith(response({
      delinquency_request_form_b64: null, delinquency_request_form_filename: null,
      delinquency_request_form_status: {
        generated: false,
        blocked_reason: 'Human Review is not yet complete: 2 recommendation(s) still pending.',
        loan_count: 0, lender_count: 0,
      },
    }))
    expect(screen.queryByRole('button', { name: FORM_LABEL })).toBeNull()
  })

  it('never renders a control from bytes alone when generated is false', async () => {
    await submitWith(response({
      delinquency_request_form_status: {
        generated: false, blocked_reason: 'Blocked.', loan_count: 0, lender_count: 0,
      },
    }))
    expect(screen.queryByRole('button', { name: FORM_LABEL })).toBeNull()
  })

  it('stays hidden when the bytes are missing even though generated is true', async () => {
    await submitWith(response({
      delinquency_request_form_b64: null,
      delinquency_request_form_status: {
        generated: true, blocked_reason: null, loan_count: 104, lender_count: 8,
      },
    }))
    expect(screen.queryByRole('button', { name: FORM_LABEL })).toBeNull()
  })

  it('renders nothing at all for a server that returned no status', async () => {
    const data = response()
    delete data.delinquency_request_form_status
    delete data.delinquency_request_form_b64
    delete data.delinquency_request_form_filename
    await submitWith(data)
    expect(screen.queryByRole('button', { name: FORM_LABEL })).toBeNull()
    expect(screen.queryByText(/Request Form not generated/)).toBeNull()
    expect(screen.getByRole('button', { name: MISMATCH_LABEL })).toBeTruthy()
    expect(screen.getByRole('button', { name: MSR_LABEL })).toBeTruthy()
  })
})

describe('a blocked run explains itself safely', () => {
  const BLOCKED = {
    delinquency_request_form_b64: null,
    delinquency_request_form_filename: null,
    delinquency_request_form_status: {
      generated: false,
      blocked_reason: 'The Delinquency Information Request Form was not generated ' +
        'because Lendistry 2025 could not be read or consolidated. An incomplete ' +
        'form would omit loans the lenders need to report on.',
      loan_count: 0, lender_count: 0,
    },
  }

  it('shows the server’s reason and no download control', async () => {
    await submitWith(response(BLOCKED))
    expect(screen.getByText(/Delinquency Information Request Form not generated/)).toBeTruthy()
    expect(screen.getByText(/could not be read or consolidated/)).toBeTruthy()
    expect(screen.queryByRole('button', { name: FORM_LABEL })).toBeNull()
  })

  it('exposes no internal diagnostic or lender-level failure structure', async () => {
    await submitWith(response(BLOCKED))
    const text = document.body.textContent
    for (const leak of ['delinquency_blocks', 'delinquency_failures',
                        'delinquency_servicers', 'error_type', 'error_detail',
                        'Traceback', 'RuntimeError', 'excluded_gps',
                        'delinquency_request_form_b64']) {
      expect(text).not.toContain(leak)
    }
  })

  it('says nothing when generation was blocked without a reason', async () => {
    await submitWith(response({
      delinquency_request_form_b64: null,
      delinquency_request_form_status: {
        generated: false, blocked_reason: null, loan_count: 0, lender_count: 0,
      },
    }))
    expect(screen.queryByText(/Request Form not generated/)).toBeNull()
  })
})

describe('every existing control is preserved', () => {
  it('keeps the corrected workbook, mismatch, MSR and Final Summary controls', async () => {
    await submitWith(response())
    for (const label of [CORRECTED_LABEL, MISMATCH_LABEL, MSR_LABEL, SUMMARY_LABEL]) {
      expect(screen.getByRole('button', { name: label })).toBeTruthy()
    }
  })

  it('leaves their filenames and bytes untouched', async () => {
    await submitWith(response())
    fireEvent.click(screen.getByRole('button', { name: CORRECTED_LABEL }))
    fireEvent.click(screen.getByRole('button', { name: MISMATCH_LABEL }))
    fireEvent.click(screen.getByRole('button', { name: MSR_LABEL }))
    fireEvent.click(screen.getByRole('button', { name: SUMMARY_LABEL }))
    expect(downloads.map((d) => d.filename)).toEqual([
      'Lendistry_CORRECTED.xlsx',
      'Submitted_vs_FBC_Mismatch_Report_July_2026.xlsx',
      'MSR_July_2026.xlsx',
      'Portfolio_Summary_Final_Audited_July_2026.xlsx',
    ])
  })

  it('still shows them when the form is blocked', async () => {
    await submitWith(response({
      delinquency_request_form_b64: null,
      delinquency_request_form_status: {
        generated: false, blocked_reason: 'Blocked.', loan_count: 0, lender_count: 0,
      },
    }))
    for (const label of [CORRECTED_LABEL, MISMATCH_LABEL, MSR_LABEL, SUMMARY_LABEL]) {
      expect(screen.getByRole('button', { name: label })).toBeTruthy()
    }
  })

  it('leaves Reviewer Name and both Approve All controls alone', async () => {
    render(<CorrectionReview recommendations={[rec()]} sessionId="s1" />)
    expect(screen.getByPlaceholderText(/name/i)).toBeTruthy()
    expect(screen.getByRole('button', { name: /Approve All Recommendations/ })).toBeTruthy()
    expect(screen.getByRole('button', { name: /Approve All Standard Corrections/ })).toBeTruthy()
  })

  it('appears in the complete-but-not-reporting-ready state, where MSR and Summary do not', async () => {
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
    expect(screen.getByRole('button', { name: FORM_LABEL })).toBeTruthy()
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
    fireEvent.click(screen.getByRole('button', { name: FORM_LABEL }))
    expect(axios.post).not.toHaveBeenCalled()
    expect(downloads).toHaveLength(1)
  })

  it('reaches the network by no other route either', async () => {
    const fetchSpy = vi.fn()
    const xhrSpy = vi.fn()
    global.fetch = fetchSpy
    vi.spyOn(XMLHttpRequest.prototype, 'open').mockImplementation(xhrSpy)
    await submitWith(response())
    fireEvent.click(screen.getByRole('button', { name: FORM_LABEL }))
    expect(fetchSpy).not.toHaveBeenCalled()
    expect(xhrSpy).not.toHaveBeenCalled()
    expect(downloads).toHaveLength(1)
  })
})
