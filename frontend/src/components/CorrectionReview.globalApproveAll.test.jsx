/**
 * Global Approve All — CLIENT-CONFIRMED (2026-08-17).
 *
 * The bulk pre-fill may sweep Standard, Data Domain, Terminal, Reversal and
 * Loan Status. It must never touch the required Ending Balance companion
 * decision, never touch Other — Flagged, never overwrite a decision already
 * made, and never submit.
 *
 * THE CENTRAL CASE, asserted end to end below: a LIQUIDATION -> ACTIVE Loan
 * Status recommendation with a required balance companion. The bulk action
 * approves the Loan Status using the engine's own value and leaves the balance
 * completely unresolved -- no Beginning Balance, no 0.00, nothing inferred --
 * and the review cannot complete until the reviewer chooses a balance by hand.
 */
import React from 'react'
import { describe, it, expect, vi, beforeEach } from 'vitest'
import { render, screen, fireEvent, within } from '@testing-library/react'

vi.mock('axios', () => ({
  default: { post: vi.fn(() => Promise.resolve({ data: {} })) },
}))
import axios from 'axios'

import CorrectionReview from './CorrectionReview.jsx'

const STANDARD = 'Standard — Delinquency'
const LOAN_STATUS = 'Loan Status'
const REVERSAL = 'Reversal / Transaction'
const DATA_DOMAIN = 'Data Domain / Exception'
const TERMINAL = 'Terminal Status Conflict'
const OTHER = 'Other — Flagged'

let seq = 0
function rec(overrides = {}) {
  seq += 1
  return {
    id: `id-${seq}`,
    servicer: 'Lendistry 2025',
    gp: `100000000${seq}`,
    field: 'Delinquency Status',
    standard: true,
    category: STANDARD,
    reported: '60 DAYS',
    calculated: '30 DAYS',
    requiresBalanceDecision: false,
    ...overrides,
  }
}

// The LIQUIDATION -> ACTIVE pair: a Loan Status recommendation and the required
// balance companion the server emits alongside it, on the SAME loan.
function liquidationReversalPair(gp = '5219499110') {
  const status = rec({
    gp,
    field: 'Loan Status',
    standard: false,
    category: LOAN_STATUS,
    reported: 'LIQUIDATION',
    calculated: 'ACTIVE',
  })
  const balance = rec({
    gp,
    field: 'Ending Balance',
    standard: false,
    category: LOAN_STATUS,
    reported: 0,
    calculated: null,
    requiresBalanceDecision: true,
    beginningBalance: 145729.87,
    priorEndingBalance: 145729.87,
    balanceChoices: [
      { action: 'restore', value: 145729.87, label: 'Restore the Beginning Balance' },
      { action: 'keep', value: 0, label: 'Keep 0.00' },
    ],
  })
  return { status, balance }
}

const setup = (recs) => render(<CorrectionReview recommendations={recs} sessionId="s1" />)
const globalButton = () => screen.getByRole('button', { name: /Approve All Recommendations/ })
const standardButton = () =>
  screen.getByRole('button', { name: /Approve All Standard Corrections/ })

// A recommendation's decision <select>, found by its row. Each rec renders one.
function decisionSelectFor(recId) {
  const row = document.getElementById(`rec-${recId}`)
  expect(row).toBeTruthy()
  return within(row).getAllByRole('combobox')[0]
}

beforeEach(() => {
  seq = 0
  axios.post.mockClear()
})

describe('Global Approve All — scope', () => {
  it('pre-fills every eligible category', () => {
    const recs = [
      rec(),
      rec({ standard: false, category: DATA_DOMAIN }),
      rec({ standard: false, category: TERMINAL }),
      rec({ standard: false, category: REVERSAL }),
    ]
    setup(recs)
    expect(globalButton()).toHaveTextContent('(4)')
    fireEvent.click(globalButton())
    recs.forEach((r) => expect(decisionSelectFor(r.id)).toHaveValue('approve'))
  })

  it('never pre-fills Other — Flagged', () => {
    const flagged = rec({ standard: false, category: OTHER })
    setup([rec(), flagged])
    fireEvent.click(globalButton())
    expect(decisionSelectFor(flagged.id)).toHaveValue('')
  })

  it('never pre-fills a recommendation with no engine value', () => {
    // A mandatory lender-reported LIQUIDATION carries no proposable status.
    const noValue = rec({
      standard: false, category: LOAN_STATUS, field: 'Loan Status', calculated: null,
    })
    setup([noValue])
    expect(globalButton()).toBeDisabled()
    fireEvent.click(globalButton())
    expect(decisionSelectFor(noValue.id)).toHaveValue('')
  })

  it('pre-fills a Loan Status recommendation with the engine value', () => {
    const status = rec({
      standard: false, category: LOAN_STATUS, field: 'Loan Status',
      reported: 'LIQUIDATION', calculated: 'ACTIVE',
    })
    setup([status])
    fireEvent.click(globalButton())
    expect(decisionSelectFor(status.id)).toHaveValue('approve')
  })
})

describe('Global Approve All — the LIQUIDATION -> ACTIVE balance companion', () => {
  it('approves the Loan Status and leaves the balance completely undecided', () => {
    const { status, balance } = liquidationReversalPair()
    setup([status, balance])

    fireEvent.click(globalButton())

    // Loan Status: approved, from the engine's own recommendation.
    expect(decisionSelectFor(status.id)).toHaveValue('approve')
    // Ending Balance: untouched.
    expect(decisionSelectFor(balance.id)).toHaveValue('')
  })

  it('pre-selects neither the Beginning Balance nor 0.00', () => {
    const { status, balance } = liquidationReversalPair()
    setup([status, balance])
    fireEvent.click(globalButton())

    // Every balance radio must remain unchecked. Nothing inferred, nothing
    // invented, and 0.00 is not treated as a safe default.
    const row = document.getElementById(`rec-${balance.id}`)
    const radios = within(row).queryAllByRole('radio')
    expect(radios.length).toBeGreaterThan(0)
    radios.forEach((radio) => expect(radio).not.toBeChecked())
  })

  it('counts the balance item as still requiring a decision', () => {
    const { status, balance } = liquidationReversalPair()
    setup([status, balance])
    fireEvent.click(globalButton())

    const panel = screen.getByTestId('still-requires-decision')
    expect(panel).toHaveTextContent('Still Requires Your Decision')
    expect(panel).toHaveTextContent(/Required balance decision/)
    expect(panel).toHaveTextContent(balance.gp)
  })

  it('blocks submission until the reviewer chooses a balance by hand', () => {
    const { status, balance } = liquidationReversalPair()
    setup([status, balance])
    fireEvent.click(globalButton())

    // The balance is undecided, so the review is incomplete.
    const row = document.getElementById(`rec-${balance.id}`)
    const radios = within(row).queryAllByRole('radio')
    radios.forEach((radio) => expect(radio).not.toBeChecked())
    expect(screen.getByTestId('still-requires-decision')).toBeInTheDocument()
  })

  it('completes normally once a permitted balance option is chosen', () => {
    const { status, balance } = liquidationReversalPair()
    setup([status, balance])
    fireEvent.click(globalButton())

    // The reviewer chooses one of the two server-supplied options by hand.
    const row = document.getElementById(`rec-${balance.id}`)
    fireEvent.click(within(row).getAllByRole('radio')[0])
    fireEvent.change(decisionSelectFor(balance.id), { target: { value: 'approve' } })

    expect(decisionSelectFor(balance.id)).toHaveValue('approve')
    expect(decisionSelectFor(status.id)).toHaveValue('approve')
    expect(screen.queryByTestId('still-requires-decision')).toBeNull()
  })

  it('treats an explicit 0.00 balance choice as a real decision', () => {
    const { status, balance } = liquidationReversalPair()
    setup([status, balance])
    const row = document.getElementById(`rec-${balance.id}`)
    // The "Keep 0.00" option is the second choice.
    fireEvent.click(within(row).getAllByRole('radio')[1])
    fireEvent.change(decisionSelectFor(balance.id), { target: { value: 'approve' } })
    expect(decisionSelectFor(balance.id)).toHaveValue('approve')
    expect(screen.queryByTestId('still-requires-decision')).toBeNull()
  })
})

describe('Global Approve All — safety properties', () => {
  it('never overwrites an existing Loan Status decision', () => {
    const status = rec({
      standard: false, category: LOAN_STATUS, field: 'Loan Status',
      reported: 'LIQUIDATION', calculated: 'ACTIVE',
    })
    setup([status, rec()])

    // The reviewer rejects the Loan Status first.
    fireEvent.change(decisionSelectFor(status.id), { target: { value: 'reject' } })
    expect(decisionSelectFor(status.id)).toHaveValue('reject')

    fireEvent.click(globalButton())
    expect(decisionSelectFor(status.id)).toHaveValue('reject')
  })

  it('never overwrites any existing decision', () => {
    const a = rec()
    const b = rec()
    setup([a, b])
    fireEvent.change(decisionSelectFor(a.id), { target: { value: 'return' } })
    fireEvent.click(globalButton())
    expect(decisionSelectFor(a.id)).toHaveValue('return')
    expect(decisionSelectFor(b.id)).toHaveValue('approve')
  })

  it('submits nothing', () => {
    setup([rec(), rec({ standard: false, category: TERMINAL })])
    fireEvent.click(globalButton())
    expect(axios.post).not.toHaveBeenCalled()
  })

  it('preserves the companion hold', () => {
    // A Standard item on a loan whose balance decision is undecided is held.
    const { status, balance } = liquidationReversalPair('7777777777')
    const sibling = rec({ gp: '7777777777' })
    setup([status, balance, sibling])
    fireEvent.click(globalButton())

    expect(decisionSelectFor(status.id)).toHaveValue('approve')
    expect(decisionSelectFor(balance.id)).toHaveValue('')
    expect(decisionSelectFor(sibling.id)).toHaveValue('')
    expect(screen.getByTestId('still-requires-decision')).toHaveTextContent(/Held/)
  })

  it('leaves the Standard action unchanged', () => {
    const standard = rec()
    const terminal = rec({ standard: false, category: TERMINAL })
    setup([standard, terminal])

    fireEvent.click(standardButton())
    expect(decisionSelectFor(standard.id)).toHaveValue('approve')
    // The Standard action must not reach a Review Required category.
    expect(decisionSelectFor(terminal.id)).toHaveValue('')
  })

  it('shows a reason for every held item', () => {
    const { status, balance } = liquidationReversalPair()
    const flagged = rec({ standard: false, category: OTHER })
    setup([status, balance, flagged])
    fireEvent.click(globalButton())

    const panel = screen.getByTestId('still-requires-decision')
    expect(panel).toHaveTextContent(/Required balance decision/)
    expect(panel).toHaveTextContent(/Other — Flagged/)
  })
})
