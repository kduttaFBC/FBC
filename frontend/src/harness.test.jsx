/**
 * Frontend test harness verification.
 *
 * INFRASTRUCTURE ONLY. This file asserts that the toolchain itself works --
 * the runner, jsdom, React rendering, the jest-dom matchers, and the
 * between-test cleanup -- so that when the global Approve All tests are added
 * next, a failure there is a failure of the feature and not of the harness.
 *
 * It touches no production module and encodes no business rule.
 */
import React, { useState } from 'react'
import { describe, it, expect } from 'vitest'
import { render, screen, fireEvent } from '@testing-library/react'

describe('frontend test harness', () => {
  it('runs a test at all', () => {
    expect(true).toBe(true)
  })

  it('provides a jsdom document', () => {
    expect(typeof document).toBe('object')
    expect(document.createElement('div')).toBeTruthy()
  })

  it('renders a React component', () => {
    render(<p>audit intelligence</p>)
    expect(screen.getByText('audit intelligence')).toBeInTheDocument()
  })

  it('exposes the jest-dom matchers', () => {
    render(<button disabled>Approve All</button>)
    expect(screen.getByRole('button', { name: 'Approve All' })).toBeDisabled()
  })

  it('tracks state across an interaction', () => {
    // The shape the bulk-approval tests will use: a click mutates local state
    // and nothing is submitted.
    function Counter() {
      const [decided, setDecided] = useState(0)
      return <button onClick={() => setDecided((n) => n + 1)}>decided {decided}</button>
    }
    render(<Counter />)
    const button = screen.getByRole('button')
    expect(button).toHaveTextContent('decided 0')
    fireEvent.click(button)
    expect(button).toHaveTextContent('decided 1')
  })

  it('cleans up between tests', () => {
    // If afterEach(cleanup) were missing, the tree rendered above would still
    // be mounted here and this query would find it.
    expect(screen.queryByRole('button')).toBeNull()
  })
})
