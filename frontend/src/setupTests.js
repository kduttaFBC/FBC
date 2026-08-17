// Vitest setup, loaded once before every frontend test file.
//
// Adds jest-dom's DOM matchers (toBeInTheDocument, toBeDisabled, ...) and
// unmounts React trees between tests so component state cannot leak from one
// assertion into the next -- which matters most for the bulk-approval tests
// this harness exists to support, since those assert on accumulated decision
// state.
import '@testing-library/jest-dom/vitest'
import { cleanup } from '@testing-library/react'
import { afterEach } from 'vitest'

afterEach(() => {
  cleanup()
})
