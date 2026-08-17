import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    strictPort: true,
    proxy: {
      '/api': 'http://localhost:8000'
    }
  },
  define: {
    'import.meta.env.VITE_API_URL': JSON.stringify(process.env.VITE_API_URL || '')
  },
  // Frontend test harness. Added as infrastructure only -- no production file
  // is touched by this block, and `npm run build` is unaffected.
  //
  // jsdom rather than a real browser: the behaviour under test is component
  // state (which recommendations a bulk action pre-fills, which it must leave
  // alone), not rendering fidelity.
  test: {
    environment: 'jsdom',
    globals: true,
    setupFiles: ['./src/setupTests.js'],
    // Restricted to src/ so a stray test file elsewhere in the repo is never
    // picked up by the frontend runner.
    include: ['src/**/*.{test,spec}.{js,jsx}']
  }
})
