import React from 'react'
import { ShieldCheck } from './Navbar.jsx'

export default function SessionEnded({ onStartNew }) {
  return (
    <div className="session-ended">
      <div className="session-ended-card">
        <div className="session-ended-logo">
          <ShieldCheck />
          <div>
            <div className="brand">Audit Intelligence</div>
            <div className="subtitle">Data Audit &amp; Compliance</div>
          </div>
        </div>
        <h2>Session ended</h2>
        <p>Thank you for using Audit Intelligence. Your session has been closed.</p>
        <button className="btn btn-primary" onClick={onStartNew}>Start New Session</button>
      </div>
    </div>
  )
}
