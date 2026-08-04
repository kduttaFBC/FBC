import React from 'react'

export function ShieldCheck() {
  return (
    <svg className="logo-icon" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path
        d="M12 2L4 5v6c0 5 3.4 8.6 8 11 4.6-2.4 8-6 8-11V5l-8-3z"
        fill="#185FA5"
        stroke="#ffffff"
        strokeWidth="1.2"
      />
      <path
        d="M8.5 12l2.5 2.5 4.5-5"
        stroke="#ffffff"
        strokeWidth="1.8"
        strokeLinecap="round"
        strokeLinejoin="round"
        fill="none"
      />
    </svg>
  )
}

export default function Navbar() {
  return (
    <nav className="navbar">
      <ShieldCheck />
      <div>
        <div className="brand">Audit Intelligence</div>
        <div className="subtitle">Data Audit &amp; Compliance</div>
      </div>
    </nav>
  )
}
