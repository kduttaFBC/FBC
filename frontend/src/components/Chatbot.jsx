import React, { useState, useRef, useEffect } from 'react'
import axios from 'axios'

function buildOpeningMessage(summary, results) {
  const issues = summary?.total_issues ?? 0
  const flagged = (results || []).filter((r) => r.status !== 'PASS')
  const checksWithIssues = flagged.length

  const top = flagged
    .filter((r) => r.issues_found > 0)
    .sort((a, b) => b.issues_found - a.issues_found)
    .slice(0, 3)
    .map((r) => `${r.name}${r.servicer ? ` (${r.servicer})` : ''} — ${r.issues_found}`)

  const summaryLine = top.length
    ? ` Top findings: ${top.join(', ')}.`
    : ' Everything looks clean.'

  return `Audit complete. I found ${issues} issue(s) across ${checksWithIssues} check(s).${summaryLine} What would you like to explore?`
}

export default function Chatbot({ auditContext, summary, results }) {
  const [messages, setMessages] = useState(() => [
    { role: 'assistant', content: buildOpeningMessage(summary, results) },
  ])
  const [input, setInput] = useState('')
  const [sending, setSending] = useState(false)
  const scrollRef = useRef(null)

  useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight
    }
  }, [messages, sending])

  const send = async () => {
    const text = input.trim()
    if (!text || sending) return

    const history = messages.map((m) => ({ role: m.role, content: m.content }))
    setMessages((m) => [...m, { role: 'user', content: text }])
    setInput('')
    setSending(true)

    try {
      const apiUrl = import.meta.env.VITE_API_URL || ''
      const { data } = await axios.post(`${apiUrl}/api/chat`, {
        message: text,
        history,
        audit_context: auditContext,
      })
      setMessages((m) => [...m, {
        role: 'assistant',
        content: data.reply,
        excelB64: data.excel_b64 || null,
        excelFilename: data.excel_filename || null,
      }])
    } catch (err) {
      const msg =
        err.response?.data?.error ||
        'Sorry, I had trouble responding. Please try again.'
      setMessages((m) => [...m, { role: 'assistant', content: msg }])
    } finally {
      setSending(false)
    }
  }

  const onKeyDown = (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      send()
    }
  }

  const downloadExcel = (base64, filename) => {
    const byteCharacters = atob(base64)
    const byteNumbers = Array.from(byteCharacters, (c) => c.charCodeAt(0))
    const byteArray = new Uint8Array(byteNumbers)
    const blob = new Blob([byteArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = filename || 'FalconBridge_Chat_Export.xlsx'
    a.click()
    URL.revokeObjectURL(url)
  }

  return (
    <div className="chatbot">
      <div className="chat-header">
        <span className="status-dot" />
        <span className="title">FalconBridge AI</span>
        <span className="powered">Powered by Claude</span>
      </div>

      <div className="chat-messages" ref={scrollRef}>
        {messages.map((m, i) => (
          <div key={i} className={`bubble ${m.role === 'user' ? 'user' : 'bot'}`}>
            {m.content}
            {m.excelB64 && (
              <div className="chat-excel-download">
                <button onClick={() => downloadExcel(m.excelB64, m.excelFilename)}>
                  ⬇ Download Excel
                </button>
              </div>
            )}
          </div>
        ))}
        {sending && <div className="bubble bot typing">FalconBridge AI is typing…</div>}
      </div>

      <div className="chat-input-row">
        <input
          type="text"
          placeholder="Ask about the audit data…"
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={onKeyDown}
          disabled={sending}
        />
        <button onClick={send} disabled={sending || !input.trim()}>Send</button>
      </div>
    </div>
  )
}
