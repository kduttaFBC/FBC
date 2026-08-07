import React, { useState, useEffect, useMemo } from "react";

/* ------------------------------------------------------------------ *
 * PROTOTYPE — Correction Review
 *
 * Populated with the four real corrections from the Peoples Bank
 * July 2026 run (Audit_Intelligence_Reconciliation_2026-08-07).
 *
 * Decisions are held in this browser only. The payload panel at the
 * bottom shows exactly what would POST to the backend, which is where
 * the audit record has to live — an approval needs reviewer identity,
 * timestamp and reason, and none of that can be attributed here.
 * ------------------------------------------------------------------ */

const LADDER = ["CURRENT", "DELINQUENT", "30 DAYS", "60 DAYS", "90 DAYS", "120 DAYS", "150+ DAYS"];

const RUNG_COLOR = {
  "CURRENT": "#2E7D64",
  "DELINQUENT": "#8A7A2E",
  "30 DAYS": "#A8702F",
  "60 DAYS": "#B15A34",
  "90 DAYS": "#A94436",
  "120 DAYS": "#93343A",
  "150+ DAYS": "#7A2739",
};

const CORRECTIONS = [
  {
    id: "c1", gp: "1278409107", field: "Delinquency Status",
    rows: [11], postings: 1, sheet: "Servicer Inputs",
    reported: "30 DAYS", prior: "DELINQUENT", calculated: "DELINQUENT",
    paymentReceived: 463.33, scheduledPayment: 450.66, balance: 57803.33,
    autoApplied: true, reviewFlag: null,
    rule: "STAT-01 deterministic calculation",
  },
  {
    id: "c2", gp: "1327409104", field: "Delinquency Status",
    rows: [12], postings: 1, sheet: "Servicer Inputs",
    reported: "DELINQUENT", prior: "CURRENT", calculated: "CURRENT",
    paymentReceived: 805.62, scheduledPayment: 779.64, balance: 99952.43,
    autoApplied: true, reviewFlag: null,
    rule: "STAT-01 deterministic calculation",
  },
  {
    id: "c3", gp: "3883228810", field: "Delinquency Status",
    rows: [31], postings: 1, sheet: "Servicer Inputs",
    reported: "DELINQUENT", prior: "CURRENT", calculated: "CURRENT",
    paymentReceived: 1347.16, scheduledPayment: 1260.57, balance: 158185.76,
    autoApplied: true, reviewFlag: null,
    rule: "STAT-01 deterministic calculation",
  },
  {
    id: "c4", gp: "4017399101", field: "Delinquency Status",
    rows: [33, 34], postings: 2, sheet: "Servicer Inputs",
    reported: "FORCLOSURE", prior: "60 DAYS", calculated: "30 DAYS",
    paymentReceived: 9876.22, scheduledPayment: 4697.01, balance: 584023.61,
    autoApplied: false,
    reviewFlag: "Reported Delinquency Status contains an invalid cross-domain value.",
    rule: "STAT-01 deterministic calculation",
  },
];

const DECISIONS = [
  { v: "pending", label: "Not yet decided" },
  { v: "approve", label: "Approve — write to workbook" },
  { v: "reject", label: "Reject — keep the reported value" },
  { v: "return", label: "Return to servicer — source defect" },
];

const money = (n) => n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/* ---------------------------------------------------------------- */

function Ladder({ reported, prior, calculated }) {
  const reportedIdx = LADDER.indexOf(reported);
  const priorIdx = LADDER.indexOf(prior);
  const calcIdx = LADDER.indexOf(calculated);
  const offLadder = reportedIdx === -1;

  return (
    <div className="ladder">
      <div className="ladder-rungs">
        {LADDER.map((rung, i) => {
          const isPrior = i === priorIdx;
          const isCalc = i === calcIdx;
          const isRep = i === reportedIdx;
          return (
            <div key={rung} className="rung">
              <div
                className={`rung-bar${isCalc ? " rung-bar-calc" : ""}`}
                style={{ background: RUNG_COLOR[rung], opacity: isCalc ? 1 : isPrior || isRep ? 0.55 : 0.14 }}
              />
              <div className="rung-label mono">{rung}</div>
              <div className="rung-marks">
                {isRep && <span className="mark mark-rep mono">reported</span>}
                {isPrior && <span className="mark mark-prior mono">last cycle</span>}
                {isCalc && <span className="mark mark-calc mono">calculated</span>}
              </div>
            </div>
          );
        })}
      </div>

      {offLadder && (
        <div className="off-ladder">
          <div className="off-ladder-chip mono">{reported}</div>
          <div className="off-ladder-note">
            The reported value has no position on this ladder. It is not a delinquency
            status — it belongs to the Loan Status field, and is misspelled there too.
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------- */

export default function CorrectionReview() {
  const [decisions, setDecisions] = useState({});
  const [notes, setNotes] = useState({});
  const [notify, setNotify] = useState({ c4: true });
  const [selectedId, setSelectedId] = useState("c4");
  const [reviewer, setReviewer] = useState("");
  const [filter, setFilter] = useState("all");
  const [loaded, setLoaded] = useState(false);
  const [saveNote, setSaveNote] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const r = await window.storage.get("review-draft");
        if (r) {
          const d = JSON.parse(r.value);
          setDecisions(d.decisions || {});
          setNotes(d.notes || {});
          setNotify(d.notify || { c4: true });
          setReviewer(d.reviewer || "");
        }
      } catch (e) {
        /* no draft yet */
      }
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    const t = setTimeout(async () => {
      try {
        await window.storage.set("review-draft", JSON.stringify({ decisions, notes, notify, reviewer }));
        setSaveNote("Draft saved");
        setTimeout(() => setSaveNote(""), 1600);
      } catch (e) {
        setSaveNote("Draft not saved");
      }
    }, 500);
    return () => clearTimeout(t);
  }, [decisions, notes, notify, reviewer, loaded]);

  const decisionOf = (id) => decisions[id] || "pending";
  const selected = CORRECTIONS.find((c) => c.id === selectedId);

  const visible = useMemo(() => {
    if (filter === "review") return CORRECTIONS.filter((c) => c.reviewFlag);
    if (filter === "pending") return CORRECTIONS.filter((c) => decisionOf(c.id) === "pending");
    if (filter === "decided") return CORRECTIONS.filter((c) => decisionOf(c.id) !== "pending");
    return CORRECTIONS;
  }, [filter, decisions]);

  const counts = useMemo(() => {
    const c = { approve: 0, reject: 0, return: 0, pending: 0 };
    CORRECTIONS.forEach((x) => { c[decisionOf(x.id)] += 1; });
    return c;
  }, [decisions]);

  const payload = useMemo(() => ({
    audit_run: "Audit_Intelligence_Reconciliation_2026-08-07",
    servicer: "Peoples Bank",
    cycle: "2026-07",
    reviewed_by: reviewer || null,
    reviewed_on: new Date().toISOString().replace(/\.\d{3}Z$/, "Z"),
    decisions: CORRECTIONS.filter((c) => decisionOf(c.id) !== "pending").map((c) => ({
      sba_gp_number: c.gp,
      field: c.field,
      excel_rows: c.rows,
      reported_value: c.reported,
      proposed_value: c.calculated,
      decision: decisionOf(c.id),
      review_reason: c.reviewFlag || null,
      notify_servicer: !!notify[c.id],
      reviewer_note: notes[c.id] || null,
    })),
  }), [decisions, notes, notify, reviewer]);

  const blocked = !reviewer.trim() && payload.decisions.length > 0;

  return (
    <div className="app">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;450;500;600&display=swap');

        .app {
          --ink:#0E1621; --ink-2:#3D4C5C; --ink-3:#6B7A8A;
          --paper:#EDF0F3; --card:#FFFFFF; --rule:#D3DAE1;
          --flag:#93343A; --ok:#2E7D64;
          font-family:'IBM Plex Sans',system-ui,sans-serif;
          background:var(--paper); color:var(--ink);
          min-height:100vh; padding:0 0 40px;
          -webkit-font-smoothing:antialiased;
        }
        .mono{font-family:'IBM Plex Mono',ui-monospace,monospace;}

        .topbar{background:var(--ink);color:#E8EDF2;padding:18px 26px;
          display:flex;flex-wrap:wrap;gap:22px;align-items:baseline;}
        .topbar h1{font-size:15px;font-weight:600;letter-spacing:.01em;margin:0;}
        .topbar .meta{font-size:11.5px;color:#93A3B4;letter-spacing:.04em;}
        .topbar .spacer{flex:1;}

        .strip{display:flex;flex-wrap:wrap;gap:0;border-bottom:1px solid var(--rule);
          background:var(--card);}
        .stat{padding:13px 22px;border-right:1px solid var(--rule);min-width:112px;}
        .stat .n{font-size:20px;font-weight:600;line-height:1.1;}
        .stat .k{font-size:10.5px;letter-spacing:.09em;text-transform:uppercase;
          color:var(--ink-3);margin-top:3px;}
        .stat-flag .n{color:var(--flag);}

        .wrap{display:grid;grid-template-columns:minmax(300px,1fr) minmax(0,1.75fr);
          gap:22px;padding:22px 26px;max-width:1320px;margin:0 auto;align-items:start;}
        @media (max-width:940px){.wrap{grid-template-columns:1fr;}}

        .panel{background:var(--card);border:1px solid var(--rule);}
        .panel-head{padding:12px 16px;border-bottom:1px solid var(--rule);
          display:flex;align-items:center;gap:10px;flex-wrap:wrap;}
        .panel-head h2{font-size:11px;letter-spacing:.1em;text-transform:uppercase;
          color:var(--ink-3);margin:0;font-weight:600;}

        select,input[type=text],textarea{font-family:inherit;font-size:13px;color:var(--ink);
          background:#fff;border:1px solid var(--rule);border-radius:2px;padding:7px 9px;}
        select:focus,input:focus,textarea:focus{outline:2px solid var(--ink);outline-offset:1px;}
        textarea{width:100%;resize:vertical;min-height:74px;line-height:1.5;}

        .qitem{display:block;width:100%;text-align:left;background:none;border:0;
          border-bottom:1px solid var(--rule);padding:13px 16px;cursor:pointer;}
        .qitem:hover{background:#F6F8FA;}
        .qitem[aria-current=true]{background:#F0F4F8;box-shadow:inset 3px 0 0 var(--ink);}
        .qitem-top{display:flex;align-items:baseline;gap:9px;}
        .qgp{font-size:13.5px;font-weight:500;letter-spacing:.01em;}
        .qmove{font-size:11.5px;color:var(--ink-2);margin-top:5px;letter-spacing:.01em;}
        .qmove b{font-weight:500;}
        .flagdot{width:6px;height:6px;border-radius:50%;background:var(--flag);
          display:inline-block;flex:none;}

        .pill{font-size:9.5px;letter-spacing:.08em;text-transform:uppercase;
          padding:2px 6px;border:1px solid currentColor;border-radius:2px;font-weight:600;}
        .pill-review{color:var(--flag);}
        .pill-applied{color:var(--ok);}
        .pill-pending{color:var(--ink-3);}

        .detail-head{padding:18px 22px;border-bottom:1px solid var(--rule);}
        .detail-gp{font-size:22px;font-weight:600;letter-spacing:-.01em;}
        .detail-sub{font-size:11.5px;color:var(--ink-3);margin-top:4px;letter-spacing:.03em;}

        .flagbox{background:#FBF1F1;border-left:3px solid var(--flag);
          padding:12px 15px;margin:0;font-size:13px;line-height:1.55;color:#5C2228;}

        .section{padding:18px 22px;border-bottom:1px solid var(--rule);}
        .section > h3{font-size:10.5px;letter-spacing:.1em;text-transform:uppercase;
          color:var(--ink-3);margin:0 0 14px;font-weight:600;}

        .ladder{display:flex;flex-direction:column;gap:0;}
        .ladder-rungs{display:flex;flex-direction:column;gap:2px;}
        .rung{display:grid;grid-template-columns:5px 90px 1fr;gap:11px;align-items:center;height:26px;}
        .rung-bar{height:26px;border-radius:1px;}
        .rung-bar-calc{box-shadow:0 0 0 1px rgba(0,0,0,.18);}
        .rung-label{font-size:11px;letter-spacing:.02em;color:var(--ink-2);}
        .rung-marks{display:flex;gap:6px;}
        .mark{font-size:9.5px;letter-spacing:.07em;text-transform:uppercase;
          padding:2px 6px;border-radius:2px;font-weight:500;}
        .mark-rep{background:#E4E9EE;color:var(--ink-2);}
        .mark-prior{background:#F0EDE2;color:#6B6132;}
        .mark-calc{background:var(--ink);color:#fff;}

        .off-ladder{margin-top:14px;padding:13px 15px;background:#FBF1F1;
          border-left:3px solid var(--flag);display:flex;gap:13px;align-items:flex-start;}
        .off-ladder-chip{font-size:12.5px;font-weight:600;color:#fff;background:var(--flag);
          padding:4px 9px;border-radius:2px;flex:none;letter-spacing:.02em;}
        .off-ladder-note{font-size:12.5px;line-height:1.55;color:#5C2228;}

        .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(148px,1fr));
          gap:1px;background:var(--rule);border:1px solid var(--rule);}
        .cell{background:#fff;padding:11px 13px;}
        .cell .k{font-size:9.5px;letter-spacing:.09em;text-transform:uppercase;
          color:var(--ink-3);margin-bottom:5px;}
        .cell .v{font-size:13.5px;letter-spacing:.01em;}

        .change{display:flex;align-items:center;gap:13px;flex-wrap:wrap;}
        .val{padding:7px 12px;border:1px solid var(--rule);border-radius:2px;
          font-size:14px;background:#fff;}
        .val-from{color:var(--ink-2);text-decoration:line-through;text-decoration-thickness:1px;}
        .val-to{border-color:var(--ink);font-weight:500;}
        .arrow{color:var(--ink-3);font-size:15px;}

        .field{margin-bottom:15px;}
        .field label{display:block;font-size:10.5px;letter-spacing:.09em;
          text-transform:uppercase;color:var(--ink-3);margin-bottom:6px;font-weight:600;}
        .check{display:flex;gap:9px;align-items:flex-start;font-size:13px;line-height:1.5;cursor:pointer;}
        .check input{margin-top:2px;width:15px;height:15px;accent-color:#0E1621;flex:none;}

        .btn{font-family:inherit;font-size:12.5px;font-weight:500;letter-spacing:.02em;
          padding:9px 15px;border:1px solid var(--ink);background:var(--ink);color:#fff;
          border-radius:2px;cursor:pointer;}
        .btn:disabled{opacity:.35;cursor:not-allowed;}
        .btn-ghost{background:none;color:var(--ink);}
        .btn-ghost:hover{background:#F0F4F8;}

        pre{font-family:'IBM Plex Mono',monospace;font-size:11.5px;line-height:1.6;
          background:var(--ink);color:#CBD6E0;padding:15px;overflow:auto;max-height:290px;
          margin:0;border-radius:2px;}
        .warn{font-size:12.5px;color:var(--flag);margin-top:10px;line-height:1.5;}
        .savenote{font-size:11px;color:var(--ink-3);letter-spacing:.04em;}
        @media (prefers-reduced-motion:no-preference){
          .qitem{transition:background .12s ease;}
        }
      `}</style>

      <div className="topbar">
        <h1>Correction review</h1>
        <span className="meta mono">PEOPLES BANK · CYCLE 2026-07 · SOUP 2024-1</span>
        <span className="spacer" />
        <span className="meta mono">AUDIT 2026-08-07</span>
      </div>

      <div className="strip">
        <div className="stat"><div className="n">{CORRECTIONS.length}</div><div className="k">Corrections</div></div>
        <div className="stat stat-flag"><div className="n">{CORRECTIONS.filter(c => c.reviewFlag).length}</div><div className="k">Need review</div></div>
        <div className="stat"><div className="n">{counts.approve}</div><div className="k">Approved</div></div>
        <div className="stat"><div className="n">{counts.reject + counts.return}</div><div className="k">Declined</div></div>
        <div className="stat"><div className="n">{counts.pending}</div><div className="k">Undecided</div></div>
      </div>

      <div className="wrap">
        {/* ---------------- queue ---------------- */}
        <div className="panel">
          <div className="panel-head">
            <h2>Queue</h2>
            <span className="spacer" style={{ flex: 1 }} />
            <select value={filter} onChange={(e) => setFilter(e.target.value)} aria-label="Filter queue">
              <option value="all">All corrections</option>
              <option value="review">Needs review</option>
              <option value="pending">Undecided</option>
              <option value="decided">Decided</option>
            </select>
          </div>

          {visible.length === 0 && (
            <div style={{ padding: "26px 16px", fontSize: 13, color: "var(--ink-3)" }}>
              Nothing matches this filter. Choose another view.
            </div>
          )}

          {visible.map((c) => {
            const d = decisionOf(c.id);
            return (
              <button
                key={c.id}
                className="qitem"
                aria-current={c.id === selectedId}
                onClick={() => setSelectedId(c.id)}
              >
                <div className="qitem-top">
                  {c.reviewFlag && <span className="flagdot" aria-hidden="true" />}
                  <span className="qgp mono">{c.gp}</span>
                  <span style={{ flex: 1 }} />
                  <span className={`pill ${c.reviewFlag ? "pill-review" : d === "pending" ? "pill-applied" : "pill-pending"}`}>
                    {c.reviewFlag ? "Review" : d === "pending" ? "Applied" : d}
                  </span>
                </div>
                <div className="qmove mono">
                  {c.reported} <span style={{ color: "var(--ink-3)" }}>→</span> <b>{c.calculated}</b>
                </div>
                <div className="qmove" style={{ color: "var(--ink-3)", marginTop: 3 }}>
                  {c.postings} posting{c.postings > 1 ? "s" : ""} · row{c.rows.length > 1 ? "s" : ""}{" "}
                  <span className="mono">{c.rows.join(", ")}</span>
                </div>
              </button>
            );
          })}
        </div>

        {/* ---------------- detail ---------------- */}
        {selected && (
          <div className="panel">
            <div className="detail-head">
              <div className="detail-gp mono">{selected.gp}</div>
              <div className="detail-sub">
                {selected.field} · {selected.sheet} · row{selected.rows.length > 1 ? "s" : ""}{" "}
                <span className="mono">{selected.rows.join(", ")}</span> · {selected.rule}
              </div>
            </div>

            {selected.reviewFlag && (
              <p className="flagbox">
                <strong>Human review required.</strong> {selected.reviewFlag} This correction was
                calculated but not written to the workbook.
              </p>
            )}

            <div className="section">
              <h3>Proposed change</h3>
              <div className="change">
                <span className="val val-from mono">{selected.reported}</span>
                <span className="arrow" aria-hidden="true">→</span>
                <span className="val val-to mono">{selected.calculated}</span>
                <span style={{ fontSize: 12.5, color: "var(--ink-3)" }}>
                  across {selected.rows.length} posting row{selected.rows.length > 1 ? "s" : ""}
                </span>
              </div>
            </div>

            <div className="section">
              <h3>Position on the delinquency ladder</h3>
              <Ladder reported={selected.reported} prior={selected.prior} calculated={selected.calculated} />
            </div>

            <div className="section">
              <h3>Evidence</h3>
              <div className="grid">
                <div className="cell"><div className="k">Servicer reported</div><div className="v mono">{selected.reported}</div></div>
                <div className="cell"><div className="k">Last cycle (SOUP)</div><div className="v mono">{selected.prior}</div></div>
                <div className="cell"><div className="k">Calculated</div><div className="v mono">{selected.calculated}</div></div>
                <div className="cell"><div className="k">Payment received</div><div className="v mono">{money(selected.paymentReceived)}</div></div>
                <div className="cell"><div className="k">Scheduled payment</div><div className="v mono">{money(selected.scheduledPayment)}</div></div>
                <div className="cell">
                  <div className="k">Payment units</div>
                  <div className="v mono">{(selected.paymentReceived / selected.scheduledPayment).toFixed(3)}</div>
                </div>
                <div className="cell"><div className="k">Ending balance</div><div className="v mono">{money(selected.balance)}</div></div>
                <div className="cell"><div className="k">Postings</div><div className="v mono">{selected.postings}</div></div>
              </div>
            </div>

            <div className="section">
              <h3>Your decision</h3>

              <div className="field">
                <label htmlFor="dec">Decision</label>
                <select
                  id="dec"
                  value={decisionOf(selected.id)}
                  onChange={(e) => setDecisions({ ...decisions, [selected.id]: e.target.value })}
                  style={{ minWidth: 280 }}
                >
                  {DECISIONS.map((d) => <option key={d.v} value={d.v}>{d.label}</option>)}
                </select>
              </div>

              <div className="field">
                <label className="check" htmlFor="notify">
                  <input
                    id="notify"
                    type="checkbox"
                    checked={!!notify[selected.id]}
                    onChange={(e) => setNotify({ ...notify, [selected.id]: e.target.checked })}
                  />
                  <span>
                    Tell the servicer about this so it does not recur next cycle
                    {selected.reviewFlag && (
                      <span style={{ color: "var(--ink-3)" }}> — checked by default for a source data defect</span>
                    )}
                  </span>
                </label>
              </div>

              <div className="field">
                <label htmlFor="note">Reviewer note</label>
                <textarea
                  id="note"
                  value={notes[selected.id] || ""}
                  placeholder={
                    selected.reviewFlag
                      ? "Why this decision, and what the servicer should change in their file."
                      : "Optional. Anything a later reader would need to understand this decision."
                  }
                  onChange={(e) => setNotes({ ...notes, [selected.id]: e.target.value })}
                />
              </div>
            </div>
          </div>
        )}
      </div>

      {/* ---------------- record ---------------- */}
      <div className="wrap" style={{ paddingTop: 0 }}>
        <div className="panel" style={{ gridColumn: "1 / -1" }}>
          <div className="panel-head">
            <h2>Audit record</h2>
            <span style={{ flex: 1 }} />
            <span className="savenote">{saveNote}</span>
          </div>
          <div className="section" style={{ borderBottom: "1px solid var(--rule)" }}>
            <div className="field" style={{ marginBottom: 0, maxWidth: 340 }}>
              <label htmlFor="rev">Reviewed by</label>
              <input
                id="rev" type="text" value={reviewer}
                placeholder="Your name"
                onChange={(e) => setReviewer(e.target.value)}
                style={{ width: "100%" }}
              />
            </div>
            {blocked && (
              <p className="warn">
                Add your name before recording {payload.decisions.length} decision
                {payload.decisions.length > 1 ? "s" : ""}. An approval without an attributable
                reviewer is not an audit record.
              </p>
            )}
          </div>

          <div className="section" style={{ borderBottom: 0 }}>
            <h3>What gets sent to the backend</h3>
            <pre>{JSON.stringify(payload, null, 2)}</pre>
            <div style={{ display: "flex", gap: 10, marginTop: 14, flexWrap: "wrap" }}>
              <button className="btn" disabled={blocked || payload.decisions.length === 0}>
                Record {payload.decisions.length || "no"} decision{payload.decisions.length === 1 ? "" : "s"}
              </button>
              <button
                className="btn btn-ghost"
                onClick={async () => {
                  setDecisions({}); setNotes({}); setNotify({ c4: true }); setReviewer("");
                  try { await window.storage.delete("review-draft"); } catch (e) {}
                }}
              >
                Clear draft
              </button>
            </div>
            <p style={{ fontSize: 12, color: "var(--ink-3)", marginTop: 13, lineHeight: 1.55, maxWidth: 620 }}>
              Prototype. Decisions live in this browser only and the button is inert. The record
              has to be written by the backend, which is the only place reviewer identity and
              timestamp can be attributed and kept.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}
