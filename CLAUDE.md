# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Full local dev (from repo root) — starts backend (:8000) + frontend (:5173) together
npm install          # once
npm run dev

# Backend only
cd backend && python3 -m uvicorn main:app --reload --port 8000

# Frontend only (Vite dev server proxies /api/* -> :8000)
npm --prefix frontend run dev

# Tests — pytest, run from backend/ (tests import modules by bare name, e.g. `import checks`)
cd backend && python3 -m pytest
python3 -m pytest test_corrections.py                                  # one file
python3 -m pytest test_corrections.py::TestClass::test_name            # one test
```

Backend requires `ANTHROPIC_API_KEY` (put it in `backend/.env`) for the chatbot; the audit pipeline itself runs without it. Python 3.11 (`backend/runtime.txt`).

## What this app actually does

**Ignore README.md's "14 checks / Customer Master + Transactions" description — it is stale.** The live app reconciles a monthly **SOUP Performance Dataset** (~25k-row loan performance export, the system of record) against one or more **servicer SRF submissions** (Servicer Inputs + Loan Terms), with an optional **Servicer Memo Summary** for deferral awareness. It produces reconciliation findings, downloadable Excel workbooks, and an AI chatbot grounded in the audited data.

## Architecture

Request flow: `frontend` → `POST /api/audit` (main.py) → per servicer file: `preflight` → `file_readers` → `checks.run_servicer_reconciliation` → `corrections` → `report`. Chat is a separate `POST /api/chat`.

### Backend modules (`backend/`)
- **main.py** — FastAPI app, the two endpoints, session cache, and all response-shaping (result capping for the chatbot, report worksheet filtering, mixed-reporting-period guard, markdown-table→Excel export for chat). Business logic lives elsewhere; main.py orchestrates and shapes.
- **file_readers.py** — Smart readers for SOUP / servicer / memo Excel files. Uses **python-calamine** (fast, but **cached values only — no formula access**). Scans up to `MAX_HEADER_SCAN_ROWS` for the real header row instead of assuming row 1. Servicer identity comes from the **upload filename convention `"{soupKey}||{displayLabel}__{originalName}"`** (set in the frontend), not from in-sheet labels. `MEMO_LENDER_ALIASES` and `SERVICER_SOUP_KEYS`-style mappings handle name spelling variants — a missing mapping silently yields zero SOUP matches, so add entries here when onboarding a servicer.
- **preflight.py** — Runs on **raw bytes before Calamine**. Because Calamine can't distinguish a formula from a pasted literal (or detect openpyxl's missing value cache), this is a **blocking re-ingestion control**: it fails workbooks with overwritten formula columns or missing cached values (e.g. this app's own corrected-workbook output, which must be opened+recalculated in Excel before re-entry). Reports only; the caller decides fail/warn.
- **checks.py** — All reconciliation checks. Each returns a result dict (`id`, `status` PASS/WARN/FAIL/INFO, `issues_found`, `records`, ...). `run_servicer_reconciliation` runs them all and tags each id/servicer as `"{CHECK}::{displayLabel}"`. Checks: POP-01, BAL-01/02, STAT-01/02, PAY-01/02, DUP-01, NEG-01, PORT-01, DQ-01. **STAT-01 (deferral-aware status/delinquency match) is the most complex** and also produces the data corrections.py and the Audited Portfolio Summary consume. Multiple postings in one cycle are consolidated via the single shared `consolidate_cycle_postings` (BAL-01/STAT-01/PAY-02/NEG-01 must not re-implement this independently).
- **corrections.py** — **Isolated** corrected-workbook generator. Reads raw facts off STAT-01's `correction_data` and applies a strict mechanical allow-list; it is the only place that writes to a copy of an uploaded workbook and never touches audit results/reports. v1 scope is deliberately narrow (Delinquency Status when unambiguous; Loan Status → DEFERRED only for confirmed full-P&I start/mid). Keep the audit engine and this module non-interfering.
- **report.py** — openpyxl Excel generation: the multi-tab reconciliation workbook plus the Original/Audited Portfolio Summary workbooks.
- **chatbot.py** — Anthropic SDK chat (model `claude-haiku-4-5`) with tool use (`lookup_loan_history`, `search_accounts_by_status`). Stateless API, so full audit context + history is sent each call. main.py keeps the full SOUP DataFrame in `SESSION_CACHE` (1h TTL) so tools can query any loan on demand.
- **business_rule_BR002.py**, **srf_validator.py** — Standalone modules **not wired into the main.py pipeline** (BR-002 is pending client approval; srf_validator is a fail-closed pre-import gate). Don't assume the audit endpoint runs them.

### Frontend (`frontend/src/`)
React + Vite, screen state machine in **App.jsx** (`upload → ready → results → ended`). App.jsx also holds the canonical `SOUP_FIELDS` / `SERVICER_FIELDS` / `MEMO_FIELDS` (mirror backend `file_readers`) and the `SERVICER_SLOTS` / `SERVICER_SOUP_KEYS` maps that build the upload filename convention. Client-side validation (SheetJS/`xlsx`) mirrors the backend's header scan so what validates in-browser matches what the backend reads.

## Conventions worth knowing
- **Client-confirmed business rules** are marked with `CLIENT-CONFIRMED (date)` comments in the code — these encode decisions (e.g. mixed-reporting-period handling, report tab inclusion tiers, deferral classification) that are not derivable from the data. Treat them as requirements, not incidental comments.
- `main.py`'s `_filter_report_results` decides which checks get an Excel tab (some always excluded, some conditional on findings). A missing tab is never evidence a check didn't run — every check's full result is always in the JSON response regardless.
- Servicers split across note series (e.g. "Lendistry 2025"/"2024") run as separate independent reconciliations against the same SOUP servicer; `soupKey` matches SOUP's Servicer column, `displayLabel` is what appears in tabs/summary/chat.
- **`correction_data` review-evidence fields are read-only.** STAT-01 exposes some already-computed figures on `correction_data` (e.g. `prior_dq`, `payment_received`, `scheduled_payment`, `servicing_fee`, `ending_balance`, `posting_count`) purely so the correction review page can display them. These are evidence for a human, never inputs — nothing in `checks.py` calculation logic (payment-sufficiency / BR-002, the delinquency ladder, `consolidate_cycle_postings`) may ever read them back. Adding a new evidence field is fine; feeding one into a calculation is not.
- **Corrections are review-gated (`corrections.py`).** `/api/audit` produces *recommendations only* — no workbook is written there. `corrections.build_recommendations` shapes STAT-01's eligible candidates; `main.py` stashes each servicer's original bytes + SHA-256 + candidate set in `SESSION_CACHE` (process-local — same single-worker assumption the chatbot's `soup_df` relies on; they fail together under `--workers >1`). `POST /api/corrections/review` then validates the reviewer's decisions, opens a fresh copy of the original bytes, writes only approved values (domain-restricted via `validate_approved_value`), and logs every approve/reject/return. `build_eligible_corrections` (the eligibility allow-list) decides only what is a *candidate* — never what gets written.
