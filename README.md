# FalconBridge Audit Intelligence

FalconBridge Audit Intelligence is a data-quality and compliance audit tool for community banks. It accepts two Excel uploads — a Customer Master file and a Banking Transactions file — runs 14 audit checks across data completeness, payment integrity, and loan math, produces a formatted, multi-tab downloadable Excel report, and includes an AI chatbot (powered by the Anthropic API) that answers questions grounded in the audited data.

## Local setup

```bash
# Backend
cd backend
pip install -r requirements.txt
echo "ANTHROPIC_API_KEY=your_key_here" > .env
uvicorn main:app --reload

# Frontend (separate terminal)
cd frontend
npm install
npm run dev
```

The app runs at: **http://localhost:5173**

The Vite dev server proxies `/api/*` to the backend at `http://localhost:8000`, so make sure the backend is running on port 8000 (the `uvicorn` default).

## How it works

1. **Upload** — drop in both Excel files. The browser validates file type, non-emptiness, and the presence of all required columns (case-insensitive) before anything is sent to the server.
2. **Map columns** — confirm or correct the auto-detected mapping between your column names and FalconBridge's standard fields.
3. **Results** — a scorecard, an expandable per-check table (14 checks), and a one-click Excel report download.
4. **Chat** — ask FalconBridge AI about the findings; it references specific records (Customer ID, Account Number) and only answers questions about this audit session.

## The 14 checks

**Customer Master:** CM-01 Missing Address, CM-02 Missing City, CM-03 Missing State, CM-04 Missing Zip, CM-05 Missing Phone.

**Transactions:** TX-01 Missed Payments, TX-02 Late Payments, TX-03 Orphan Transactions, TX-04 Duplicate Unique IDs, TX-05 Interest Rate Out of Range, TX-06 Invalid Amount Financed, TX-07 Monthly Payment Mismatch (amortization), TX-08 Payment Marked Yes But No Actual Date, TX-09 Additional Payment Amount.

## Deploy to Render

1. Push this repo to GitHub and create a new Blueprint in Render, linking `render.yaml`.
2. Set the `ANTHROPIC_API_KEY` environment variable in the Render dashboard for the `falconbridge-backend` service.
3. Point the frontend at the deployed backend (update the API base / proxy as needed for production; CORS is open for the POC and should be locked down to your frontend domain after deployment).

## Sample files

`Customer_Master.xlsx` and `Bank_Accounts.xlsx` can be used to exercise the full flow end to end.

## Tech stack

- **Frontend:** React + Vite, plain CSS, axios, SheetJS (`xlsx`) for in-browser validation.
- **Backend:** Python, FastAPI, pandas, openpyxl, uvicorn.
- **AI:** Anthropic Python SDK (`anthropic`), model `claude-sonnet-4-6`.
- **Report output:** openpyxl.
