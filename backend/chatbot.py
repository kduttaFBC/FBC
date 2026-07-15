"""
FalconBridge Audit Intelligence — AI chatbot backed by the Anthropic API.

Model: claude-sonnet-4-6 (per the FalconBridge spec).
Sends the full audit context as a system prompt and the entire conversation
history on every call (the Messages API is stateless).
"""

import os
import json

import anthropic
import pandas as pd

MODEL = "claude-haiku-4-5-20251001"
MAX_TOKENS = 1536
MAX_TOOL_ROUNDS = 3

LOOKUP_TOOL = {
    "name": "lookup_loan_history",
    "description": (
        "Looks up the full monthly payment/status history for ANY SBA GP Number "
        "directly from the complete SOUP Performance Dataset for this audit session — "
        "whether or not that loan was flagged by any check, and regardless of which "
        "servicer it belongs to. Use this whenever asked about a specific loan that "
        "isn't already covered in the pre-loaded flagged-loan history."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "sba_gp_number": {
                "type": "string",
                "description": "The SBA GP Number to look up, e.g. '2599617005'.",
            }
        },
        "required": ["sba_gp_number"],
    },
}

SOUP_LOOKUP_COLUMNS = [
    "SBA GP Number", "Servicer", "Payment Month", "Loan Status",
    "Delinquency Status", "Beginning Actual Balance", "Ending Actual Balance",
]


def _lookup_loan_history(soup_df, sba_gp_number):
    """Returns a CSV string of every SOUP row for the given account, or a
    clear not-found message — never silently returns nothing."""
    if soup_df is None:
        return "No live SOUP dataset is available for this session (it may have expired — try re-running the audit)."

    gp = str(sba_gp_number).strip()
    ids = soup_df["SBA GP Number"].astype("string").str.strip()
    mask = ids == gp
    matched = soup_df[mask]

    if matched.empty:
        return f"No records found for SBA GP Number '{gp}' in the SOUP Performance Dataset. Double-check the account number."

    matched = matched.copy()
    if "Payment Month" in matched.columns:
        matched = matched.sort_values("Payment Month")
    keep = [c for c in SOUP_LOOKUP_COLUMNS if c in matched.columns]
    if keep:
        matched = matched[keep]
    for col in matched.columns:
        if str(matched[col].dtype).startswith("datetime"):
            matched[col] = matched[col].astype("string")
    return matched.to_csv(index=False)

_client = None


def _get_client():
    global _client
    if _client is None:
        api_key = os.environ.get("ANTHROPIC_API_KEY") or os.getenv("ANTHROPIC_API_KEY")
        if not api_key:
            raise ValueError(f"ANTHROPIC_API_KEY not found. Env vars: {list(os.environ.keys())}")
        _client = anthropic.Anthropic(api_key=api_key)
    return _client


def _build_system_prompt(audit_context):
    """audit_context carries the structured audit results plus truncated CSV
    samples of SOUP and each servicer file. Keeping the flagged records front
    and centre keeps answers grounded and the token cost low."""
    summary = audit_context.get("summary", {})
    results = audit_context.get("results", [])
    soup_csv = audit_context.get("soup_csv", "")
    flagged_loan_history_csv = audit_context.get("flagged_loan_history_csv", "")
    servicer_csvs = audit_context.get("servicer_csvs", {})

    servicer_csv_block = "\n\n".join(
        f"--- {name} (Servicer Inputs sample) ---\n{csv}"
        for name, csv in servicer_csvs.items()
    )

    return (
        "You are FalconBridge AI, a data-quality and compliance assistant for SBA loan "
        "servicing. You have just completed a reconciliation between the SOUP Performance "
        "Dataset and one or more servicers' SRF (Servicer Remittance Form) submissions, "
        "matched on SBA GP Number.\n\n"
        "Answer ONLY questions about the data from this audit session. If asked "
        "anything outside the scope of this data, reply exactly: "
        '"I can only answer questions about the data from this audit session."\n\n'
        "When relevant, reference specific loans by SBA GP Number and name the servicer "
        "they belong to. Be concise and precise — this is a compliance context.\n\n"
        "IMPORTANT: for record/row counts, always use the exact figures in "
        "summary.servicer_record_counts and summary.soup_records — never estimate by "
        "counting rows in the CSV samples below, since those are truncated and will "
        "give a wrong answer.\n\n"
        "IMPORTANT: for questions about a specific loan's payment history over time "
        "(multiple months, trends, when a status changed, etc.), first check the "
        "'FLAGGED LOAN FULL PAYMENT HISTORY' section below — it already contains every "
        "month's SOUP record for every loan that appeared in this audit's flagged results. "
        "If the loan you need is NOT in that section (e.g. it wasn't flagged, or belongs to "
        "a servicer that wasn't part of this audit run), use the lookup_loan_history tool "
        "to query the complete SOUP Performance Dataset directly for that account — this "
        "works for ANY SBA GP Number, flagged or not. Only report a loan as 'not found' "
        "after actually calling the tool and getting a not-found result back — never assume "
        "a loan doesn't exist just because it's missing from the pre-loaded section.\n\n"
        "=== AUDIT SUMMARY (JSON) ===\n"
        f"{json.dumps(summary, indent=2)}\n\n"
        "=== AUDIT RESULTS (JSON) ===\n"
        f"{json.dumps(results, indent=2, default=str)}\n\n"
        "=== FLAGGED LOAN FULL PAYMENT HISTORY (CSV, every month, every flagged loan) ===\n"
        f"{flagged_loan_history_csv}\n\n"
        "=== SOUP PERFORMANCE DATASET SAMPLE (CSV, small generic snapshot, truncated) ===\n"
        f"{soup_csv}\n\n"
        "=== SERVICER FILE SAMPLES (CSV, truncated) ===\n"
        f"{servicer_csv_block}\n"
    )


def chat(message, history, audit_context, soup_df=None):
    """
    message: the new user message (str)
    history: list of {"role": "user"|"assistant", "content": str}
    audit_context: dict with summary, results, and truncated CSV samples
    soup_df: the full SOUP DataFrame for this session, if still cached
             server-side — enables the lookup_loan_history tool for ANY
             account, not just pre-loaded flagged loans.
    Returns the assistant reply (str).
    """
    system_prompt = _build_system_prompt(audit_context or {})

    messages = []
    for turn in history or []:
        role = turn.get("role")
        content = turn.get("content", "")
        if role in ("user", "assistant") and content:
            messages.append({"role": role, "content": content})
    messages.append({"role": "user", "content": message})

    client = _get_client()
    tools = [LOOKUP_TOOL] if soup_df is not None else []

    for _ in range(MAX_TOOL_ROUNDS):
        response = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=system_prompt,
            messages=messages,
            tools=tools,
        )

        if response.stop_reason != "tool_use":
            return "".join(block.text for block in response.content if block.type == "text")

        # Claude wants to call a tool — execute it and feed the result back
        messages.append({"role": "assistant", "content": response.content})
        tool_results = []
        for block in response.content:
            if block.type != "tool_use":
                continue
            if block.name == "lookup_loan_history":
                gp = block.input.get("sba_gp_number", "")
                result_text = _lookup_loan_history(soup_df, gp)
            else:
                result_text = f"Unknown tool: {block.name}"
            tool_results.append({
                "type": "tool_result",
                "tool_use_id": block.id,
                "content": result_text,
            })
        messages.append({"role": "user", "content": tool_results})

    # Fallback if it somehow keeps calling tools past the round limit
    return "I looked into this but couldn't finish putting together an answer. Please try rephrasing your question."
