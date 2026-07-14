"""
FalconBridge Audit Intelligence — AI chatbot backed by the Anthropic API.

Model: claude-sonnet-4-6 (per the FalconBridge spec).
Sends the full audit context as a system prompt and the entire conversation
history on every call (the Messages API is stateless).
"""

import os
import json

import anthropic

MODEL = "claude-haiku-4-5-20251001"
MAX_TOKENS = 1024

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
        "(multiple months, trends, when a status changed, etc.), use the "
        "'FLAGGED LOAN FULL PAYMENT HISTORY' section below — it contains every month's "
        "SOUP record for every loan that appeared in this audit's flagged results. The "
        "generic SOUP sample further below is just a small snapshot of the whole file and "
        "will NOT reliably contain a specific loan's full history.\n\n"
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


def chat(message, history, audit_context):
    """
    message: the new user message (str)
    history: list of {"role": "user"|"assistant", "content": str}
    audit_context: dict with summary, results, and truncated CSV samples
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
    response = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=system_prompt,
        messages=messages,
    )

    return "".join(block.text for block in response.content if block.type == "text")
