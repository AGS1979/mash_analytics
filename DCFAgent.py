import os
import re
import json
import time
import requests
import tempfile

import pandas as pd
import numpy as np
import yfinance as yf

import pdfplumber
from docx import Document
from pptx import Presentation
from openpyxl import Workbook
# ─────────────────────────────────────────
# CONFIG: DeepSeek API
# ─────────────────────────────────────────
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_CHAT_URL = "https://api.deepseek.com/v1/chat/completions"


if not DEEPSEEK_API_KEY:
    raise RuntimeError("❌ DEEPSEEK_API_KEY is not set in the environment.")


# ─────────────────────────────────────────
# DEEPSEEK CHAT CALL
# ─────────────────────────────────────────
def call_deepseek(prompt, temperature=0.2, max_tokens=1000):
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}"
    }
    body = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": temperature,
        "max_tokens": max_tokens
    }
    response = requests.post(DEEPSEEK_CHAT_URL, headers=headers, json=body)
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"].strip()


# ─────────────────────────────────────────
# 1. Get Stock Ticker from Name
# ─────────────────────────────────────────
def get_ticker_from_name(company_name: str) -> str:
    prompt = f"What is the primary US stock ticker for {company_name}? Only respond with the ticker symbol. For example: 'RTX'"
    response = call_deepseek(prompt)

    # Fallback regex extraction in case LLM still responds with extra text
    match = re.search(r'\b[A-Z]{1,5}\b', response)
    if match:
        return match.group(0)
    raise ValueError(f"Could not extract a valid ticker from: {response}")


def get_current_share_price(ticker: str) -> float:
    try:
        data = yf.Ticker(ticker).history(period="1d")
        return round(data['Close'][-1], 2)
    except Exception as e:
        raise RuntimeError(f"Failed to fetch current share price for {ticker}: {e}")


# ─────────────────────────────────────────
# 2. Save Uploaded Files
# ─────────────────────────────────────────
def save_uploaded_files(uploaded_files):
    saved_paths = []
    for file in uploaded_files:
        suffix = os.path.splitext(file.filename)[-1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            file.save(tmp.name)
            saved_paths.append(tmp.name)
    return saved_paths


# ─────────────────────────────────────────
# 3. Extract Text from PDF, DOCX, PPTX
# ─────────────────────────────────────────
def extract_text_from_documents(file_paths):
    extracted = {}
    for path in file_paths:
        if path.endswith(".pdf"):
            with pdfplumber.open(path) as pdf:
                text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        elif path.endswith(".docx"):
            doc = Document(path)
            text = "\n".join([para.text for para in doc.paragraphs])
        elif path.endswith(".pptx"):
            prs = Presentation(path)
            text = "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
        else:
            text = ""
        extracted[path] = text
    return extracted


# ─────────────────────────────────────────
# 4. Extract Financial Line Items
# ─────────────────────────────────────────
def extract_line_items_from_text(text_dict, line_items):
    results = {}
    for item in line_items:
        combined_text = " ".join(text_dict.values())[:16000]
        prompt = f"""Extract historical values for the following line item: {item}.
If available, give for the last 3–5 years. Format:
Line Item | Year | Value

Avoid follow-up suggestions like "Let me know if you'd like help" or "You may refer to other sources." Only state what is available in the text.

Text:
{combined_text}"""
        try:
            results[item] = call_deepseek(prompt)
        except Exception as e:
            results[item] = f"Error: {e}"
    return results


# ─────────────────────────────────────────
# 5. Resolve Assumptions
# ─────────────────────────────────────────
def resolve_assumptions(source, user_inputs, extracted_text):
    if source == "own":
        required = ["WACC", "terminal_rate_or_multiple", "model_type", "forecast_years"]
        if not all(k in user_inputs and user_inputs[k] for k in required):
            raise ValueError("Missing fields in own assumptions.")
        return user_inputs

    elif source == "management":
        prompt = """
    From the following management commentary, extract forward-looking guidance that can inform a DCF model. Prioritize any references to:
    - Free Cash Flow (FCF or FCFF) targets for upcoming years
    - Revenue or EBITDA guidance
    - Operating margin targets
    - Capital expenditure plans
    - Growth assumptions
    - Terminal value guidance (exit multiples or perpetuity growth)
    - Forecast horizon
    - Management's expected macro assumptions (e.g., interest rates, inflation, defense budgets)

    Respond in a structured format like:

    Free Cash Flow Forecast:
    Year | FCFF ($B)
    2024 | X.X
    2025 | X.X
    ...

    Other DCF Inputs:
    - WACC estimate (if any): X%
    - Terminal growth (if any): X%
    - Forecast period: N years
    - Capex trend: ...
    - Revenue CAGR: ...

    If guidance is only qualitative (e.g., "double FCF by 2026"), provide a narrative but attempt to estimate figures reasonably.

    Here is the extracted text:
    """ + " ".join(extracted_text.values())[:16000]

        try:
            guidance = call_deepseek(prompt)
            if any(kw in guidance.lower() for kw in ["not found", "no guidance", "unknown"]):
                return None  # fallback
            return {"source": "mgmt", "response": guidance}
        except Exception as e:
            return None


    elif source == "llm":
        prompt = "Suggest reasonable DCF assumptions (WACC, terminal growth/multiple, forecast period) for a typical public company in its industry."
        return {"source": "llm", "response": call_deepseek(prompt)}

    else:
        raise ValueError("Invalid assumption source.")


# ─────────────────────────────────────────
# 6. Stubbed DCF Model
# ─────────────────────────────────────────
def run_dcf_model(line_item_data: dict, assumptions: dict, cmp: float) -> dict:
    """
    Sends extracted financial data and assumptions to DeepSeek to compute DCF values for Bull, Base, Bear cases.
    """
    # Format the extracted financials into a string block
    financials_block = ""
    for item, raw in line_item_data.items():
        financials_block += f"\n# {item}\n{raw.strip()}\n"

    # Format the assumptions
    assumption_text = ""
    if "source" in assumptions and "response" in assumptions:
        assumption_text = assumptions["response"]
    elif all(k in assumptions for k in ["WACC", "terminal_rate_or_multiple", "model_type", "forecast_years"]):
        assumption_text = f"""
WACC: {assumptions['WACC']}
Forecast Period: {assumptions['forecast_years']} years
DCF Type: {assumptions['model_type']}
Terminal {"Growth Rate" if assumptions['model_type'] == "perpetuity" else "Exit Multiple"}: {assumptions['terminal_rate_or_multiple']}
"""
    else:
        raise ValueError("Invalid or incomplete assumptions.")

    # Compose final prompt
        # Add current price to prompt
    cmp_text = f"The current market price (CMP) of the stock is approximately ${cmp} as of today.\n"

    prompt = f"""
Using the following extracted financials, assumptions, and current share price, generate a 3-scenario Discounted Cash Flow (DCF) analysis:
- Scenarios: Bull, Base, Bear
- Each should include: Enterprise Value (EV), Equity Value, and Fair Value Per Share.
- Ensure all scenarios are reasonable in light of the current share price (${cmp}).
- Avoid assigning downside in the Bull Case unless highly justified.

## Current Share Price
{cmp_text}

## Financials
{financials_block}

## Assumptions
{assumption_text}

Format output like:
Scenario | Enterprise Value | Equity Value | Per Share Value
Bull Case | 150000 | 130000 | 95
...
"""


# Call DeepSeek
try:
    response = call_deepseek(prompt)
    print("\n🔎 Raw DeepSeek DCF Response:\n", response)

    parsed = {}
    for line in response.splitlines():
        if "|" in line and "Scenario" not in line and not line.strip().startswith("---"):
            parts = [re.sub(r'[\*\$]', '', p.strip()) for p in line.split("|")]
            if len(parts) == 4:
                try:
                    scenario = parts[0]
                    ev = float(parts[1].replace(",", "").replace("B", "e9").replace("M", "e6"))
                    eqv = float(parts[2].replace(",", "").replace("B", "e9").replace("M", "e6"))
                    ps = float(parts[3].replace(",", "").replace("B", "e9").replace("M", "e6"))
                    parsed[scenario] = {
                        "EV": ev,
                        "Equity Value": eqv,
                        "Per Share": ps
                    }
                except ValueError:
                    continue

    if not parsed:
        raise ValueError(f"DeepSeek returned no valid output. Response:\n{response}")
    return parsed

except Exception as e:
    raise RuntimeError(f"LLM-based DCF calculation failed: {e}")
