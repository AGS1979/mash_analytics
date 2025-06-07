import os
import re
import time
import json
import requests
import traceback
import yfinance as yf
import pandas as pd

from PyPDF2 import PdfReader
from docx import Document

# ─────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_CHAT_URL = os.getenv("DEEPSEEK_CHAT_URL", "https://api.deepseek.com/v1/chat/completions")

if not DEEPSEEK_API_KEY:
    raise RuntimeError("Please set DEEPSEEK_API_KEY in your environment.")

HEADERS = {
    "Content-Type": "application/json",
    "Authorization": f"Bearer {DEEPSEEK_API_KEY}"
}

# ─────────────────────────────────────────────────────────────
# TICKER + PRICE
# ─────────────────────────────────────────────────────────────
def get_ticker_from_name(company_name):
    prompt = f"""Return only the stock ticker for this public company in Yahoo Finance format:\n{company_name}\nExample: AAPL"""
    resp = requests.post(DEEPSEEK_CHAT_URL, headers=HEADERS, json={
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}]
    })
    return resp.json()["choices"][0]["message"]["content"].strip()

def get_current_share_price(ticker):
    try:
        data = yf.Ticker(ticker).history(period="1d")
        return round(data['Close'].iloc[-1], 2)
    except Exception:
        return None

# ─────────────────────────────────────────────────────────────
# FILE HANDLING
# ─────────────────────────────────────────────────────────────
def save_uploaded_files(files):
    paths = []
    for f in files:
        path = f"/tmp/{int(time.time() * 1000)}_{f.filename}"
        f.save(path)
        paths.append(path)
    return paths

def extract_text_from_documents(paths):
    content = {}
    for path in paths:
        try:
            if path.endswith(".pdf"):
                reader = PdfReader(path)
                text = "\n".join([page.extract_text() or "" for page in reader.pages])
            elif path.endswith(".docx"):
                doc = Document(path)
                text = "\n".join([para.text for para in doc.paragraphs])
            else:
                text = ""
        except:
            text = ""
        content[os.path.basename(path)] = text.strip()
    return content

# ─────────────────────────────────────────────────────────────
# LLM WRAPPERS
# ─────────────────────────────────────────────────────────────
def call_deepseek(prompt, temperature=0.2, max_tokens=2000):
    body = {
        "model": "deepseek-chat",
        "temperature": temperature,
        "max_tokens": max_tokens,
        "messages": [{"role": "user", "content": prompt}]
    }
    resp = requests.post(DEEPSEEK_CHAT_URL, headers=HEADERS, json=body)
    if resp.status_code != 200:
        raise RuntimeError(f"DeepSeek API error {resp.status_code}: {resp.text}")
    return resp.json()['choices'][0]['message']['content'].strip()

# ─────────────────────────────────────────────────────────────
# LINE ITEM + DCF EXTRACTION
# ─────────────────────────────────────────────────────────────
def extract_line_items(prompt, extracted_text):
    full_text = "\n".join(extracted_text.values())[:18000]
    q = f"""
You are a financial analyst. Extract only the requested historical financial line items using exact values from the provided text.

If something is not present, write "Not Found".

Line Items:
{prompt}

Raw Filing Text:
{full_text}
"""
    return call_deepseek(q)

def resolve_assumptions(source, user_assumptions, extracted_text):
    if source == "own":
        return user_assumptions

    elif source == "management":
        prompt = "Extract WACC, terminal growth or terminal multiple, and forecast period from the company's filings."
        _ = extract_line_items(prompt, extracted_text)  # fallback to user if DeepSeek fails
        return {
            "WACC": user_assumptions.get("WACC") or "9.0",
            "terminal_rate_or_multiple": user_assumptions.get("terminal_rate_or_multiple") or "2.5",
            "model_type": user_assumptions.get("model_type") or "perpetuity",
            "forecast_years": user_assumptions.get("forecast_years") or "5"
        }

    else:
        return {
            "WACC": "9.0",
            "terminal_rate_or_multiple": "2.5",
            "model_type": "perpetuity",
            "forecast_years": "5"
        }

def run_dcf_model(line_item_block, assumptions, cmp):
    prompt = f"""
You are a valuation analyst. Use the historical line items and DCF assumptions below to compute a 3-scenario DCF valuation (Bull, Base, Bear).

---
Line Items:
{line_item_block}

Assumptions:
WACC: {assumptions.get("WACC")}%
Forecast Years: {assumptions.get("forecast_years")}
Terminal Value Method: {assumptions.get("model_type")}
Terminal Rate or Multiple: {assumptions.get("terminal_rate_or_multiple")}
Current Share Price: ${cmp}
---

Return the output as a markdown-formatted table with these fields:
- Scenario
- Enterprise Value
- Net Debt
- Equity Value
- Shares Outstanding
- Fair Value Per Share
- Upside or Downside vs Current Price

Keep it professional. Avoid casual language. Do not offer suggestions.
"""
    return call_deepseek(prompt, temperature=0.1, max_tokens=1800)
