import os
import re
import json
import time
import requests
import tempfile

import pandas as pd
import numpy as np
import yfinance as yf

from PyPDF2 import PdfReader
from docx import Document
from pptx import Presentation


DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_CHAT_URL = os.getenv("DEEPSEEK_CHAT_URL", "https://api.deepseek.com/v1/chat/completions")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_CHAT_URL = "https://api.openai.com/v1/chat/completions"

HEADERS_DEEPSEEK = {
    "Content-Type": "application/json",
    "Authorization": f"Bearer {DEEPSEEK_API_KEY}"
}

HEADERS_OPENAI = {
    "Content-Type": "application/json",
    "Authorization": f"Bearer {OPENAI_API_KEY}"
}


def get_ticker(company_name: str) -> str:
    """
    Extract the most likely ticker from LLM response and validate it via yfinance.
    """
    prompt = (
      f"Please provide only the stock ticker symbol (e.g., GD, AAPL) for the company named '{company_name}', "
      "with no explanation or extra text."
    )

    response = call_llm(prompt)
    print("🔍 Raw LLM response for ticker:", response)

    # Extract possible tickers (uppercase 1–5 letters optionally with .AX etc.)
    matches = re.findall(r"\b[A-Z]{1,5}(?:\.[A-Z]{1,2})?\b", response)

    tried = []
    for candidate in matches:
        tried.append(candidate)
        try:
            df = yf.download(candidate, period="1d", progress=False)
            if not df.empty and "Close" in df.columns:
                print(f"✅ Validated ticker: {candidate}")
                return candidate
        except Exception as e:
            print(f"❌ {candidate} failed with error: {e}")
            continue

    raise ValueError(f"❌ No valid ticker found for company: {company_name}. Tried: {tried}")



def get_current_price(ticker: str) -> float:

    try:
        print(f"📈 Fetching current price for ticker: {ticker}")
        df = yf.download(ticker, period="5d", interval="1d", progress=False)

        if df.empty or "Close" not in df.columns:
            print("❌ yfinance returned empty data or missing 'Close' column.")
            raise ValueError(f"No valid price data found for ticker {ticker}.")

        latest_close = df["Close"].dropna().iloc[-1]
        print(f"💲 Latest closing price for {ticker}: {latest_close}")
        return round(latest_close, 2)

    except Exception as e:
        print(f"❌ Failed to fetch price for {ticker}: {e}")
        raise ValueError("-1")  # Let the calling function handle it


def call_llm(prompt: str, temperature=0.2, max_tokens=1000, provider="deepseek") -> str:
    body = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": temperature,
        "max_tokens": max_tokens
    } if provider == "deepseek" else {
        "model": "gpt-4o",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": temperature,
        "max_tokens": max_tokens
    }

    headers = HEADERS_DEEPSEEK if provider == "deepseek" else HEADERS_OPENAI
    url = DEEPSEEK_CHAT_URL if provider == "deepseek" else OPENAI_CHAT_URL
    response = requests.post(url, headers=headers, data=json.dumps(body))
    return response.json()["choices"][0]["message"]["content"]


def extract_text_from_documents(filepaths):
    text = ""
    for path in filepaths:
        if path.endswith(".pdf"):
            with open(path, "rb") as f:
                reader = PdfReader(f)
                for page in reader.pages:
                    text += page.extract_text() or ""
        elif path.endswith(".docx"):
            doc = Document(path)
            text += "\n".join(p.text for p in doc.paragraphs)
        elif path.endswith(".pptx"):
            prs = Presentation(path)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        text += shape.text + "\n"
    return text


def extract_financial_data(text, line_items):
    prompt = (
        "Extract the following line items from the financial documents text below:\n"
        f"{', '.join(line_items)}\n"
        "Return results in JSON format with line item as key and year-wise values as nested dictionary."
        "\n\nTEXT:\n" + text[:8000]
    )
    response = call_llm(prompt, provider="deepseek")
    try:
        return json.loads(response)
    except:
        return {}


def generate_forecast_scenarios(text, financials, assumptions):
    prompt = (
        "Based on the following financial data and document excerpts, generate 3 financial scenarios "
        "(bull, base, bear) of Free Cash Flow to Firm (FCFF) over the next "
        f"{assumptions.get('forecast_years', 5)} years.\n"
        "You should consider trends, segment outlooks, industry guidance, and recent news.\n"
        "Output should include FCFF per year and a one-line justification per case.\n"
        f"FINANCIALS:\n{json.dumps(financials)}\n\nTEXT:\n{text[:10000]}"
    )
    return call_llm(prompt, provider="openai")


def calculate_dcf_scenarios(forecast_json, assumptions, cash, debt, shares, cmp):
    if isinstance(forecast_json, str):
        forecast_json = json.loads(forecast_json)

    wacc = float(assumptions["WACC"])
    terminal_growth = float(assumptions["terminal_rate_or_multiple"])
    years = int(assumptions["forecast_years"])

    def discount_fcffs(fcffs):
        return sum(f / ((1 + wacc/100) ** (i+1)) for i, f in enumerate(fcffs))

    output = {}
    for scenario in ["bull", "base", "bear"]:
        fcffs = forecast_json[scenario]["fcff"]
        terminal_fcff = fcffs[-1] * (1 + terminal_growth/100) / ((wacc - terminal_growth)/100)
        dcf_value = discount_fcffs(fcffs) + terminal_fcff / ((1 + wacc/100) ** years)
        equity_value = dcf_value + cash - debt
        fair_value = equity_value / shares
        output[scenario] = {
            "fcff": fcffs,
            "dcf_value": round(dcf_value, 2),
            "equity_value": round(equity_value, 2),
            "fair_value": round(fair_value, 2),
            "cmp": cmp,
            "justification": forecast_json[scenario].get("justification", "")
        }
    return output


def format_html_output(dcf_result, financials, ticker, cmp):
    html = f"<h2>📊 DCF Valuation for {ticker}</h2><p>Current Market Price: <strong>${cmp}</strong></p><table border='1' cellpadding='8' cellspacing='0'><tr><th>Scenario</th><th>Fair Value</th><th>Upside</th><th>Justification</th></tr>"
    for k, v in dcf_result.items():
        upside = round((v['fair_value'] - cmp)/cmp * 100, 2)
        html += f"<tr><td>{k.title()}</td><td>${v['fair_value']}</td><td>{upside}%</td><td>{v['justification']}</td></tr>"
    html += "</table>"
    return html


def generate_excel_output(dcf_result):
    from io import BytesIO
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    df = pd.DataFrame(dcf_result).T
    df.to_excel(writer, sheet_name="DCF Scenarios")
    writer.close()
    output.seek(0)
    return output.read()
