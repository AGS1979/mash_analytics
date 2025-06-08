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
import html as html_lib


DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_CHAT_URL = os.getenv("DEEPSEEK_CHAT_URL", "https://api.deepseek.com/v1/chat/completions")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_CHAT_URL = "https://api.openai.com/v1/chat/completions"
FMP_API_KEY = os.getenv("FMP_API_KEY", "")
FMP_BASE_URL = "https://financialmodelingprep.com/api/v3"
HEADERS_FMP = {"User-Agent": "Avinash Singh <avinashg.singh@aranca.com>"}


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


def extract_financial_data(text, line_item_queries, ticker=None):
    try:
        print("📊 [STEP 4] Extracting financial data using OpenAI...")

        line_items = [item.strip() for item in line_item_queries if item.strip()]
        print("🧾 Line items requested:", line_items)

        prompt = (
            "You are a financial data analyst. Your task is to extract historical actual financial values "
            "for the requested line items from the given company filing excerpts.\n"
            "Use only the actual reported figures — avoid forecasts, guidance, or narrative commentary.\n"
            "Extract values for the past 3–5 years, if available.\n"
            "Return the data in **JSON format only**, with this structure:\n\n"
            "{\n"
            '  "revenue": {"2024": 12345, "2023": 11800},\n'
            '  "net_income": {"2024": 3400, "2023": 2900},\n'
            '  "cash": {"2024": 2100},\n'
            '  "debt": {"2024": 5000},\n'
            '  "diluted_eps": {"2024": 13.63}\n'
            "}\n\n"
            "Use **lowercase snake_case** keys. Return figures in **millions of USD**, unless otherwise stated.\n"
            "Do NOT include any text, commentary, or explanation outside the JSON.\n\n"
            f"Requested Items:\n- " + "\n- ".join(line_items) + "\n\n"
            f"Company Filing Excerpt (first 8000 characters):\n{text[:8000]}"
        )

        response = call_llm(prompt, provider="openai", max_tokens=1600)
        print("📥 Raw OpenAI response (first 1000 chars):\n", response[:1000])
        cleaned = re.sub(r"```json|```", "", response).strip()
        json_match = re.search(r"{.*}", cleaned, re.DOTALL)
        if not json_match:
            print("❌ No valid JSON object found in OpenAI response.")
            return {}

        raw_json = json_match.group(0)
        parsed_data = json.loads(raw_json)
        print("✅ Parsed JSON from OpenAI.")

        # Synonym map
        synonyms = {
            "total_cash": "cash",
            "cash_and_cash_equivalents": "cash",
            "cashandshortterminvestments": "cash",
            "available_cash": "cash",
            "total_debt": "debt",
            "borrowings": "debt",
            "eps": "diluted_eps",
            "basic_eps": "diluted_eps",
        }

        # Normalize keys
        normalized = {}
        for k in list(parsed_data.keys()):
            norm_k = k.lower().strip().replace(" ", "_")
            key_to_use = synonyms.get(norm_k, norm_k)
            val = parsed_data[k]
            if isinstance(val, dict):
                normalized.setdefault(key_to_use, {}).update(val)

        # Check if we need fallback
        required = ["cash", "debt"]
        missing = [k for k in required if k not in normalized]
        if missing and ticker and "." not in ticker:
            print(f"⚠️ Missing fields: {missing} — attempting FMP fallback for {ticker}")

            def fetch_fmp(endpoint):
                url = f"{FMP_BASE_URL}/{endpoint}/{ticker}?limit=1&apikey={FMP_API_KEY}"
                res = requests.get(url, headers=HEADERS_FMP)
                return res.json()[0] if res.ok and res.json() else {}

            bs = fetch_fmp("balance-sheet-statement")
            is_ = fetch_fmp("income-statement")
            print("🧾 Raw FMP balance sheet:", bs)
            print("🧾 Raw FMP income statement:", is_)

            if "cash" not in normalized and "cashAndShortTermInvestments" in bs:
                normalized["cash"] = {"2024": round(bs["cashAndShortTermInvestments"] / 1e6, 2)}
            if "debt" not in normalized:
                total_debt = (bs.get("shortTermDebt", 0) or 0) + (bs.get("longTermDebt", 0) or 0)
                if total_debt > 0:
                    normalized["debt"] = {"2024": round(total_debt / 1e6, 2)}
            if "diluted_eps" not in normalized and "eps" in is_:
                normalized["diluted_eps"] = {"2024": round(is_["eps"], 2)}
            if "net_income" not in normalized and "netIncome" in is_:
                normalized["net_income"] = {"2024": round(is_["netIncome"] / 1e6, 2)}

        print("✅ Final normalized financials:\n", json.dumps(normalized, indent=2))
        return normalized

    except Exception as e:
        print(f"❌ Extraction error: {e}")
        return {}


def generate_forecast_scenarios(text, financials, assumptions):
    prompt = (
        "Based on the following financial data and company document excerpts, generate 3 forecast scenarios "
        "(bull, base, and bear) for Free Cash Flow to Firm (FCFF) over the next "
        f"{assumptions.get('forecast_years', 5)} years.\n"
        "Return your output strictly in **JSON format only** with the following structure:\n\n"
        "{\n"
        '  "bull": {\n'
        '    "fcff": [3600, 4000, 4400, 4800, 5200],\n'
        '    "justification": "Brief explanation here."\n'
        "  },\n"
        '  "base": {\n'
        '    "fcff": [3400, 3700, 4000, 4300, 4600],\n'
        '    "justification": "Brief explanation here."\n'
        "  },\n"
        '  "bear": {\n'
        '    "fcff": [3000, 3200, 3400, 3600, 3800],\n'
        '    "justification": "Brief explanation here."\n'
        "}\n\n"
        "Use integers in millions of USD. Do not include any markdown, bullet points, or non-JSON text.\n\n"
        f"FINANCIALS:\n{json.dumps(financials)}\n\nTEXT:\n{text[:10000]}"
    )

    return call_llm(prompt, provider="openai")


def calculate_dcf_scenarios(forecast_json, assumptions, cash, debt, cmp, financials=None):
    try:
        if isinstance(forecast_json, str):
            forecast_json = forecast_json.strip()

            # Clean markdown-style code block if present
            if forecast_json.startswith("```json"):
                forecast_json = forecast_json[len("```json"):].strip()
            elif forecast_json.startswith("```"):
                forecast_json = forecast_json[len("```"):].strip()
            if forecast_json.endswith("```"):
                forecast_json = forecast_json[:-3].strip()

            forecast_json = json.loads(forecast_json)

    except Exception as e:
        print("🔥 [FATAL] Failed to parse forecast JSON:", str(e))
        print("🧾 Forecast response was:\n", forecast_json)
        raise

    # Extract WACC and terminal growth assumptions
    wacc = float(assumptions["WACC"])
    terminal_growth = float(assumptions["terminal_rate_or_multiple"])
    years = int(assumptions["forecast_years"])

    # Discounting helper
    def discount_fcffs(fcffs):
        return sum(f / ((1 + wacc / 100) ** (i + 1)) for i, f in enumerate(fcffs))

    # Compute shares = net_income / EPS
    try:
        if not financials:
            raise ValueError("Missing `financials` for computing shares")

        ni_years = financials.get("net_income", {})
        eps_years = financials.get("diluted_eps", {})
        common_years = set(ni_years.keys()) & set(eps_years.keys())

        if not common_years:
            raise ValueError("No common year between net income and EPS")

        latest_year = max(common_years)
        net_income = financials["net_income"][latest_year]
        eps = financials["diluted_eps"][latest_year]

        shares = round(net_income / eps, 2)
        print(f"🧠 Computed shares = {net_income} / {eps} = {shares}")
    except Exception as e:
        print(f"❌ Could not compute shares from EPS and Net Income: {e}")
        raise

    # Build DCF output
    output = {}
    for scenario in ["bull", "base", "bear"]:
        fcffs = forecast_json[scenario]["fcff"]
        terminal_fcff = fcffs[-1] * (1 + terminal_growth / 100) / ((wacc - terminal_growth) / 100)
        dcf_value = discount_fcffs(fcffs) + terminal_fcff / ((1 + wacc / 100) ** years)
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
    # Ensure cmp is a scalar float (not a Series)
    if isinstance(cmp, (pd.Series, np.ndarray)):
        cmp_value = float(cmp.dropna().iloc[-1])
    else:
        cmp_value = float(cmp)

    html = f"<h2>📊 DCF Valuation for {ticker}</h2><p>Current Market Price: <strong>${cmp_value}</strong></p>"
    html += "<table border='1' cellpadding='8' cellspacing='0'><tr><th>Scenario</th><th>Fair Value</th><th>Upside</th><th>Justification</th></tr>"

    for k, v in dcf_result.items():
        upside = round((v['fair_value'] - cmp_value) / cmp_value * 100, 2)
        justification = html_lib.escape(v['justification'])
        html += f"<tr><td>{k.title()}</td><td>${v['fair_value']}</td><td>{upside}%</td><td>{justification}</td></tr>"


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
