import os
import re
import json
import time
import requests

import pandas as pd
import numpy as np
import yfinance as yf

from PyPDF2 import PdfReader
import camelot
from openpyxl import Workbook
from docx import Document
from bs4 import BeautifulSoup

from InvMemo import PDFQueryEngine  # ✅ Correct import based on your file name


OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_CHAT_URL = "https://api.openai.com/v1/chat/completions"
HEADERS_OPENAI = {
    "Content-Type": "application/json",
    "Authorization": f"Bearer {OPENAI_API_KEY}"
}


def get_ticker(company_name: str) -> str:
    prompt = (
        f"Please provide only the stock ticker symbol (e.g., GD, AAPL) for the company named '{company_name}', "
        "with no explanation or extra text."
    )
    response = call_llm(prompt)
    print("🔍 Raw LLM response for ticker:", response)

    matches = re.findall(r"\b[A-Z]{1,5}(?:\.[A-Z]{1,2})?\b", response)
    for candidate in matches:
        try:
            df = yf.download(candidate, period="1d", progress=False)
            if not df.empty and "Close" in df.columns:
                print(f"✅ Validated ticker: {candidate}")
                return candidate
        except Exception as e:
            print(f"❌ {candidate} failed with error: {e}")
            continue
    raise ValueError(f"❌ No valid ticker found for company: {company_name}")


def get_current_price(ticker: str) -> float:
    df = yf.download(ticker, period="5d", interval="1d", progress=False)
    if df.empty or "Close" not in df.columns:
        raise ValueError(f"No valid price data found for ticker {ticker}.")
    return round(df["Close"].dropna().iloc[-1], 2)


def extract_year_value(field, year):
    return field.get(str(year), 0) if isinstance(field, dict) else field


def call_llm(prompt: str, temperature=0.2, max_tokens=1000) -> str:
    body = {
        "model": "gpt-4o",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": temperature,
        "max_tokens": max_tokens
    }
    response = requests.post(OPENAI_CHAT_URL, headers=HEADERS_OPENAI, data=json.dumps(body))
    return response.json()["choices"][0]["message"]["content"]


def extract_financials_with_pdfquery(filepaths):
    import re
    from InvMemo import PDFQueryEngine
    from types import MethodType

    def extract_year_value_from_text(answer, target_year=2024):
        pattern = re.findall(r"(\d{4})[:\s\-]+([-+]?\d[\d,\.]*)", answer)
        data = {int(yr): float(val.replace(",", "")) for yr, val in pattern}
        return round(data.get(target_year, 0), 2)

    engine = PDFQueryEngine()
    engine.chunks = []
    
    for path in filepaths:
        engine.chunks.extend(engine.extract_text_from_pdf(path))
    
    engine.chunk_texts = [text for _, text in engine.chunks]
    engine.embeddings = engine.embed_texts(engine.chunk_texts)
    engine.index = engine.build_faiss_index(engine.embeddings)

    def answer_query_bound(self, query: str) -> str:
        query_embedding = self.embedder.encode([query], convert_to_numpy=True)
        distances, indices = self.index.search(query_embedding, 5)

        keywords = ["cash", "debt", "borrowings", "shares", "equity", "capital", "EPS", "revenue", "income"]
        retrieved_chunks = [
            chunk for i in indices[0] for chunk in [self.chunk_texts[i]]
            if any(k in chunk.lower() for k in keywords)
        ]
        if not retrieved_chunks:
            retrieved_chunks = [self.chunk_texts[i] for i in indices[0]]

        context = "\n\n".join(retrieved_chunks)

        prompt = {
            "model": "gpt-4o",
            "messages": [
                {"role": "system", "content": "You are a financial analyst."},
                {"role": "user", "content": f"Based only on the context below, answer this: {query}\n\nContext:\n{context}"}
            ],
            "temperature": 0,
            "max_tokens": 300
        }

        response = requests.post(
            OPENAI_CHAT_URL,
            headers=HEADERS_OPENAI,
            data=json.dumps(prompt)
        )

        return response.json()["choices"][0]["message"]["content"].strip()

    engine.answer_query = MethodType(answer_query_bound, engine)

    print("📄 Using PDFQueryEngine for financial extraction...")

    QUERIES = {
        "cash": [
            "What was the total cash or cash equivalents for the company in 2024? Provide only the number.",
            "Only return total cash (or cash and equivalents) for 2024 from the context. Return only the number."
        ],
        "debt": [
            "What is the total debt (or borrowings) for 2024? Just return the number.",
            "Give the total debt or borrowings for the company in 2024. Return only the number."
        ],
        "shares": [
            "How many diluted shares outstanding were there in 2024? Only return the number.",
            "Provide diluted shares outstanding for 2024. Return just the number."
        ],
        "net_income": ["Only provide net income for 2024, 2023, 2022, 2021, 2020. Only return a number."],
        "diluted_eps": ["Only provide diluted earnings per share (EPS) for 2024. Only return a number."],
        "revenue": ["Only provide total revenue for the company for 2024, 2023, 2022, 2021, 2020. Only return a number."],
        "operating_income": ["Only provide net income for 2024, 2023, 2022, 2021, 2020. Only return a number."],
        "EBITDA": ["Only provide EBITDA for 2024, 2023, 2022, 2021, 2020. Only return a number."],
        "segment_revenue": ["Only provide segment revenue for 2024, 2023, 2022, 2021, 2020. Only return a number."],
        "segment_EBITDA": ["Only provide segment EBITDA for 2024, 2023, 2022, 2021, 2020. Only return a number."],
        "free_cash_flow": ["Only provide free cash flow for 2024, 2023, 2022, 2021, 2020. Only return a number."],
        "capex": ["Only provide capital expenditure for 2024, 2023, 2022, 2021, 2020. Only return a number."]
    }

    MULTI_YEAR_KEYS = {
        "net_income", "operating_income", "EBITDA", "segment_revenue",
        "segment_EBITDA", "free_cash_flow", "capex"
    }

    results = {}
    for key, prompts in QUERIES.items():
        for prompt in prompts:
            try:
                answer = engine.answer_query(prompt)
                print(f"📌 {key} prompt: {prompt} → {answer}")

                if key in MULTI_YEAR_KEYS:
                    val = extract_year_value_from_text(answer, target_year=2024)
                    if val:
                        results[key] = val
                        break
                else:
                    number = re.findall(r"[-+]?\d*\.\d+|\d+", answer.replace(",", ""))
                    if number:
                        val = float(number[0])
                        if key == "shares":
                            if val > 1_000_000:
                                val = round(val / 1e6, 2)  # e.g. 416,000,000 → 416.0
                            elif val > 1_000:
                                val = round(val / 1e3, 2)  # e.g. 277,485 → 277.49

                        results[key] = round(val, 2)
                        break

            except Exception as e:
                continue
        else:
            print(f"❌ Could not extract {key} with any prompt.")
            results[key] = 0

    if "shares" in results and results["shares"] < 10:
        print(f"⚠️ Possible shares extraction issue — too low: {results['shares']}M")

    # Prevent any critical field from being 0 unless all attempts failed
    for k in ["shares", "cash", "debt"]:
        if results.get(k, 0) == 0:
            print(f"⚠️ Warning: {k} is 0 — extracted value may be missing or invalid.")

    return results




def generate_forecast_scenarios(text, financials, assumptions):
    prompt = (
        f"You are an equity research analyst preparing a 3-scenario DCF valuation for the company '{financials.get('company_name', 'the company')}'.\n\n"
        f"Use the extracted financials and the uploaded document excerpts to generate realistic Bull, Base, and Bear cases for Free Cash Flow to Firm (FCFF) over the next {assumptions['forecast_years']} years.\n"
        "Focus on company-specific drivers such as growth initiatives, market expansion, product pipeline, R&D investments, capital efficiency, and macroeconomic risks. Avoid generic assumptions and only use information grounded in the text.\n\n"
        "Return in **pure JSON** format (no explanations, no markdown, no comments), like:\n"
        '{ "bull": { "fcff": [...], "justification": "..." }, "base": {...}, "bear": {...} }\n\n'
        f"FINANCIALS:\n{json.dumps(financials)}\n\nDOCUMENTS:\n{text[:10000]}"
    )

    raw = call_llm(prompt, max_tokens=1600)

    try:
        clean = raw.strip().strip("```json").strip("```").strip()
        return json.loads(clean)
    except Exception as e:
        print("❌ Failed to parse LLM JSON response:", raw[:500])
        raise ValueError("LLM did not return valid JSON. Try re-running or upload more complete documents.")



def calculate_dcf_scenarios(forecast_json, assumptions, cash, debt, shares, cmp):
    if isinstance(forecast_json, str):
        forecast_json = re.sub(r"```json|```", "", forecast_json).strip()
        forecast_json = json.loads(forecast_json)

    wacc = float(assumptions["WACC"])
    terminal_growth = float(assumptions["terminal_rate_or_multiple"])
    years = int(assumptions["forecast_years"])

    def discount_fcffs(fcffs):
        return sum(f / ((1 + wacc / 100) ** (i + 1)) for i, f in enumerate(fcffs))

    output = {}
    for scenario in ["bull", "base", "bear"]:
        fcffs = forecast_json[scenario]["fcff"]
        terminal_fcff = fcffs[-1] * (1 + terminal_growth / 100) / ((wacc - terminal_growth) / 100)
        dcf_value = discount_fcffs(fcffs) + terminal_fcff / ((1 + wacc / 100) ** years)
        cash_val = extract_year_value(cash, 2024)
        debt_val = extract_year_value(debt, 2024)
        shares_val = extract_year_value(shares, 2024)
        if not shares_val or shares_val <= 0:
            raise ValueError("❌ Invalid or missing 'shares outstanding' value. Cannot proceed with DCF.")


        equity_value = dcf_value + cash_val - debt_val
        fair_value = equity_value / shares_val if shares_val else 0


        output[scenario] = {
            "fcff": fcffs,
            "dcf_value": round(dcf_value, 2),
            "equity_value": round(equity_value, 2),
            "fair_value": round(fair_value, 2),
            "cmp": cmp,
            "justification": forecast_json[scenario].get("justification", "")
        }

    return output


def format_html_output(dcf_result, ticker, cmp):
    cmp_val = float(cmp) if not isinstance(cmp, (float, int)) else cmp
    html = f"<h2>📊 DCF Valuation for {ticker}</h2><p>Current Market Price: <strong>${cmp_val:.2f}</strong></p>"
    html += "<table border='1' cellpadding='8' cellspacing='0'><tr><th>Scenario</th><th>Fair Value</th><th>Upside</th><th>Justification</th></tr>"
    for scenario, data in dcf_result.items():
        cmp_val = float(data["cmp"]) if not isinstance(data["cmp"], (float, int)) else data["cmp"]
        upside = round((data["fair_value"] - cmp_val) / cmp_val * 100, 2)
        upside_str = f"+{upside}%" if upside > 0 else f"{upside}%"

        html += f"<tr><td>{scenario.title()}</td><td>${data['fair_value']}</td><td>{upside_str}</td><td>{data['justification']}</td></tr>"
    html += "</table>"
    return html


def generate_excel_output(dcf_result):
    from io import BytesIO
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    pd.DataFrame(dcf_result).T.to_excel(writer, sheet_name="DCF Scenarios")
    writer.close()
    output.seek(0)
    return output.read()
