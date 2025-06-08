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
    from InvMemo import PDFQueryEngine  # ✅ uses your existing class

    engine = PDFQueryEngine()  # no arguments since your class doesn’t accept any
    engine.chunks = []
    for path in filepaths:
        engine.chunks.extend(engine.extract_text_from_pdf(path))
    engine.chunk_texts = [text for _, text in engine.chunks]
    engine.embeddings = engine.embed_texts(engine.chunk_texts)
    engine.index = engine.build_faiss_index(engine.embeddings)
    
    from types import MethodType

    def answer_query_bound(self, query: str) -> str:
        query_embedding = self.embedder.encode([query], convert_to_numpy=True)
        distances, indices = self.index.search(query_embedding, 5)
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

    queries = {
        "cash": "Only provide the total cash or cash and cash equivalents for 2024. Only return a number.",
        "debt": "Only provide the total debt or borrowings for 2024. Only return a number.",
        "shares": "Only provide diluted shares outstanding for 2024. Only return a number.",
        "net_income": "Only provide net income for 2024, 2023, 2022, 2021, 2020. Only return a number.",
        "diluted_eps": "Only provide diluted earnings per share (EPS) for 2024. Only return a number.",
        "revenue": "Only provide total revenue for the company for 2024, 2023, 2022, 2021, 2020. Only return a number.",
        "operating_income": "Only provide net income for 2024, 2023, 2022, 2021, 2020. Only return a number.",
        "EBITDA": "Only provide EBITDA for 2024, 2023, 2022, 2021, 2020. Only return a number.",
        "segment_revenue": "Only provide segment revenue for 2024, 2023, 2022, 2021, 2020. Only return a number.",
        "segment_EBITDA": "Only provide segment EBITDA for 2024, 2023, 2022, 2021, 2020. Only return a number.",
        "free_cash_flow": "Only provide free cash flow for 2024, 2023, 2022, 2021, 2020. Only return a number.",
        "capex": "Only provide capital expenditure for 2024, 2023, 2022, 2021, 2020. Only return a number."
    }

    results = {}
    for key, query in queries.items():
        try:
            answer = engine.answer_query(query)
            print(f"📌 {key}: {answer}")
            number = re.findall(r"[-+]?\d*\.\d+|\d+", answer.replace(",", ""))
            results[key] = float(number[0]) if number else 0
        except Exception as e:
            print(f"❌ Failed to extract {key}: {e}")
            results[key] = 0

    if results.get("shares", 0) > 1000:
        results["shares"] = round(results["shares"] / 1e6, 2)

    return results



def generate_forecast_scenarios(text, financials, assumptions):
    prompt = (
        f"Based on the following financial data and company document excerpts, generate 3 forecast scenarios "
        f"(bull, base, and bear) for Free Cash Flow to Firm (FCFF) over the next {assumptions['forecast_years']} years.\n"
        "Return in pure JSON format like:\n"
        '{ "bull": { "fcff": [...], "justification": "..." }, "base": {...}, "bear": {...} }\n\n'
        f"FINANCIALS:\n{json.dumps(financials)}\n\nTEXT:\n{text[:8000]}"
    )
    return call_llm(prompt, max_tokens=1600)


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


def format_html_output(dcf_result, ticker, cmp):
    html = f"<h2>📊 DCF Valuation for {ticker}</h2><p>Current Market Price: <strong>${cmp}</strong></p>"
    html += "<table border='1' cellpadding='8' cellspacing='0'><tr><th>Scenario</th><th>Fair Value</th><th>Upside</th><th>Justification</th></tr>"
    for scenario, data in dcf_result.items():
        upside = round((data["fair_value"] - cmp) / cmp * 100, 2)
        html += f"<tr><td>{scenario.title()}</td><td>${data['fair_value']}</td><td>{upside}%</td><td>{data['justification']}</td></tr>"
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
