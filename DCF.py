import os
import requests
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from PyPDF2 import PdfReader
from docx import Document
from openai import OpenAI
from bs4 import BeautifulSoup
import hashlib

# ========== CONFIG ==========
FMP_API_KEY = os.getenv("FMP_API_KEY")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=OPENAI_API_KEY)

# ========== Ticker & Price ==========
def get_fmp_ticker(company_name):
    prompt = f"""
What is the FMP-compatible stock ticker for the company: {company_name}?
Only return the raw ticker symbol (e.g., AAPL or MSFT). Do not include any explanation or extra text.
"""
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0
    )
    raw_output = response.choices[0].message.content.strip().upper()
    ticker = raw_output.split()[0].strip('".:,')
    return ticker

def get_current_price(ticker):
    url = f"https://financialmodelingprep.com/api/v3/quote/{ticker}?apikey={FMP_API_KEY}"
    response = requests.get(url).json()
    if response and isinstance(response, list):
        return round(response[0].get("price", 0), 2)
    return None

# ========== Financials & Excel ==========
def get_fmp_data(ticker):
    def fetch(endpoint):
        url = f"https://financialmodelingprep.com/api/v3/{endpoint}/{ticker}?period=annual&apikey={FMP_API_KEY}"
        return requests.get(url).json()

    income = fetch("income-statement")
    balance = fetch("balance-sheet-statement")
    cashflow = fetch("cash-flow-statement")

    data = pd.DataFrame()
    for year in range(min(len(income), len(balance), len(cashflow))):
        row = {
            "Year": income[year]["calendarYear"],
            "Revenue": income[year].get("revenue"),
            "EBITDA": income[year].get("ebitda"),
            "Operating Income": income[year].get("operatingIncome"),
            "Net Income": income[year].get("netIncome"),
            "EPS (Diluted)": income[year].get("epsdiluted"),
            "Shares Outstanding": income[year].get("weightedAverageShsOutDil"),
            "Cash": balance[year].get("cashAndCashEquivalents"),
            "Short-term Investments": balance[year].get("shortTermInvestments"),
            "Receivables": balance[year].get("netReceivables"),
            "Inventory": balance[year].get("inventory"),
            "Payables": balance[year].get("accountPayables"),
            "Short-term Debt": balance[year].get("shortTermDebt"),
            "Long-term Debt": balance[year].get("longTermDebt"),
            "CapEx": cashflow[year].get("capitalExpenditure"),
            "FCF": cashflow[year].get("freeCashFlow"),
            "Change in WC": cashflow[year].get("changeInWorkingCapital")
        }
        data = pd.concat([data, pd.DataFrame([row])], ignore_index=True)
    return data

def save_excel(data):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="Financials")
    output.seek(0)
    return output

# ========== Document Parsing ==========
def extract_text_from_files(uploaded_files):
    chunked_docs = []
    for file in uploaded_files:
        if hasattr(file, "name") and file.name.endswith(".pdf"):
            reader = PdfReader(file)
            for i, page in enumerate(reader.pages):
                text = page.extract_text()
                if text:
                    chunked_docs.append(f"[PDF: {file.name}, Page {i+1}]\n{text}")
        elif hasattr(file, "name") and file.name.endswith(".docx"):
            doc = Document(file)
            paras = [para.text for para in doc.paragraphs if para.text.strip()]
            chunked_docs.append(f"[DOCX: {file.name}]\n" + "\n".join(paras))
        elif hasattr(file, "name") and file.name.endswith(".txt"):
            text = file.read().decode("utf-8")
            chunked_docs.append(f"[TXT: {file.name}]\n{text}")
    return "\n\n".join(chunked_docs)


def extract_kpi_drivers(documents_text):
    prompt = f"""
Extract 2025 forward-looking guidance from RTX's earnings call and investor document below. Focus on:

- Revenue range or growth %
- EBITDA or segment margins
- CapEx guidance
- Free Cash Flow target
- EPS guidance
- Segment-level performance
- Any comments supporting these assumptions

Return the extracted drivers in clean bullet points with values, and tag source if visible (e.g. "Slide 13" or "Page 14").

📄 RTX Investor PDF Extract:
{documents_text}
"""
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.0
    )
    return response.choices[0].message.content.strip()



# ========== LLM-Based DCF ==========

def generate_dcf_logic(financials_df, documents_text_annotated, wacc, current_price, mode):
    import hashlib

    shares = financials_df['Shares Outstanding'].dropna().astype(float).iloc[0]
    net_debt = (
        financials_df['Short-term Debt'].dropna().astype(float).iloc[0] +
        financials_df['Long-term Debt'].dropna().astype(float).iloc[0] -
        financials_df['Cash'].dropna().astype(float).iloc[0]
    )
    latest_revenue = financials_df['Revenue'].dropna().astype(float).iloc[0]
    latest_fcf = financials_df['FCF'].dropna().astype(float).iloc[0]

    # Deterministic prompt seed
    data_to_hash = str(financials_df.to_dict()) + documents_text_annotated + str(wacc)
    prompt_seed = hashlib.md5(data_to_hash.encode()).hexdigest()[:8]

    prompt = f"""
[SEED: {prompt_seed}]

You are an equity analyst creating a detailed, **document-driven** DCF forecast for RTX using uploaded guidance and actual financials.

🎯 Objective:
Build Bull, Base, and Bear forecasts using 2025 guidance and segment updates from the uploaded RTX investor document.

📘 Key Inputs:
- CMP: ${current_price}
- WACC: {wacc}%
- Shares Outstanding: {shares:,.0f}
- Net Debt: ${net_debt:,.0f}
- Latest Revenue: ${latest_revenue:,.0f}
- Latest FCF: ${latest_fcf:,.0f}
- Forecast Years: 2025–2029
- Terminal Growth: Bull (2.5%), Base (2.0%), Bear (1.5%)

📂 Source Data:
{documents_text_annotated if documents_text_annotated.strip() else "No documents provided. Use trailing 3Y trends only."}

🔍 Step-by-Step Instructions:
1. Extract the following from the uploaded document:
   - 2025 Sales guidance range
   - Organic growth %
   - CapEx guidance
   - Adjusted operating/EBITDA margin
   - Free Cash Flow guidance

2. Use that to justify:
   - Revenue CAGR (Bull, Base, Bear)
   - EBITDA Margin (Bull, Base, Bear)
   - CapEx % (use doc value: ~3.2%)
   - D&A = 5% of revenue
   - WC change = 2%
   - Tax = 25%

3. For each year 2025–2029 and each scenario:
   - Revenue = Prior × CAGR
   - EBITDA = Revenue × Margin
   - EBIT = EBITDA - D&A
   - NOPAT = EBIT × (1 - Tax)
   - FCF = NOPAT - CapEx - ΔWC

4. Compute:
   - Terminal Value = 2029 FCF × (1 + g) / (WACC - g)
   - PV of FCFs and TV
   - Enterprise Value = PV(FCFs + TV)
   - Equity Value = EV - Net Debt
   - Per-share value = Equity / Shares

✏️ Output:
Start with short scenario justification using document quotes (e.g., “per Slide 13, RTX expects 4–6% organic sales growth”).

Then output tables:

| Year | Revenue | EBITDA Margin | CapEx % | FCF |
|------|---------|----------------|--------|-----|

Finally:

| Scenario | Per-Share Value | Upside (%) |
|----------|-----------------|------------|

⚠️ If you hallucinate numbers (e.g., $150B terminal value) or don’t explain source, re-run.
    """

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0
    )
    return response.choices[0].message.content.strip()



# ========== Cleaned HTML Output ==========
def clean_and_format_dcf_output(raw_text, current_price):
    def extract_per_share_values(text):
        lines = text.splitlines()
        per_share_values = {"Bull": None, "Base": None, "Bear": None}
        for scenario in per_share_values:
            for line in lines:
                if scenario.lower() in line.lower() and ("per-share" in line.lower() or "per share" in line.lower()):
                    try:
                        value = float(line.split("$")[-1].strip())
                        per_share_values[scenario] = value
                    except:
                        continue
        return per_share_values

    def get_upside_text(fv, cmp):
        pct = ((fv - cmp) / cmp) * 100
        direction = "Upside" if pct >= 0 else "Downside"
        color = "green" if pct >= 0 else "red"
        return f'<span style="color:{color};">{pct:+.2f}% {direction}</span>'

    soup = BeautifulSoup(raw_text, "html.parser")
    text = soup.get_text()
    if "To generate the Bull" in text:
        text = text.split("To generate the Bull")[1]
        text = "Bull" + text[text.find("Bull"):]

    per_share_values = extract_per_share_values(text)
    cmp = current_price

    comparison_html = f"""
      <div class='dcf-scenario'>
        <h3 class='dcf-title'>📊 Comparison to CMP (${cmp:.2f})</h3>
        <ul>
    """
    for scenario, val in per_share_values.items():
        if val is not None:
            comparison_html += f"<li><strong>{scenario}:</strong> ${val:.2f} → {get_upside_text(val, cmp)}</li>"
    comparison_html += "</ul></div>"

    clean_text = text.replace("\\frac", "").replace("^", "**").replace("∗", "*")
    clean_text = clean_text.replace("millionmillion", "million")
    clean_text = clean_text.replace("\n", "<br>")

    return f"<div class='dcf-container'><div class='dcf-scenario'>{clean_text.strip()}</div>{comparison_html}</div>"
