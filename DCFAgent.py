# DCFAgent.py

import os
import re
import time
import json
import requests

import pandas as pd

from PyPDF2 import PdfReader
import camelot
from openpyxl import Workbook
from docx import Document

# ────────────────────────────────────────────────────────────────────────────────
# CONFIGURATION: DEEPSEEK API
# ────────────────────────────────────────────────────────────────────────────────
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_CHAT_URL = os.getenv(
    "DEEPSEEK_CHAT_URL",
    "https://api.deepseek.com/v1/chat/completions"
)
if not DEEPSEEK_API_KEY:
    raise RuntimeError("Please set DEEPSEEK_API_KEY in your environment.")

# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Call DeepSeek Chat API
# ────────────────────────────────────────────────────────────────────────────────
def deepseek_chat(prompt: str, max_tokens: int = 512, temperature: float = 0.1) -> str:
    """
    Sends prompt to DeepSeek Chat and returns the assistant's reply text.
    Lower temperature yields more deterministic numeric outputs.
    """
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": max_tokens,
        "temperature": temperature
    }
    try:
        resp = requests.post(DEEPSEEK_CHAT_URL, headers=headers, json=payload, timeout=90)
        resp.raise_for_status()
        data = resp.json()
        return data["choices"][0]["message"]["content"]
    except requests.exceptions.RequestException as e:
        print(f"❌ Error communicating with DeepSeek API: {e}")
        return ""
    except Exception as e:
        print(f"❌ An unexpected error occurred with DeepSeek API response: {e}")
        return ""

# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Extract raw text from a PDF
# ────────────────────────────────────────────────────────────────────────────────
def extract_text_from_pdf(pdf_path: str) -> str:
    all_text = []
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            txt = page.extract_text() or ""
            all_text.append(txt)
    except Exception as e:
        print(f"❌ Error extracting text from PDF {pdf_path}: {e}")
    return "\n".join(all_text)

# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Extract raw text from a DOCX
# ────────────────────────────────────────────────────────────────────────────────
def extract_text_from_docx(docx_path: str) -> str:
    all_text = []
    try:
        doc = Document(docx_path)
        for para in doc.paragraphs:
            all_text.append(para.text)
    except Exception as e:
        print(f"❌ Error extracting text from DOCX {docx_path}: {e}")
    return "\n".join(all_text)

# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Identify page numbers containing a given keyword in PDF
# ────────────────────────────────────────────────────────────────────────────────
def find_pages_with_keyword(pdf_path: str, keyword: str) -> list[int]:
    pages = set()
    try:
        reader = PdfReader(pdf_path)
        for idx, page in enumerate(reader.pages):
            txt = page.extract_text() or ""
            if re.search(keyword, txt, re.IGNORECASE):
                pages.add(idx)
    except Exception as e:
        print(f"⚠️ Warning: Error finding keyword '{keyword}' in PDF {pdf_path}: {e}")
    return sorted(list(pages))

# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Extract tables from PDF pages
# ────────────────────────────────────────────────────────────────────────────────
def extract_tables_from_pdf(pdf_path: str, page_indices: list[int]) -> list[pd.DataFrame]:
    dfs = []
    if not page_indices:
        return dfs
    pages_str = ",".join(str(i + 1) for i in page_indices)
    try:
        tables = camelot.read_pdf(pdf_path, pages=pages_str, flavor="lattice", split_text=True, line_scale=40)
        if not tables or all(tbl.df.empty for tbl in tables):
            tables = camelot.read_pdf(pdf_path, pages=pages_str, flavor="stream", split_text=True, line_scale=40)
        for t in tables:
            if not t.df.empty:
                dfs.append(t.df)
        print(f"[EXTRACT DEBUG] Extracted {len(dfs)} tables from pages {pages_str} of {pdf_path}.")
    except Exception as e:
        print(f"❌ Error extracting tables from PDF {pdf_path} pages {pages_str}: {e}")
    return dfs

# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Parse numerical value from string, handling various formats
# ────────────────────────────────────────────────────────────────────────────────
def parse_numeric_value(s: str) -> float | None:
    """Safely converts a string to a float, handling commas, dollar signs, parentheses for negatives."""
    if not isinstance(s, str):
        return None
    s = s.strip().replace(",", "").replace("$", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except ValueError:
        return None

# ────────────────────────────────────────────────────────────────────────────────
# (A) Build a single “file_context” from multiple uploaded PDFs / DOCXs
# ────────────────────────────────────────────────────────────────────────────────
def prepare_pdf_context(file_paths: list[str]) -> dict:
    """
    Given a list of paths to PDF or DOCX financial reports, returns:
      - 'all_text': a big string of every page’s extracted text
      - 'tables': a list of 7×7 CSV snippets from each detected table
    """
    all_text_segments = []
    all_table_snippets = []

    for path in file_paths:
        name, ext = os.path.splitext(path.lower())
        file_tag = os.path.basename(path)

        if ext == ".pdf":
            # Extract all text
            try:
                reader = PdfReader(path)
                for page_num, page in enumerate(reader.pages, start=1):
                    txt = page.extract_text() or ""
                    all_text_segments.append(f"--- [{file_tag} PAGE {page_num}] ---\n{txt}")
            except Exception as e:
                print(f"⚠️ Could not extract text from PDF {path}: {e}")

            # Extract tables on every page
            try:
                num_pages = len(PdfReader(path).pages)
                for pg_idx in range(1, num_pages + 1):
                    try:
                        tables = camelot.read_pdf(path, pages=str(pg_idx), flavor="lattice", split_text=True, line_scale=40)
                        if not tables or all(tbl.df.empty for tbl in tables):
                            tables = camelot.read_pdf(path, pages=str(pg_idx), flavor="stream", split_text=True, line_scale=40)
                        for tbl in tables:
                            df = tbl.df
                            if df.shape[0] > 0 and df.shape[1] > 0:
                                snippet = df.iloc[:7, :7].to_csv(index=False)
                                all_table_snippets.append(f"--- Table snippet from {file_tag} page {pg_idx} ---\n{snippet}")
                    except Exception:
                        continue
            except Exception as e:
                print(f"⚠️ Could not run Camelot on {path}: {e}")

        elif ext == ".docx":
            try:
                doc = Document(path)
                text_body = "\n".join([p.text for p in doc.paragraphs])
                all_text_segments.append(f"--- [{file_tag}] DOCX Text ---\n{text_body}")
            except Exception as e:
                print(f"⚠️ Could not extract text from DOCX {path}: {e}")

        else:
            print(f"⚠️ Unsupported file type for {path}, skipping.")

    return {
        "all_text": "\n\n".join(all_text_segments),
        "tables": all_table_snippets
    }

# ────────────────────────────────────────────────────────────────────────────────
# (B) Query the PDF context for a specific financial figure or table‐based answer
# ────────────────────────────────────────────────────────────────────────────────
def query_pdf_with_context(file_context: dict, question: str) -> str:
    """
    Ask a natural-language question about the user's uploaded files.
    file_context = { "all_text": str, "tables": [csv_snippet, ...] }
    Returns the LLM’s answer (ideally a single numeric string or 'N/A').
    """
    text_block = file_context.get("all_text", "")
    tables_block = "\n\n".join(file_context.get("tables", []))

    prompt = (
        "You are a financial data assistant. Below is all the text and every detected table (in CSV) "
        "from the user’s uploaded financial reports. If the user asks a question like "
        "“What was 2023 revenue?” or “What is the latest total debt?”, you should search this content "
        "and give a single numeric answer (e.g. “$41,227,600,000” or “N/A”). Do not add any extra commentary.\n\n"
        f"--- ALL EXTRACTED TEXT BELOW ---\n{text_block}\n\n"
        f"--- ALL TABLE SNIPPETS BELOW ---\n{tables_block}\n\n"
        f"QUESTION: {question}\n"
        "ANSWER:"
    )

    raw = deepseek_chat(prompt, max_tokens=256, temperature=0.1)
    return raw.strip()

# ────────────────────────────────────────────────────────────────────────────────
# (C) Build FCFF Forecast Using Guidance, Consensus, & Segment Growth
# ────────────────────────────────────────────────────────────────────────────────
def build_fcff_forecast(
    hist_year: int,
    last_historical_fcf: float,
    segment_summary_text: str
) -> dict[int, float]:
    """
    Returns a dict {year: fcf} for the next 5 fiscal years.
    Year 1 & 2: No external guidance/consensus; assume fallback growth (5% by default or LLM-estimated).
    Years 3-5: LLM‐estimated long-term growth rate (default 3%).
    """
    def get_midpoint(text: str) -> float | None:
        if not text or not isinstance(text, str):
            return None
        txt = text.replace(",", "").replace("$", "").lower()
        nums = re.findall(r"(\d+\.?\d*)\s*(million|billion|m|b)?", txt)
        vals = []
        for num_str, scale in nums:
            try:
                val = float(num_str)
                if scale in ("million", "m"):
                    val *= 1_000_000
                elif scale in ("billion", "b"):
                    val *= 1_000_000_000
                vals.append(val)
            except:
                continue
        if not vals:
            return None
        return (vals[0] + vals[1]) / 2.0 if len(vals) >= 2 else vals[0]

    forecast = {}
    y1 = hist_year + 1
    y2 = hist_year + 2

    # Determine initial growth rate via LLM (default 5%)
    initial_growth_rate = 0.05
    if segment_summary_text and segment_summary_text.strip() != "N/A":
        prompt_initial = (
            f"You are a financial modeler. Based on this segment performance summary:\n\n"
            f"{segment_summary_text}\n\n"
            f"Estimate an initial year‐over‐year growth rate (decimal) for Free Cash Flow (Year {hist_year}→{y1}). "
            f"Return just a decimal (e.g., 0.065). If unsure, return 0.05."
        )
        raw_initial = deepseek_chat(prompt_initial, max_tokens=10, temperature=0.0)
        m = re.search(r"0\.\d+", raw_initial)
        if m:
            try:
                r = float(m.group(0))
                initial_growth_rate = r if 0.01 <= r <= 0.15 else 0.05
            except:
                initial_growth_rate = 0.05

    # Year 1 = last_historical_fcf * (1 + initial_growth_rate)
    fcf_y1 = last_historical_fcf * (1 + initial_growth_rate)
    forecast[y1] = fcf_y1

    # Year 2 = fcf_y1 * (1 + initial_growth_rate)
    fcf_y2 = fcf_y1 * (1 + initial_growth_rate)
    forecast[y2] = fcf_y2

    # Determine long‐term growth via LLM (default 3%)
    long_term_rate = 0.03
    if segment_summary_text and segment_summary_text.strip() != "N/A":
        prompt_long = (
            f"You are a financial modeler. Based on this segment performance summary and long‐term outlook (Years {hist_year+3}-{hist_year+5}):\n\n"
            f"{segment_summary_text}\n\n"
            f"Estimate a long‐term annual growth rate (decimal) for Free Cash Flow for Years {hist_year+3} to {hist_year+5}. "
            f"Return just a decimal (e.g. 0.03). If unsure, return 0.03."
        )
        raw_long = deepseek_chat(prompt_long, max_tokens=10, temperature=0.0)
        m2 = re.search(r"0\.\d+", raw_long)
        if m2:
            try:
                r2 = float(m2.group(0))
                long_term_rate = r2 if 0.01 <= r2 <= 0.10 else 0.03
            except:
                long_term_rate = 0.03

    # Years 3–5
    for i in range(3, 6):
        year = hist_year + i
        prev = forecast[year - 1]
        forecast[year] = prev * (1 + long_term_rate)

    return forecast

# ────────────────────────────────────────────────────────────────────────────────
# (D) Calculate DCF
# ────────────────────────────────────────────────────────────────────────────────
def calculate_dcf(
    forecasted_fcff: dict[int, float],
    terminal_growth_rate: float,
    wacc: float,
    net_debt: float,
    shares_outstanding: int,
    current_share_price: float | None = None
) -> dict:
    """
    Performs a Discounted Cash Flow calculation.

    Args:
        forecasted_fcff (dict): {year: FCF_value}
        terminal_growth_rate (float): Perpetual growth rate (e.g. 0.02)
        wacc (float): Discount rate (e.g. 0.08)
        net_debt (float): Total Debt – Cash
        shares_outstanding (int)
        current_share_price (float | None): For deviation calculation

    Returns:
        dict with PVs, EV, Equity Value, intrinsic value per share, etc.
    """
    if wacc <= 0 or terminal_growth_rate >= wacc:
        return {"error": "Invalid DCF inputs: WACC must exceed terminal growth."}

    sorted_years = sorted(forecasted_fcff.keys())
    n = len(sorted_years)
    pv_fcff = 0.0

    for i, year in enumerate(sorted_years):
        fcf = forecasted_fcff[year]
        discount = (1 + wacc) ** (i + 1)
        pv = fcf / discount
        pv_fcff += pv

    last_year = sorted_years[-1]
    last_fcf = forecasted_fcff[last_year]
    terminal_value = last_fcf * (1 + terminal_growth_rate) / (wacc - terminal_growth_rate)
    pv_tv = terminal_value / ((1 + wacc) ** n)

    enterprise_value = pv_fcff + pv_tv
    equity_value = enterprise_value - net_debt
    intrinsic_per_share = equity_value / shares_outstanding

    result = {
        "forecasted_fcff": forecasted_fcff,
        "present_value_of_explicit_fcff": pv_fcff,
        "terminal_value": terminal_value,
        "present_value_of_terminal_value": pv_tv,
        "enterprise_value": enterprise_value,
        "net_debt": net_debt,
        "equity_value": equity_value,
        "shares_outstanding": shares_outstanding,
        "intrinsic_value_per_share": intrinsic_per_share
    }

    if current_share_price is not None:
        deviation = ((intrinsic_per_share - current_share_price) / current_share_price) * 100
        result["current_share_price"] = current_share_price
        result["deviation_from_current_price_pct"] = deviation

    return result

# ────────────────────────────────────────────────────────────────────────────────
# (E) Full DCF Agent That Combines All Steps
# ────────────────────────────────────────────────────────────────────────────────
def run_dcf_model(
    company_name: str,
    assumptions: dict,  # Expected keys:
                        #   'base_wacc', 'base_terminal_growth_rate',
                        #   'bull_growth_multiplier', 'bull_wacc_adjust', 'bull_terminal_growth_adjust',
                        #   'bear_growth_multiplier', 'bear_wacc_adjust', 'bear_terminal_growth_adjust'
    file_paths: list[str],
    current_share_price: float | None = None,
    pdf_answers: dict = None
) -> tuple[str, dict]:
    """
    Performs a 3-case (Base / Bull / Bear) FCFF-based DCF valuation:
      1) Ask LLM for ticker.
      2) Build file_context from uploaded PDFs/DOCXs.
      3) Query the PDF for revenue, FCF, total debt, cash, shares outstanding, fiscal year.
      4) Build 5-year FCFF forecast.
      5) Calculate DCF for Base / Bull / Bear scenarios.
      6) Return summary message and detailed results.
    """
    results = {}
    print(f"🚀 Starting DCF analysis for {company_name}...")

    # 1) Determine ticker via LLM
    ticker_prompt = f"You are a stock ticker lookup assistant. For the company name \"{company_name}\", provide the stock ticker symbol (just the ticker)."
    raw_ticker = deepseek_chat(ticker_prompt, max_tokens=8, temperature=0.0).strip().upper()
    if not raw_ticker:
        return "❌ Error: Could not determine ticker symbol for the company.", {}
    ticker = raw_ticker
    print(f"✅ Determined ticker: {ticker}")

    # Fetch current share price if not provided
    if current_share_price is None:
        current_share_price = None  # No automatic fetch; user must provide or skip.
        print("⚠️ Current share price not provided; skipping price comparison.")
    else:
        print(f"✅ Using provided current market price: ${current_share_price:,.2f}")

    # 2) Build file_context
    file_context = prepare_pdf_context(file_paths)

    # Helper to parse LLM’s numeric answer
    # Updated get_number_from_pdf
    def get_number_from_pdf(question: str) -> float | None:
        answer = None
        if pdf_answers and question in pdf_answers:
            answer = pdf_answers[question]
        else:
            answer = query_pdf_with_context(file_context, question)

        m = re.search(r"([\d,]+(?:\.\d+)?)\s*(million|billion|M|B)?", answer, re.IGNORECASE)
        if not m:
            return None
        num_str, scale = m.groups()
        try:
            num = float(num_str.replace(",", ""))
        except:
            return None
        if scale:
            s = scale.lower()
            if s in ("billion", "b"):
                num *= 1_000_000_000
            elif s in ("million", "m"):
                num *= 1_000_000
        return num


    # 3) Query for required financial inputs
    # a) Revenue
    revenue_q = "What was the most recent annual revenue (in USD) reported in these uploaded files?"
    latest_revenue = get_number_from_pdf(revenue_q)
    if latest_revenue is not None:
        print(f"✅ Revenue from PDF: ${latest_revenue:,.0f}")
    else:
        return "❌ Error: Could not find Revenue via PDF query.", {}

    # b) Free Cash Flow
    fcf_q = "What was the most recent Free Cash Flow (FCF) reported in these files?"
    latest_fcf = get_number_from_pdf(fcf_q)
    if latest_fcf is not None:
        print(f"✅ FCF from PDF: ${latest_fcf:,.0f}")
    else:
        return "❌ Error: Could not find Free Cash Flow via PDF query.", {}

    # c) Total Debt
    debt_q = "What was the most recent Total Debt (in USD) reported?"
    latest_debt = get_number_from_pdf(debt_q)
    if latest_debt is not None:
        print(f"✅ Total Debt from PDF: ${latest_debt:,.0f}")
    else:
        return "❌ Error: Could not find Total Debt via PDF query.", {}

    # d) Cash & Cash Equivalents
    cash_q = "What was the most recent Cash and Cash Equivalents (in USD) reported?"
    latest_cash = get_number_from_pdf(cash_q)
    if latest_cash is not None:
        print(f"✅ Cash & CE from PDF: ${latest_cash:,.0f}")
    else:
        return "❌ Error: Could not find Cash & CE via PDF query.", {}

    # e) Shares Outstanding
    shares_q = "How many shares outstanding (number of shares) are reported in these files?"
    shares_val = get_number_from_pdf(shares_q)
    if shares_val is not None:
        shares_outstanding = int(shares_val)
        print(f"✅ Shares Outstanding from PDF: {shares_outstanding:,}")
    else:
        return "❌ Error: Could not find Shares Outstanding via PDF query.", {}

    # f) Fiscal Year
    year_q = "What is the fiscal year of the most recent report (e.g. 2023 or 2024)?"
    year_val = get_number_from_pdf(year_q)
    if year_val is not None:
        hist_year = int(year_val)
        print(f"✅ Fiscal Year from PDF: {hist_year}")
    else:
        return "❌ Error: Could not find fiscal year via PDF query.", {}

    # Compute net debt
    net_debt = latest_debt - latest_cash
    print(f"✅ Computed Net Debt = ${latest_debt:,.0f} - ${latest_cash:,.0f} = ${net_debt:,.0f}")

    # 4) Query PDF for segment margins text (to feed into forecast)
    seg_q = "For each business segment, what is the most recent operating margin (in %) reported?"
    segment_margins_text = query_pdf_with_context(file_context, seg_q)
    print(f"✅ Segment margins text:\n{segment_margins_text}")

    # 5) Build a 5-year FCFF forecast
    print("\n📈 Building 5-year FCFF forecast...")
    base_forecast_fcff = build_fcff_forecast(
        hist_year=hist_year,
        last_historical_fcf=latest_fcf,
        segment_summary_text=segment_margins_text
    )
    if not base_forecast_fcff:
        return "❌ Error: Could not build FCFF forecast.", {}
    print(f"✅ Base FCFF Forecast: {base_forecast_fcff}")

    # 6) Calculate DCF for Base, Bull, Bear
    print("\n📊 Calculating DCF valuation for Base, Bull, and Bear cases...")
    results = {}

    # Base Case
    base_wacc = assumptions.get("base_wacc", 0.08)
    base_tgr = assumptions.get("base_terminal_growth_rate", 0.02)
    print("\n--- Base Case ---")
    base_result = calculate_dcf(
        forecasted_fcff=base_forecast_fcff,
        terminal_growth_rate=base_tgr,
        wacc=base_wacc,
        net_debt=net_debt,
        shares_outstanding=shares_outstanding,
        current_share_price=current_share_price
    )
    results["base_case"] = base_result

    # Bull Case
    bull_multiplier = assumptions.get("bull_growth_multiplier", 1.10)
    bull_fcff = {yr: val * bull_multiplier for yr, val in base_forecast_fcff.items()}
    bull_wacc = base_wacc + assumptions.get("bull_wacc_adjust", -0.005)
    bull_tgr = base_tgr + assumptions.get("bull_terminal_growth_adjust", 0.005)
    print("\n--- Bull Case ---")
    bull_result = calculate_dcf(
        forecasted_fcff=bull_fcff,
        terminal_growth_rate=bull_tgr,
        wacc=bull_wacc,
        net_debt=net_debt,
        shares_outstanding=shares_outstanding,
        current_share_price=current_share_price
    )
    results["bull_case"] = bull_result

    # Bear Case
    bear_multiplier = assumptions.get("bear_growth_multiplier", 0.90)
    bear_fcff = {yr: val * bear_multiplier for yr, val in base_forecast_fcff.items()}
    bear_wacc = base_wacc + assumptions.get("bear_wacc_adjust", 0.005)
    bear_tgr = base_tgr + assumptions.get("bear_terminal_growth_adjust", -0.005)
    print("\n--- Bear Case ---")
    bear_result = calculate_dcf(
        forecasted_fcff=bear_fcff,
        terminal_growth_rate=bear_tgr,
        wacc=bear_wacc,
        net_debt=net_debt,
        shares_outstanding=shares_outstanding,
        current_share_price=current_share_price
    )
    results["bear_case"] = bear_result

    # Build final message
    msg = f"✅ DCF analysis complete for {company_name} ({ticker})."
    if current_share_price is not None:
        msg += f"\n   Current Share Price: ${current_share_price:,.2f}"
        if "intrinsic_value_per_share" in base_result:
            iv = base_result["intrinsic_value_per_share"]
            dev = base_result.get("deviation_from_current_price_pct", 0.0)
            msg += f"\n   Base Intrinsic Value: ${iv:,.2f}  (Deviation: {dev:+.2f}%)"
        if "intrinsic_value_per_share" in bull_result:
            ivb = bull_result["intrinsic_value_per_share"]
            msg += f"\n   Bull Intrinsic Value: ${ivb:,.2f}"
        if "intrinsic_value_per_share" in bear_result:
            ivr = bear_result["intrinsic_value_per_share"]
            msg += f"\n   Bear Intrinsic Value: ${ivr:,.2f}"

    return msg, results
