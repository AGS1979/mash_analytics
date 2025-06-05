# DCFAgent.py

import os
import json
import time
import re
import requests

import pandas as pd
import numpy as np
import yfinance as yf

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
def deepseek_chat(prompt: str, max_tokens: int = 512) -> str:
    """
    Sends prompt to DeepSeek Chat and returns the assistant's reply text.
    """
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": max_tokens
    }
    resp = requests.post(DEEPSEEK_CHAT_URL, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Fetch shares outstanding & current price via yfinance
# ────────────────────────────────────────────────────────────────────────────────
def get_shares_outstanding(ticker: str) -> int | None:
    try:
        info = yf.Ticker(ticker).info
        return info.get("sharesOutstanding")
    except Exception:
        return None


def get_current_share_price(ticker: str) -> float | None:
    try:
        info = yf.Ticker(ticker).info
        return info.get("regularMarketPrice")
    except Exception:
        return None


def get_yfinance_fcf(ticker: str) -> float | None:
    """
    Fallback: Pull most recent Free Cash Flow from yfinance cashflow statement.
    """
    try:
        cf = yf.Ticker(ticker).cashflow
        # yfinance cashflow columns are dates; take the most recent column
        most_recent = cf.columns[0]
        # "Free Cash Flow" or "Free Cash Flow (FCF)" might be a row index
        for label in ["Free Cash Flow", "FreeCashFlow", "Free Cash Flow (FCF)"]:
            if label in cf.index:
                val = cf.loc[label, most_recent]
                return float(val)  # already in USD absolute
    except Exception:
        return None
    return None


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Extract raw text from a PDF
# ────────────────────────────────────────────────────────────────────────────────
def extract_text_from_pdf(pdf_path: str) -> str:
    """
    Extracts and concatenates all textual content from a PDF file.
    """
    all_text = []
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            txt = page.extract_text() or ""
            all_text.append(txt)
    except Exception:
        pass
    return "\n".join(all_text)


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Extract raw text from a DOCX
# ────────────────────────────────────────────────────────────────────────────────
def extract_text_from_docx(docx_path: str) -> str:
    """
    Extracts and concatenates all textual content from a DOCX file.
    """
    all_text = []
    try:
        doc = Document(docx_path)
        for para in doc.paragraphs:
            all_text.append(para.text)
    except Exception:
        pass
    return "\n".join(all_text)


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Identify page numbers containing a given keyword in PDF
# ────────────────────────────────────────────────────────────────────────────────
def find_pages_with_keyword(pdf_path: str, keyword: str) -> list[int]:
    """
    Returns a list of 0-based page indices where keyword appears in the text.
    Keyword match is case-insensitive.
    """
    pages = []
    try:
        reader = PdfReader(pdf_path)
        for idx, page in enumerate(reader.pages):
            txt = page.extract_text() or ""
            if re.search(keyword, txt, re.IGNORECASE):
                pages.append(idx)
    except Exception:
        pass
    return pages


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Extract tables from PDF pages
# ────────────────────────────────────────────────────────────────────────────────
def extract_tables_from_pdf(pdf_path: str, page_indices: list[int]) -> list[pd.DataFrame]:
    """
    Uses Camelot to extract all tables on the specified 1-based pages.
    Returns a list of DataFrames.
    """
    dfs = []
    if not page_indices:
        return dfs
    pages_str = ",".join(str(i + 1) for i in page_indices)
    try:
        tables = camelot.read_pdf(pdf_path, pages=pages_str, flavor='stream')
        for t in tables:
            dfs.append(t.df)
    except Exception:
        pass
    return dfs


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Attempt to parse Free Cash Flow from a DataFrame
# ────────────────────────────────────────────────────────────────────────────────
def parse_fcf_from_df(df: pd.DataFrame) -> float | None:
    """
    Searches df for a row containing 'Free Cash Flow' (case-insensitive)
    and returns the first numeric value in that row (USD).
    Detects if table headers mention millions/billions anywhere in the first 5 rows.
    """
    # 1) Determine scale by scanning up to first 5 rows for "in millions"/"in billions"
    scale = 1
    header_rows = min(5, len(df))
    for i in range(header_rows):
        row_str = " ".join(str(cell) for cell in df.iloc[i].tolist()).lower()
        if "in millions" in row_str:
            scale = 1_000_000
            break
        elif "in billions" in row_str:
            scale = 1_000_000_000
            break

    # 2) Look for "Free Cash Flow" in any row
    for _, row in df.iterrows():
        row_str = " ".join(str(cell) for cell in row.tolist())
        if re.search(r"free\s+cash\s+flow", row_str, re.IGNORECASE):
            # Once found, grab the first number in that row
            for cell in row.tolist()[1:]:
                s = str(cell).replace(",", "").replace("$", "").replace("(", "-").replace(")", "")
                try:
                    val = float(s)
                    return val * scale
                except Exception:
                    continue
    return None


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Attempt to parse Revenue & Net Income from Income Statement DataFrame
# ────────────────────────────────────────────────────────────────────────────────
def parse_income_from_df(df: pd.DataFrame) -> tuple[float | None, float | None]:
    """
    Searches df for 'Revenue' and 'Net Income' rows (case-insensitive).
    Returns a tuple (revenue, net_income) in USD.
    Detects scale by scanning up to first 5 rows.
    """
    rev = None
    ni = None

    # Determine scale by scanning up to first 5 rows
    scale = 1
    header_rows = min(5, len(df))
    for i in range(header_rows):
        row_str = " ".join(str(cell) for cell in df.iloc[i].tolist()).lower()
        if "in millions" in row_str:
            scale = 1_000_000
            break
        elif "in billions" in row_str:
            scale = 1_000_000_000
            break

    for _, row in df.iterrows():
        row_str = " ".join(str(cell) for cell in row.tolist())
        if rev is None and re.search(r"(^|\s)revenue($|\s)", row_str, re.IGNORECASE):
            for cell in row.tolist()[1:]:
                s = str(cell).replace(",", "").replace("$", "").replace("(", "-").replace(")", "")
                try:
                    rev_val = float(s)
                    rev = rev_val * scale
                    break
                except Exception:
                    continue
        if ni is None and re.search(r"(^|\s)net\s+income($|\s)", row_str, re.IGNORECASE):
            for cell in row.tolist()[1:]:
                s = str(cell).replace(",", "").replace("$", "").replace("(", "-").replace(")", "")
                try:
                    ni_val = float(s)
                    ni = ni_val * scale
                    break
                except Exception:
                    continue
        if rev is not None and ni is not None:
            break
    return rev, ni


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Attempt to parse Total Debt & Cash from a Balance Sheet DataFrame
# ────────────────────────────────────────────────────────────────────────────────
def parse_balance_sheet_from_df(df: pd.DataFrame) -> tuple[float | None, float | None]:
    """
    Searches df for 'Total Debt' and 'Cash and cash equivalents' (case-insensitive).
    Returns (total_debt, cash_ce) in USD.
    Detects scale by scanning up to first 5 rows.
    """
    total_debt = None
    cash_ce = None

    # Determine scale by scanning up to first 5 rows
    scale = 1
    header_rows = min(5, len(df))
    for i in range(header_rows):
        row_str = " ".join(str(cell) for cell in df.iloc[i].tolist()).lower()
        if "in millions" in row_str:
            scale = 1_000_000
            break
        elif "in billions" in row_str:
            scale = 1_000_000_000
            break

    for _, row in df.iterrows():
        row_str = " ".join(str(cell) for cell in row.tolist())
        if total_debt is None and re.search(r"total\s+debt", row_str, re.IGNORECASE):
            for cell in row.tolist()[1:]:
                s = str(cell).replace(",", "").replace("$", "").replace("(", "-").replace(")", "")
                try:
                    td_val = float(s)
                    total_debt = td_val * scale
                    break
                except Exception:
                    continue
        if cash_ce is None and re.search(r"cash\s+and\s+cash\s+equivalents", row_str, re.IGNORECASE):
            for cell in row.tolist()[1:]:
                s = str(cell).replace(",", "").replace("$", "").replace("(", "-").replace(")", "")
                try:
                    c_val = float(s)
                    cash_ce = c_val * scale
                    break
                except Exception:
                    continue
        if total_debt is not None and cash_ce is not None:
            break
    return total_debt, cash_ce


# ────────────────────────────────────────────────────────────────────────────────
# CORE HELPER: Extract financials (year, revenue, net_income, fcf, debt, cash) from PDF or DOCX
# ────────────────────────────────────────────────────────────────────────────────
def extract_financials_from_file(file_path: str) -> dict | None:
    """
    Attempts to extract {year, revenue, net_income, free_cash_flow, total_debt, cash_ce} 
    from a 10-K/10-Q PDF or a DOCX.
    1. If PDF:
       a) Extract full text → find year via "Year Ended <Mon> <Day>, <Year>"
       b) Find pages with "Cash Flow" → extract tables → parse FCF
       c) Find pages with "Income Statement" → extract tables → parse revenue, net income
       d) Find pages with "Balance Sheet" → extract tables → parse total_debt & cash_ce
       e) Fallback: search lines for "$" amounts beside "Free Cash Flow", etc.
    2. If DOCX:
       a) Extract full text → same line-based search for year, FCF, revenue, net income, debt, cash
    Returns a dict {
      "year": int,
      "revenue": float,
      "net_income": float,
      "free_cash_flow": float,
      "total_debt": float,
      "cash_ce": float
    } or None.
    """
    _, ext = os.path.splitext(file_path.lower())
    full_text = ""
    year = None
    revenue = None
    net_income = None
    fcf = None
    total_debt = None
    cash_ce = None

    # ────────── DEBUG ────────────────────────────────────────────────────────────
    print(f"[EXTRACT DEBUG] Starting extraction for: {file_path!r} (ext={ext})")
    # ─────────────────────────────────────────────────────────────────────────────

    # 1) Extract raw text
    if ext == ".pdf":
        full_text = extract_text_from_pdf(file_path)
    elif ext == ".docx":
        full_text = extract_text_from_docx(file_path)
    else:
        return None

    if not full_text.strip():
        print(f"[EXTRACT DEBUG] Unsupported extension or empty text: {file_path!r}")
        return None

    print(f"[EXTRACT DEBUG]   raw_text[0:300]: {repr(full_text[:300].replace(chr(10), ' '))} …")

    # 2) Find year via common patterns
    ymatches = re.findall(r"Year\s+Ended\s+[A-Za-z]+\s+\d{1,2},\s*(\d{4})", full_text, re.IGNORECASE)
    if ymatches:
        try:
            year = int(sorted({int(y) for y in ymatches})[-1])
        except Exception:
            year = None
    if year is None:
        y2 = re.findall(r"(\d{4})\s+Consolidated\s+Balance\s+Sheets", full_text, re.IGNORECASE)
        if y2:
            try:
                year = int(sorted({int(y) for y in y2})[-1])
            except Exception:
                year = None

    # 3) If PDF, attempt table extraction first
    if ext == ".pdf":
        # a) Cash Flow → parse FCF
        cf_pages = find_pages_with_keyword(
            file_path,
            r"Consolidated\s+Statements\s+of\s+Cash\s+Flows|Cash\s+Flow"
        )
        tables_cf = extract_tables_from_pdf(file_path, cf_pages)
        for df in tables_cf:
            val = parse_fcf_from_df(df)
            if val is not None:
                fcf = val
                break

        # b) Income Statement → parse revenue, net_income
        inc_pages = find_pages_with_keyword(
            file_path,
            r"Consolidated\s+Statements\s+of\s+Income|Consolidated\s+Statements\s+of\s+Comprehensive\s+Income|Income\s+Statement"
        )
        tables_inc = extract_tables_from_pdf(file_path, inc_pages)
        for df in tables_inc:
            r, ni = parse_income_from_df(df)
            if r is not None:
                revenue = r
            if ni is not None:
                net_income = ni
            if revenue is not None and net_income is not None:
                break

        # c) Balance Sheet → parse total_debt, cash_ce
        bs_pages = find_pages_with_keyword(
            file_path,
            r"Consolidated\s+Balance\s+Sheets|Balance\s+Sheet"
        )
        tables_bs = extract_tables_from_pdf(file_path, bs_pages)
        for df in tables_bs:
            td, c = parse_balance_sheet_from_df(df)
            if td is not None:
                total_debt = td
            if c is not None:
                cash_ce = c
            if total_debt is not None and cash_ce is not None:
                break

    # 4) Fallback line-by-line search if any value still missing
    if fcf is None or revenue is None or net_income is None or total_debt is None or cash_ce is None:
        for line in full_text.splitlines():
            # Free Cash Flow
            if fcf is None and re.search(r"free\s+cash\s+flow", line, re.IGNORECASE):
                m = re.search(r"\$[\s,]*([\d,]+(?:\.\d+)?)", line)
                if m:
                    try:
                        fcf_val = float(m.group(1).replace(",", ""))
                        fcf = fcf_val  # assume already absolute if no scale indicator
                    except:
                        pass

            # Revenue
            if revenue is None and re.search(r"(^|\s)revenue($|\s)", line, re.IGNORECASE):
                m = re.search(r"\$[\s,]*([\d,]+(?:\.\d+)?)", line)
                if m:
                    try:
                        rev_val = float(m.group(1).replace(",", ""))
                        revenue = rev_val
                    except:
                        pass

            # Net Income
            if net_income is None and re.search(r"(^|\s)net\s+income($|\s)", line, re.IGNORECASE):
                m = re.search(r"\$[\s,]*([\d,]+(?:\.\d+)?)", line)
                if m:
                    try:
                        ni_val = float(m.group(1).replace(",", ""))
                        net_income = ni_val
                    except:
                        pass

            # Total Debt
            if total_debt is None and re.search(r"(^|\s)total\s+debt($|\s)", line, re.IGNORECASE):
                m = re.search(r"\$[\s,]*([\d,]+(?:\.\d+)?)", line)
                if m:
                    try:
                        td_val = float(m.group(1).replace(",", ""))
                        total_debt = td_val
                    except:
                        pass

            # Cash & Cash Equivalents
            if cash_ce is None and re.search(r"cash\s+and\s+cash\s+equivalents", line, re.IGNORECASE):
                m = re.search(r"\$[\s,]*([\d,]+(?:\.\d+)?)", line)
                if m:
                    try:
                        c_val = float(m.group(1).replace(",", ""))
                        cash_ce = c_val
                    except:
                        pass

            if fcf is not None and revenue is not None and net_income is not None and total_debt is not None and cash_ce is not None:
                break

    # 5) If still missing everything, return None
    if year is None and fcf is None and revenue is None and net_income is None and total_debt is None and cash_ce is None:
        return None

    return {
        "year": year,
        "revenue": revenue,
        "net_income": net_income,
        "free_cash_flow": fcf,
        "total_debt": total_debt,
        "cash_ce": cash_ce
    }


# ────────────────────────────────────────────────────────────────────────────────
# MAIN FUNCTION: run_dcf_model (company_name-based)
# ────────────────────────────────────────────────────────────────────────────────
def run_dcf_model(
    company_name: str,
    assumptions: dict,
    file_paths: list[str]
) -> tuple[str, dict]:
    """
    Performs a 3-case (Bull / Base / Bear) FCFF-based DCF valuation for the given company_name.
    1) Uses DeepSeek to convert company_name → ticker.
    2) Extracts historical financials from each uploaded file via extract_financials_from_file.
    3) Aggregates at most 5 most recent years of financials.
    4) Uses user-supplied assumption dict (or defaults) for growth_rates, waccs, terminal_multiples.
    5) Projects Free Cash Flow to Firm (FCFF) for 5 years + terminal value → discounts at scenario WACC → computes Enterprise Value.
    6) Subtracts Net Debt → divides by Shares Outstanding → gets Equity Value per share for each scenario.
    7) Writes an Excel under ./reports named "<TICKER>_DCF_<timestamp>.xlsx".
    8) Returns (output_path, summary_dict).
    """
    # ────────────────────────────────────────────────────────────────────────────
    # 1) Convert company_name → ticker via DeepSeek
    # ────────────────────────────────────────────────────────────────────────────
    ticker_prompt = (
        f"You are a stock ticker lookup assistant.\n"
        f"For the company name \"{company_name}\", provide the primary U.S. stock ticker (e.g. 'AAPL').\n"
        f"Reply with exactly the ticker symbol (no extra text)."
    )
    raw_ticker = deepseek_chat(ticker_prompt, max_tokens=32).strip()
    ticker = raw_ticker.upper()

    # ────────────────────────────────────────────────────────────────────────────
    # 1.5) Fetch shares outstanding & current share price
    # ────────────────────────────────────────────────────────────────────────────
    shares_outstanding = get_shares_outstanding(ticker)
    current_price = get_current_share_price(ticker)

    # ────────────────────────────────────────────────────────────────────────────
    # 2) Extract pandas dicts for each file (year, revenue, net_income, free_cash_flow, debt, cash)
    # ────────────────────────────────────────────────────────────────────────────
    extracted = []
    for path in file_paths:
        fin = extract_financials_from_file(path)
        # Require at least a year and an FCF (so we can project)
        if fin is not None and fin.get("year") is not None and fin.get("free_cash_flow") is not None:
            extracted.append(fin)
    if not extracted:
        raise ValueError("No valid financial data could be extracted from uploaded files.")

    # Keep only 5 most recent years
    extracted = sorted(extracted, key=lambda x: x["year"] or 0, reverse=True)[:5]
    historical_list = sorted(extracted, key=lambda x: x["year"] or 0)  # ascending

    # Extract arrays
    hist_years = [e["year"] for e in historical_list]
    hist_fcfs = [e["free_cash_flow"] for e in historical_list]
    hist_revs = [e.get("revenue") for e in historical_list]
    hist_nis = [e.get("net_income") for e in historical_list]
    # For debt & cash, take the most recent (largest year) if available
    most_recent = historical_list[-1]
    total_debt = most_recent.get("total_debt", 0) or 0
    cash_ce = most_recent.get("cash_ce", 0) or 0
    net_debt = total_debt - cash_ce

    last_year = hist_years[-1]
    last_fcf = hist_fcfs[-1]

    # ────────────────────────────────────────────────────────────────────────────
    # 2.5) FCF VALIDATION & FALLBACK
    # ────────────────────────────────────────────────────────────────────────────
    # If the extracted FCF is implausibly low (e.g., < $500M for a large-cap), fetch from yfinance.
    # We assume that if last_fcf < 500 million but company is large, it’s wrong.
    if last_fcf is None or last_fcf < 500_000_000:
        print(f"[WARNING] Extracted FCF (${last_fcf}) too low for {ticker}. Attempting fallback via yfinance.")
        yf_fcf = get_yfinance_fcf(ticker)
        if yf_fcf is not None and yf_fcf > 0:
            last_fcf = yf_fcf
            print(f"[INFO] Using fallback FCF from yfinance: ${last_fcf:,.0f}")
        else:
            # If yfinance fails, let user know, and still proceed with the low number (risking inaccuracy)
            print(f"[WARNING] yfinance FCF lookup failed or returned None. Continuing with extracted FCF = ${last_fcf}")

    print(f"[DEBUG] Base Free Cash Flow used for projection: ${last_fcf:,.0f} (Fiscal Year {last_year})")
    print(f"[DEBUG] Net Debt used: ${net_debt:,.0f}")

    # ────────────────────────────────────────────────────────────────────────────
    # 3) Scenario assumptions: growth_rates, waccs, terminal_multiples
    # ────────────────────────────────────────────────────────────────────────────
    # 3a) Growth Rates
    if "growth_rates" in assumptions:
        growth_rates = {
            "bull": assumptions["growth_rates"].get("bull"),
            "base": assumptions["growth_rates"].get("base"),
            "bear": assumptions["growth_rates"].get("bear"),
        }
    else:
        hist_str = ", ".join(f"{e['year']}:{e['free_cash_flow']:,}" for e in historical_list)
        gr_prompt = (
            f"For ticker {ticker}, historical Free Cash Flows to Firm (USD) for years: {hist_str}.\n"
            f"Current share price: ${current_price:.2f} per share.\n"
            f"Net Debt (Debt - Cash): ${net_debt:,.2f}.\n"
            f"Propose forward 5-year annual FCFF growth rates under bull, base, bear,\n"
            f"such that the resulting DCF per share is in the same ballpark as the current price.\n"
            f"Reply with JSON {{\"bull\":0.XX,\"base\":0.XX,\"bear\":0.XX}}."
        )
        raw_gr = deepseek_chat(gr_prompt, max_tokens=256)
        try:
            growth_rates = json.loads(raw_gr)
        except Exception:
            print("[WARNING] Failed to parse growth rates from DeepSeek. Using defaults 8%,5%,2%.")
            growth_rates = {"base": 0.05, "bull": 0.08, "bear": 0.02}

    for scenario in ["bull", "base", "bear"]:
        if growth_rates.get(scenario) is None:
            growth_rates[scenario] = growth_rates.get("base", 0.05)

    # 3b) WACCs and Terminal Multiples
    if "waccs" in assumptions and "terminal_multiples" in assumptions:
        waccs = {
            "bull": assumptions["waccs"].get("bull"),
            "base": assumptions["waccs"].get("base"),
            "bear": assumptions["waccs"].get("bear"),
        }
        terminal_mults = {
            "bull": assumptions["terminal_multiples"].get("bull"),
            "base": assumptions["terminal_multiples"].get("base"),
            "bear": assumptions["terminal_multiples"].get("bear"),
        }
    else:
        recent_three = historical_list[-3:]
        three_str = ", ".join(f"{e['year']}:{e['free_cash_flow']:,}" for e in recent_three)
        wa_prompt = (
            f"For ticker {ticker}, last three years’ FCFF (USD): {three_str}.\n"
            f"Current share price: ${current_price:.2f} per share.\n"
            f"Net Debt: ${net_debt:,.2f}.\n"
            f"Provide WACC and terminal multiples under bull, base, bear,\n"
            f"such that the 5-year DCF per share remains near the market price.\n"
            f"Reply JSON {{"
            f"\"bull\":{{\"wacc\":0.XX,\"terminal_multiple\":YY}}, "
            f"\"base\":{{\"wacc\":0.XX,\"terminal_multiple\":YY}}, "
            f"\"bear\":{{\"wacc\":0.XX,\"terminal_multiple\":YY}}}}."
        )
        raw_wa = deepseek_chat(wa_prompt, max_tokens=256)
        try:
            wa_data = json.loads(raw_wa)
            waccs = {
                "bull": wa_data["bull"]["wacc"],
                "base": wa_data["base"]["wacc"],
                "bear": wa_data["bear"]["wacc"],
            }
            terminal_mults = {
                "bull": wa_data["bull"]["terminal_multiple"],
                "base": wa_data["base"]["terminal_multiple"],
                "bear": wa_data["bear"]["terminal_multiple"],
            }
        except Exception:
            print("[WARNING] Failed to parse WACC/terminal multiples from DeepSeek. Using defaults.")
            waccs = {"bull": 0.09, "base": 0.10, "bear": 0.11}
            terminal_mults = {"bull": 14, "base": 12, "bear": 10}

    for scenario in ["bull", "base", "bear"]:
        if waccs.get(scenario) is None:
            waccs[scenario] = waccs.get("base", 0.10)
        if terminal_mults.get(scenario) is None:
            terminal_mults[scenario] = terminal_mults.get("base", 12)

    # ────────────────────────────────────────────────────────────────────────────
    # 4) PROJECT FCFF & DISCOUNT for EACH SCENARIO
    # ────────────────────────────────────────────────────────────────────────────
    dcf_results = {}
    for scenario in ["bull", "base", "bear"]:
        gr = growth_rates[scenario]
        wacc = waccs[scenario]
        tm = terminal_mults[scenario]

        # Project 5 years of FCFF (starting from last_fcf)
        projections = []
        for i in range(1, 6):
            year_i = last_year + i
            fcf_i = last_fcf * ((1 + gr) ** i)
            projections.append({"year": year_i, "fcf": fcf_i})

        # Terminal value at year 5 (using last projected FCFF)
        terminal_value = projections[-1]["fcf"] * tm

        # Discount each year's FCFF + terminal → Enterprise Value (PV)
        ev = sum(
            proj["fcf"] / ((1 + wacc) ** j)
            for j, proj in enumerate(projections, start=1)
        )
        ev += terminal_value / ((1 + wacc) ** 5)

        # Subtract Net Debt → Equity Value
        equity_value = ev - net_debt

        # Per-share
        npv_per_share = None
        if shares_outstanding and shares_outstanding > 0:
            npv_per_share = equity_value / shares_outstanding

        dcf_results[scenario] = {
            "growth_rate": gr,
            "wacc": wacc,
            "terminal_multiple": tm,
            "projections": projections,
            "terminal_value": terminal_value,
            "enterprise_value": ev,
            "equity_value": equity_value,
            "npv_per_share": npv_per_share
        }

    # ────────────────────────────────────────────────────────────────────────────
    # 5) CREATE EXCEL WORKBOOK UNDER ./reports
    # ────────────────────────────────────────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    for scenario in ["bull", "base", "bear"]:
        data = dcf_results[scenario]
        ws = wb.create_sheet(f"{scenario.capitalize()} Case")

        # Header info
        ws.append(["Scenario", scenario.capitalize()])
        ws.append(["Assumptions", "Value"])
        ws.append(["Growth Rate", data["growth_rate"]])
        ws.append(["WACC", data["wacc"]])
        ws.append(["Terminal Multiple", data["terminal_multiple"]])
        ws.append([])

        # Projections
        ws.append(["Year", "Projected FCFF (USD)"])
        for proj in data["projections"]:
            ws.append([proj["year"], proj["fcf"]])
        ws.append([])

        # Terminal & EV
        ws.append(["Terminal Value (Year 5)", data["terminal_value"]])
        ws.append([])
        ws.append(["Enterprise Value (PV of FCFF + Terminal)", data["enterprise_value"]])
        ws.append([])

        # Net Debt & Equity Value
        ws.append(["Net Debt (Debt - Cash)", net_debt])
        ws.append(["Equity Value", data["equity_value"]])
        ws.append([])

        # Per-share
        if data["npv_per_share"] is not None:
            ws.append(["NPV per Share", data["npv_per_share"]])

    reports_dir = os.path.join(os.getcwd(), "reports")
    os.makedirs(reports_dir, exist_ok=True)
    filename = f"{ticker}_DCF_{int(time.time())}.xlsx"
    output_path = os.path.join(reports_dir, filename)
    wb.save(output_path)

    # ────────────────────────────────────────────────────────────────────────────
    # 6) BUILD SUMMARY DICT
    # ────────────────────────────────────────────────────────────────────────────
    summary_dict = {
        "company_name": company_name,
        "ticker": ticker,
        "current_share_price": current_price,
        "shares_outstanding": shares_outstanding,
        "net_debt": net_debt,
        "historical_years": hist_years,
        "historical_fcfs": hist_fcfs,
        "historical_revenues": hist_revs,
        "historical_net_incomes": hist_nis,
        "scenarios": {
            scenario: {
                # Include an "npv" key so your existing JS (which does sc.npv) still works:
                "npv": dcf_results[scenario]["equity_value"],

                # Keep the other new fields in case you want to reference them later:
                "enterprise_value": dcf_results[scenario]["enterprise_value"],
                "equity_value": dcf_results[scenario]["equity_value"],
                "npv_per_share": dcf_results[scenario]["npv_per_share"],
                "terminal_value": dcf_results[scenario]["terminal_value"],
                "wacc": dcf_results[scenario]["wacc"],
                "growth_rate": dcf_results[scenario]["growth_rate"],
                "terminal_multiple": dcf_results[scenario]["terminal_multiple"]
            }
            for scenario in ["bull", "base", "bear"]
        }
    }

    # ────────────────────────────────────────────────────────────────────────────
    # 7) Add plain-English “reasoning” paragraphs for each scenario
    # ────────────────────────────────────────────────────────────────────────────
    reasonings = {}
    for scenario in ["bull", "base", "bear"]:
        data = dcf_results[scenario]
        gr = data["growth_rate"]
        wacc = data["wacc"]
        tm = data["terminal_multiple"]
        ev_val = data["enterprise_value"]
        eq_val = data["equity_value"]
        per_share = data["npv_per_share"]

        if scenario == "bull":
            label = "Bull Case"
        elif scenario == "base":
            label = "Base Case"
        else:
            label = "Bear Case"

        reason_str = (
            f"<strong>{label}:</strong> We assumed an annual FCFF growth rate of "
            f"{gr*100:.1f}%, a discount rate (WACC) of {wacc*100:.1f}%, "
            f"and a terminal multiple of {tm}×. Under these assumptions, the present value of projected FCFF "
            f"and terminal value is ${ev_val:,.2f}; after subtracting net debt of ${net_debt:,.2f}, "
            f"the implied equity value is ${eq_val:,.2f}, or ${per_share:,.2f} per share."
        )
        reasonings[scenario] = reason_str

    summary_dict["reasonings"] = reasonings

    return output_path, summary_dict
