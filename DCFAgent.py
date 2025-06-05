# DCFAgent.py

import os
import re
import time
import json
import requests

import pandas as pd
import numpy as np
import yfinance as yf

from PyPDF2 import PdfReader
import camelot
from openpyxl import Workbook # Not used in the provided snippet, kept for completeness
from docx import Document
from bs4 import BeautifulSoup

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
    Adjusted temperature for more deterministic numerical outputs.
    """
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": max_tokens,
        "temperature": temperature # Lower temperature for factual extraction
    }
    try:
        resp = requests.post(DEEPSEEK_CHAT_URL, headers=headers, json=payload, timeout=90) # Increased timeout
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
# HELPER: Fetch shares outstanding & current price via yfinance
# ────────────────────────────────────────────────────────────────────────────────
def get_shares_outstanding(ticker: str) -> int | None:
    try:
        info = yf.Ticker(ticker).info
        return info.get("sharesOutstanding")
    except Exception as e:
        print(f"⚠️ Warning: Could not fetch shares outstanding for {ticker} from yfinance: {e}")
        return None


def get_current_share_price(ticker: str) -> float | None:
    try:
        info = yf.Ticker(ticker).info
        return info.get("regularMarketPrice")
    except Exception as e:
        print(f"⚠️ Warning: Could not fetch current share price for {ticker} from yfinance: {e}")
        return None


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Fetch “official” net debt via yfinance
# ────────────────────────────────────────────────────────────────────────────────
def get_yfinance_net_debt(ticker: str) -> float | None:
    """
    Attempts to fetch totalDebt and cash from yfinance Ticker.info and returns netDebt = totalDebt - cash.
    """
    try:
        info = yf.Ticker(ticker).info
        total_debt = info.get("totalDebt")
        cash = info.get("cash")
        if total_debt is not None and cash is not None:
            return float(total_debt) - float(cash)
    except Exception as e:
        print(f"⚠️ Warning: Could not fetch net debt for {ticker} from yfinance: {e}")
    return None


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Fetch fallback FCF from yfinance “cashflow” DataFrame
# ────────────────────────────────────────────────────────────────────────────────
def get_yfinance_fcf(ticker: str) -> float | None:
    """
    Pulls most recent Free Cash Flow from yfinance cashflow statement (annual).
    """
    try:
        cf = yf.Ticker(ticker).cashflow
        if cf is None or cf.empty:
            print(f"⚠️ Warning: yfinance cashflow data not available for {ticker}.")
            return None
        # yfinance columns are dates; first column is most recent
        most_recent_col = cf.columns[0]
        for label in ("Free Cash Flow", "FreeCashFlow"): # Check both common labels
            if label in cf.index:
                fcf_val = cf.loc[label, most_recent_col]
                return float(fcf_val)
    except Exception as e:
        print(f"⚠️ Warning: Could not fetch FCF for {ticker} from yfinance: {e}")
        return None
    return None


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
    pages = set() # Use set to avoid duplicate page numbers
    try:
        reader = PdfReader(pdf_path)
        for idx, page in enumerate(reader.pages):
            txt = page.extract_text() or ""
            # Use re.IGNORECASE for case-insensitivity
            if re.search(keyword, txt, re.IGNORECASE):
                pages.add(idx) # Store 0-indexed page number
    except Exception as e:
        print(f"⚠️ Warning: Error finding keyword '{keyword}' in PDF {pdf_path}: {e}")
    return sorted(list(pages)) # Return sorted 0-indexed page numbers


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Extract tables from PDF pages
# ────────────────────────────────────────────────────────────────────────────────
def extract_tables_from_pdf(pdf_path: str, page_indices: list[int]) -> list[pd.DataFrame]:
    dfs = []
    if not page_indices:
        return dfs
    # camelot expects 1-indexed page numbers
    pages_str = ",".join(str(i + 1) for i in page_indices)
    try:
        # First attempt “lattice”; if that fails, fall back to “stream”
        tables = camelot.read_pdf(pdf_path, pages=pages_str, flavor="lattice", split_text=True, line_scale=40)
        if not tables or all(len(tbl.df)==0 for tbl in tables):
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
    """Safely converts a string to a float, handling commas, dollar signs, and parentheses for negatives."""
    if not isinstance(s, str):
        return None
    s = s.strip().replace(",", "").replace("$", "")
    if s.startswith("(") and s.endswith(")"): # Handle (123.45) for negative
        s = "-" + s[1:-1]
    try:
        return float(s)
    except ValueError:
        return None

# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Parse Free Cash Flow from a DataFrame with better scale detection
# ────────────────────────────────────────────────────────────────────────────────
def parse_fcf_from_df(df: pd.DataFrame) -> float | None:
    """
    Searches df for a row containing 'Free Cash Flow' (case-insensitive)
    and returns the most recent numeric value in that row (USD).
    Detects if table headers mention 'in millions' or 'in billions' anywhere in the first 5 rows.
    """
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

    fcf_row = None
    for _, row in df.iterrows():
        # Check if the first cell or any cell in the row contains the keyword
        row_cells_str = " ".join(str(cell) for cell in row.tolist())
        if re.search(r"free\s+cash\s+flow|fcf", row_cells_str, re.IGNORECASE):
            fcf_row = row
            break

    if fcf_row is None:
        return None

    # Try to find the most recent year's data by looking for numeric columns, prioritizing the rightmost
    for col_idx in reversed(range(len(fcf_row))):
        val = parse_numeric_value(str(fcf_row.iloc[col_idx]))
        if val is not None:
            return val * scale
    return None


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Parse Revenue & Net Income from Income Statement DataFrame
# ────────────────────────────────────────────────────────────────────────────────
def parse_income_from_df(df: pd.DataFrame) -> tuple[float | None, float | None]:
    rev = None
    ni = None

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
        row_cells_str = " ".join(str(cell) for cell in row.tolist())

        if rev is None and re.search(r"(^|\s)(total\s+)?revenue($|\s)|(^|\s)sales($|\s)", row_cells_str, re.IGNORECASE):
            for cell in reversed(row.tolist()): # Prioritize most recent
                val = parse_numeric_value(str(cell))
                if val is not None:
                    rev = val * scale
                    break

        if ni is None and re.search(r"(^|\s)net\s+income($|\s)|(^|\s)net\s+earnings($|\s)|(^|\s)profit\s+for\s+the\s+period($|\s)", row_cells_str, re.IGNORECASE):
            for cell in reversed(row.tolist()): # Prioritize most recent
                val = parse_numeric_value(str(cell))
                if val is not None:
                    ni = val * scale
                    break
        if rev is not None and ni is not None:
            break

    return rev, ni


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Parse Total Debt & Cash from Balance Sheet DataFrame
# ────────────────────────────────────────────────────────────────────────────────
def parse_balance_sheet_from_df(df: pd.DataFrame) -> tuple[float | None, float | None]:
    total_debt = None
    cash_ce = None

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
        row_cells_str = " ".join(str(cell) for cell in row.tolist())

        if total_debt is None and re.search(r"(^|\s)total\s+debt($|\s)|long-term\s+debt", row_cells_str, re.IGNORECASE):
            for cell in reversed(row.tolist()): # Prioritize most recent
                val = parse_numeric_value(str(cell))
                if val is not None:
                    total_debt = val * scale
                    break

        if cash_ce is None and re.search(r"cash\s+and\s+cash\s+equivalents|cash\s+at\s+bank", row_cells_str, re.IGNORECASE):
            for cell in reversed(row.tolist()): # Prioritize most recent
                val = parse_numeric_value(str(cell))
                if val is not None:
                    cash_ce = val * scale
                    break
        if total_debt is not None and cash_ce is not None:
            break

    return total_debt, cash_ce


# ────────────────────────────────────────────────────────────────────────────────
# CORE HELPER: Extract financials (year, revenue, net_income, fcf, debt, cash) from PDF or DOCX
# ────────────────────────────────────────────────────────────────────────────────
def extract_financials_from_file(file_path: str, ticker: str = "") -> dict | None:
    """
    Attempts to extract {year, revenue, net_income, free_cash_flow, total_debt, cash_ce}
    from a 10-K/10-Q PDF or a DOCX.
    """
    _, ext = os.path.splitext(file_path.lower())
    full_text = ""
    year = None
    revenue = None
    net_income = None
    fcf = None
    total_debt = None
    cash_ce = None

    print(f"[EXTRACT DEBUG] Starting extraction for: {file_path!r} (ext={ext})")

    if ext == ".pdf":
        full_text = extract_text_from_pdf(file_path)
    elif ext == ".docx":
        full_text = extract_text_from_docx(file_path)
    else:
        print(f"[EXTRACT DEBUG] Unsupported file type: {ext}. Skipping.")
        return None

    if not full_text.strip():
        print(f"[EXTRACT DEBUG] No text found in {file_path}. Skipping.")
        return None

    # 1) Find year via "Year Ended ... <Year>" or "Fiscal Year"
    ymatches = re.findall(r"(?:Year\s+Ended|Fiscal\s+Year)\s+[A-Za-z]+\s+\d{1,2},\s*(\d{4})|(\d{4})\s+Consolidated\s+Balance\s+Sheets", full_text, re.IGNORECASE)
    # ymatches will be a list of tuples like (year1_group1, year1_group2)
    potential_years = []
    for match in ymatches:
        for y_str in match:
            if y_str: # Only take non-empty groups
                try:
                    potential_years.append(int(y_str))
                except ValueError:
                    pass
    if potential_years:
        year = sorted(list(set(potential_years)), reverse=True)[0] # Get most recent unique year

    print(f"[EXTRACT DEBUG] Identified year: {year}")

    # 2) If PDF, attempt table extraction (more robust than line-by-line)
    if ext == ".pdf":
        # a) Cash Flow → parse FCF
        cf_pages = find_pages_with_keyword(
            file_path,
            r"Consolidated\s+Statements\s+of\s+Cash\s+Flows|Cash\s+Flows|Statements\s+of\s+Cash\s+Flow"
        )
        tables_cf = extract_tables_from_pdf(file_path, cf_pages)
        for df in tables_cf:
            val = parse_fcf_from_df(df)
            if val is not None:
                fcf = float(val)
                print(f"[EXTRACT DEBUG] FCF found from table (Cash Flow): ${fcf:,.0f}")
                break

        # b) Income Statement → parse revenue, net income
        inc_pages = find_pages_with_keyword(
            file_path,
            r"Consolidated\s+Statements\s+of\s+Income|Consolidated\s+Statements\s+of\s+Operations|Income\s+Statement"
        )
        tables_inc = extract_tables_from_pdf(file_path, inc_pages)
        for df in tables_inc:
            r, ni = parse_income_from_df(df)
            if r is not None:
                revenue = float(r)
                print(f"[EXTRACT DEBUG] Revenue found from table (Income): ${revenue:,.0f}")
            if ni is not None:
                net_income = float(ni)
                print(f"[EXTRACT DEBUG] Net Income found from table (Income): ${net_income:,.0f}")
            if revenue is not None and net_income is not None:
                break

        # c) Balance Sheet → parse total debt, cash & equivalents
        bs_pages = find_pages_with_keyword(
            file_path,
            r"Consolidated\s+Balance\s+Sheets|Balance\s+Sheet"
        )
        tables_bs = extract_tables_from_pdf(file_path, bs_pages)
        for df in tables_bs:
            td, c = parse_balance_sheet_from_df(df)
            if td is not None:
                total_debt = float(td)
                print(f"[EXTRACT DEBUG] Total Debt found from table (Balance Sheet): ${total_debt:,.0f}")
            if c is not None:
                cash_ce = float(c)
                print(f"[EXTRACT DEBUG] Cash & Equivalents found from table (Balance Sheet): ${cash_ce:,.0f}")
            if total_debt is not None and cash_ce is not None:
                break

    # 3) Fallback line-by-line search for any missing items (less reliable, but better than nothing)
    # The regex here will be very basic, assuming a "$1,234,567" format near the keyword
    print("[EXTRACT DEBUG] Attempting line-by-line fallback for missing data...")
    for line in full_text.splitlines():
        line = line.strip()
        # Free Cash Flow
        if fcf is None and re.search(r"free\s+cash\s+flow|fcf", line, re.IGNORECASE):
            m = re.search(r"\$?\s*([\d,\(]+\.?\d*)\s*(million|billion|M|B)?", line, re.IGNORECASE)
            if m:
                val = parse_numeric_value(m.group(1))
                if val is not None:
                    scale_modifier = 1
                    if m.group(2):
                        if m.group(2).lower() in ("million", "m"):
                            scale_modifier = 1_000_000
                        elif m.group(2).lower() in ("billion", "b"):
                            scale_modifier = 1_000_000_000
                    fcf = val * scale_modifier
                    print(f"[EXTRACT DEBUG] FCF found via regex fallback: ${fcf:,.0f}")


        # Revenue
        if revenue is None and re.search(r"(^|\s)(total\s+)?revenue($|\s)|(^|\s)sales($|\s)", line, re.IGNORECASE):
            m = re.search(r"\$?\s*([\d,\(]+\.?\d*)\s*(million|billion|M|B)?", line, re.IGNORECASE)
            if m:
                val = parse_numeric_value(m.group(1))
                if val is not None:
                    scale_modifier = 1
                    if m.group(2):
                        if m.group(2).lower() in ("million", "m"):
                            scale_modifier = 1_000_000
                        elif m.group(2).lower() in ("billion", "b"):
                            scale_modifier = 1_000_000_000
                    revenue = val * scale_modifier
                    print(f"[EXTRACT DEBUG] Revenue found via regex fallback: ${revenue:,.0f}")

        # Net Income
        if net_income is None and re.search(r"(^|\s)net\s+income($|\s)|(^|\s)net\s+earnings($|\s)|(^|\s)profit\s+for\s+the\s+period($|\s)", line, re.IGNORECASE):
            m = re.search(r"\$?\s*([\d,\(]+\.?\d*)\s*(million|billion|M|B)?", line, re.IGNORECASE)
            if m:
                val = parse_numeric_value(m.group(1))
                if val is not None:
                    scale_modifier = 1
                    if m.group(2):
                        if m.group(2).lower() in ("million", "m"):
                            scale_modifier = 1_000_000
                        elif m.group(2).lower() in ("billion", "b"):
                            scale_modifier = 1_000_000_000
                    net_income = val * scale_modifier
                    print(f"[EXTRACT DEBUG] Net Income found via regex fallback: ${net_income:,.0f}")

        # Total Debt
        if total_debt is None and re.search(r"(^|\s)total\s+debt($|\s)|long-term\s+debt", line, re.IGNORECASE):
            m = re.search(r"\$?\s*([\d,\(]+\.?\d*)\s*(million|billion|M|B)?", line, re.IGNORECASE)
            if m:
                val = parse_numeric_value(m.group(1))
                if val is not None:
                    scale_modifier = 1
                    if m.group(2):
                        if m.group(2).lower() in ("million", "m"):
                            scale_modifier = 1_000_000
                        elif m.group(2).lower() in ("billion", "b"):
                            scale_modifier = 1_000_000_000
                    total_debt = val * scale_modifier
                    print(f"[EXTRACT DEBUG] Total Debt found via regex fallback: ${total_debt:,.0f}")

        # Cash & Equivalents
        if cash_ce is None and re.search(r"cash\s+and\s+cash\s+equivalents|cash\s+at\s+bank", line, re.IGNORECASE):
            m = re.search(r"\$?\s*([\d,\(]+\.?\d*)\s*(million|billion|M|B)?", line, re.IGNORECASE)
            if m:
                val = parse_numeric_value(m.group(1))
                if val is not None:
                    scale_modifier = 1
                    if m.group(2):
                        if m.group(2).lower() in ("million", "m"):
                            scale_modifier = 1_000_000
                        elif m.group(2).lower() in ("billion", "b"):
                            scale_modifier = 1_000_000_000
                    cash_ce = val * scale_modifier
                    print(f"[EXTRACT DEBUG] Cash & Equivalents found via regex fallback: ${cash_ce:,.0f}")

        # If everything found, no need to continue parsing lines
        if all(x is not None for x in [fcf, revenue, net_income, total_debt, cash_ce, year]):
            break

    # If FCF still missing, try yfinance as fallback
    if fcf is None and ticker: # Only try yfinance if we have a ticker
        yf_fcf = get_yfinance_fcf(ticker)
        if yf_fcf is not None:
            fcf = float(yf_fcf)
            print(f"[INFO] Using yfinance FCF fallback for {ticker}: ${fcf:,.0f}")

    # 4) If nothing found, give up (or if year is crucial and missing)
    if year is None and fcf is None and revenue is None and net_income is None and total_debt is None and cash_ce is None:
        print(f"❌ Error: No valid financial data extracted from {file_path}.")
        return None

    # Ensure all crucial fields for DCF are present or set to None explicitly
    extracted_data = {
        "year": year,
        "revenue": revenue,
        "net_income": net_income,
        "free_cash_flow": fcf,
        "total_debt": total_debt,
        "cash_ce": cash_ce
    }
    print(f"[EXTRACT DEBUG] Final extracted data for {file_path}: {extracted_data}")
    return extracted_data


# ────────────────────────────────────────────────────────────────────────────────
# 1. Segment Extraction & Summarization
# ────────────────────────────────────────────────────────────────────────────────
def find_segment_pages(pdf_path: str) -> list[int]:
    keywords = [
        r"Management['’]s\s+Discussion\s*(?:&|\band\b)\s*Analysis",
        r"\bMD&A\b",
        r"\bBusiness\s+Segments\b",
        r"\bSegment\s+Results\b",
        r"\bSegments?\b" # Broad match, might need refinement if too generic
    ]
    pages = set()
    try:
        reader = PdfReader(pdf_path)
        for idx, page in enumerate(reader.pages):
            text = (page.extract_text() or "")
            for kw in keywords:
                if re.search(kw, text, re.IGNORECASE):
                    pages.add(idx)
    except Exception as e:
        print(f"⚠️ Warning: Error finding segment keywords in {pdf_path}: {e}")
    return sorted(list(pages))


def extract_segment_tables(pdf_path: str, segment_pages: list[int]) -> list[pd.DataFrame]:
    dfs = []
    if not segment_pages:
        return dfs
    pages_str = ",".join(str(i + 1) for i in segment_pages)
    try:
        tables = camelot.read_pdf(pdf_path, pages=pages_str, flavor="lattice", split_text=True, line_scale=40)
        if not tables or all(len(tbl.df)==0 for tbl in tables):
            print(f"[EXTRACT DEBUG] No lattice tables, trying stream for segment tables on pages {pages_str}.")
            tables = camelot.read_pdf(pdf_path, pages=pages_str, flavor="stream", split_text=True, line_scale=40)
        for tbl in tables:
            if not tbl.df.empty:
                dfs.append(tbl.df)
        print(f"[EXTRACT DEBUG] Extracted {len(dfs)} segment tables from pages {pages_str} of {pdf_path}.")
    except Exception as e:
        print(f"❌ Error extracting segment tables from PDF {pdf_path} pages {pages_str}: {e}")
    return dfs


def summarize_segment_performance(dfs: list[pd.DataFrame], ticker: str) -> str:
    """
    Given a list of DataFrames (from camelot) that contain segment revenue & margin tables,
    ask DeepSeek to summarize revenue CAGRs and average margins for each segment.
    """
    if not dfs:
        return "No segment tables found."

    table_strs = []
    for i, df in enumerate(dfs):
        # Limit to first 7 rows & 7 columns for prompt brevity, to capture more data
        small = df.iloc[:7, :7].to_csv(index=False)
        table_strs.append(f"--- Table {i+1} ---\n" + small)

    prompt = (
        f"You are a financial data summarization assistant. For ticker {ticker},\n"
        f"here are extracted tables, potentially containing segment revenue, operating profit, or similar financial data.\n"
        f"Each table is in CSV format.\n"
        f"Analyze the tables and for each distinct business segment you find, provide:\n"
        f"  • The latest reported annual revenue (in USD, with original scale like 'millions' or 'billions' if available).\n"
        f"  • The compound annual growth rate (CAGR) of revenue over the most recent 3 available years (e.g., 2021-2023).\n"
        f"  • The average operating margin (Operating Profit / Revenue) over the most recent 3 available years.\n"
        f"If data for a calculation is missing or cannot be reliably extracted, state 'N/A'.\n"
        f"Return a clear, concise bullet-point summary for each segment. Do not include any introductory or concluding remarks.\n"
        "\n"
        f"TABLES:\n"
        + "\n\n".join(table_strs)
    )

    raw = deepseek_chat(prompt, max_tokens=768, temperature=0.2) # Increased max_tokens, slightly higher temp for interpretation
    return raw


# ────────────────────────────────────────────────────────────────────────────────
# 2. Management Guidance Extraction (MD&A)
# ────────────────────────────────────────────────────────────────────────────────
def find_mda_pages(pdf_path: str) -> list[int]:
    """
    Returns pages that contain 'Management’s Discussion & Analysis' or 'MD&A'.
    """
    keywords = [r"Management’s\s+Discussion\s+and\s+Analysis", r"MD&A"]
    pages = set()
    try:
        reader = PdfReader(pdf_path)
        for idx, page in enumerate(reader.pages):
            txt = page.extract_text() or ""
            for kw in keywords:
                if re.search(kw, txt, re.IGNORECASE):
                    pages.add(idx)
    except Exception as e:
        print(f"⚠️ Warning: Error finding MD&A keywords in {pdf_path}: {e}")
    return sorted(list(pages))


def extract_mda_text(pdf_path: str, mda_pages: list[int]) -> str:
    """
    Concatenate all the text from the identified MD&A pages.
    """
    if not mda_pages:
        return ""
    collected = []
    try:
        reader = PdfReader(pdf_path)
        for idx in mda_pages:
            if idx < len(reader.pages): # Ensure page index is valid
                page = reader.pages[idx]
                txt = page.extract_text() or ""
                collected.append(txt)
    except Exception as e:
        print(f"❌ Error extracting MD&A text from PDF {pdf_path} pages {mda_pages}: {e}")
    return "\n\n".join(collected)


def summarize_management_guidance(mda_text: str, ticker: str) -> dict:
    """
    Given the MD&A text, ask DeepSeek to extract forward guidance as JSON:
      {
        "next_year_overall_revenue": "...",
        "next_year_fcff": "...",
        "segment_guidance": { "Collins Aerospace": "...", "Pratt & Whitney": "...", "Raytheon": "..." }
      }
    """
    if not mda_text.strip():
        print("[INFO] No MD&A text to summarize guidance from.")
        return {
            "next_year_overall_revenue": None,
            "next_year_fcff": None,
            "segment_guidance": {
                "Collins Aerospace": None,
                "Pratt & Whitney": None,
                "Raytheon": None
            }
        }

    # Limit MD&A text length for prompt (DeepSeek's context window)
    mda_snippet = mda_text[:8000] # Use a reasonable slice

    prompt = (
        f"You are a financial analyst. Below is the MD&A section from the {ticker} 10-K.\n\n"
        f"{mda_snippet}\n\n"
        f"Extract any specific quantitative forward guidance (e.g., specific values, percentage ranges) for:\n"
        f"1) Next fiscal year overall revenue. If a range is provided, capture it as a string (e.g., '7.5 to 8.2 billion').\n"
        f"2) Next fiscal year Free Cash Flow to Firm (FCFF) or Free Cash Flow (FCF). Capture as a string.\n"
        f"3) Any segment-level guidance (e.g., revenue growth %, margin expectations) for segments like 'Collins Aerospace', 'Pratt & Whitney', 'Raytheon'. Capture as strings.\n"
        f"Return your answer strictly as a valid JSON object with the following keys:\n"
        f"  \"next_year_overall_revenue\": string or null,\n"
        f"  \"next_year_fcff\": string or null,\n"
        f"  \"segment_guidance\": {{\"Collins Aerospace\": string or null, \"Pratt & Whitney\": string or null, \"Raytheon\": string or null, \"Other Segments\": string or null}}\n"
        f"If a particular data point is not explicitly mentioned or is purely qualitative (e.g., 'expected to grow'), set its value to null.\n"
        f"Example output: {{\n"
        f"  \"next_year_overall_revenue\": \"$70.5 billion to $72.0 billion\",\n"
        f"  \"next_year_fcff\": \"Approximately $4.5 billion\",\n"
        f"  \"segment_guidance\": {{\n"
        f"    \"Collins Aerospace\": \"Revenue growth mid-single digits\",\n"
        f"    \"Pratt & Whitney\": \"Operating profit expected to improve\",\n"
        f"    \"Raytheon\": null,\n"
        f"    \"Other Segments\": null\n"
        f"  }}\n"
        f"}}"
    )
    raw = deepseek_chat(prompt, max_tokens=1024, temperature=0.1) # Lower temperature for JSON output
    guidance = {
        "next_year_overall_revenue": None,
        "next_year_fcff": None,
        "segment_guidance": {
            "Collins Aerospace": None,
            "Pratt & Whitney": None,
            "Raytheon": None,
            "Other Segments": None # Added a generic "Other" category
        }
    }
    try:
        parsed_guidance = json.loads(raw)
        guidance.update(parsed_guidance)
        # Ensure segment_guidance subkeys exist
        if "segment_guidance" in parsed_guidance and isinstance(parsed_guidance["segment_guidance"], dict):
            for seg, val in parsed_guidance["segment_guidance"].items():
                if seg in guidance["segment_guidance"]:
                    guidance["segment_guidance"][seg] = val
                else: # Catch any other segments LLM might find
                    guidance["segment_guidance"][seg] = val

    except json.JSONDecodeError as e:
        print(f"❌ Error: Could not parse management guidance JSON. LLM response was:\n{raw}\nError: {e}")
    except Exception as e:
        print(f"❌ An unexpected error occurred processing guidance: {e}")
    return guidance


# ────────────────────────────────────────────────────────────────────────────────
# 3. Analyst Consensus Scraping from Yahoo Finance
# ────────────────────────────────────────────────────────────────────────────────
def scrape_yahoo_finance_consensus(ticker: str) -> dict:
    """
    ATTENTION: This web scraping function is HIGHLY FRAGILE and subject to breaking
    if Yahoo Finance changes its HTML structure. For production use, consider
    using a dedicated financial data API (e.g., Alpha Vantage, Financial Modeling Prep).

    Basic scrape of Yahoo Finance 'Analysis' page for forward revenue/FCF estimates.
    """
    print(f"[INFO] Attempting to scrape Yahoo Finance consensus for {ticker}...")
    consensus = {"revenue": {}, "freeCashFlow": {}}
    url = f"https://finance.yahoo.com/quote/{ticker}/analysis?p={ticker}"
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"}
        resp = requests.get(url, headers=headers, timeout=10)
        resp.raise_for_status() # Raise an exception for HTTP errors
        soup = BeautifulSoup(resp.text, "html.parser")

        # Yahoo Finance tables typically have specific attributes or are within certain divs
        # This selector targets tables that are part of the 'Estimates' section.
        # This is still a heuristic and might break.
        estimate_tables = soup.find_all("table", class_="W(100%) M(0)") # This class name is prone to change!

        for table in estimate_tables:
            # Check if this table contains "Revenue Estimate" or "Free Cash Flow Estimate"
            table_text = table.get_text()
            if "Revenue Estimate" not in table_text and "Free Cash Flow Estimate" not in table_text and "FCF Estimate" not in table_text:
                continue # Skip tables that don't seem to contain the relevant estimates

            rows = table.find_all("tr")
            # The first row often contains the years, e.g., 'Current Year', 'Next Year', 'Next 5 Years (per annum)'
            # Let's try to dynamically identify the year columns
            header_row = rows[0] if rows else None
            if header_row:
                headers = [th.get_text(strip=True) for th in header_row.find_all("th")]
                # Attempt to map headers to years (e.g., "2024", "2025")
                year_cols = {}
                current_year_fallback = 2024 # A reasonable current year fallback for now
                for i, h in enumerate(headers):
                    if re.match(r"\d{4}", h): # Matches "2024" directly
                        year_cols[int(h)] = i
                    elif "Current Year" in h:
                        year_cols[current_year_fallback] = i
                    elif "Next Year" in h:
                        year_cols[current_year_fallback + 1] = i
                    # Add more sophisticated mapping if needed for "Next 5 Years", etc.

            for row in rows:
                cols = [col.get_text(strip=True) for col in row.find_all("td")]
                if not cols:
                    continue
                label = cols[0]

                # Assuming values are in millions unless specified
                scale = 1_000_000 # Default to millions

                # Yahoo finance might use abbreviations like "B" for billions, "M" for millions in their numbers
                # Let's improve the parsing to handle that
                def parse_yfinance_value(val_str):
                    val_str = val_str.replace(",", "")
                    if val_str.endswith('B'):
                        return float(val_str[:-1]) * 1_000_000_000
                    elif val_str.endswith('M'):
                        return float(val_str[:-1]) * 1_000_000
                    elif val_str.endswith('T'): # Trillions? Unlikely for single company estimates
                        return float(val_str[:-1]) * 1_000_000_000_000
                    try:
                        # Assume if no suffix, it's already in the base unit (e.g., dollars or millions from table context)
                        # This part is tricky as context for millions/billions is often in table headers
                        # For simplicity, if no suffix, assume it's raw value, then apply default scale if needed
                        return float(val_str)
                    except ValueError:
                        return None

                if "Revenue Estimate" in label:
                    # Fill revenue estimates based on identified year columns
                    for year, idx in year_cols.items():
                        if idx < len(cols):
                            val = parse_yfinance_value(cols[idx])
                            if val is not None:
                                consensus["revenue"][str(year)] = val
                                print(f"[INFO] Yahoo Finance Revenue Estimate for {year}: ${val:,.0f}")

                if "Free Cash Flow Estimate" in label or "FCF Estimate" in label:
                    # Fill FCF estimates based on identified year columns
                    for year, idx in year_cols.items():
                        if idx < len(cols):
                            val = parse_yfinance_value(cols[idx])
                            if val is not None:
                                consensus["freeCashFlow"][str(year)] = val
                                print(f"[INFO] Yahoo Finance FCF Estimate for {year}: ${val:,.0f}")
        if not consensus["revenue"] and not consensus["freeCashFlow"]:
            print(f"⚠️ Warning: No revenue or FCF consensus found on Yahoo Finance for {ticker}. HTML structure might have changed.")

    except requests.exceptions.HTTPError as e:
        print(f"❌ Error during Yahoo Finance scrape (HTTP {e.response.status_code}): {e}")
    except requests.exceptions.ConnectionError as e:
        print(f"❌ Error during Yahoo Finance scrape (Connection): {e}")
    except requests.exceptions.Timeout as e:
        print(f"❌ Error during Yahoo Finance scrape (Timeout): {e}")
    except Exception as e:
        print(f"❌ An unexpected error occurred during Yahoo Finance scrape: {e}")
    return consensus


# ────────────────────────────────────────────────────────────────────────────────
# 4. Build FCFF Forecast Using Guidance, Consensus, & Segment Growth
# ────────────────────────────────────────────────────────────────────────────────
def build_fcff_forecast(
    hist_year: int,
    last_historical_fcf: float,
    mgmt_guidance: dict,
    yahoo_consensus: dict,
    segment_summary_text: str
) -> dict[int, float]:
    """
    Returns a dict {year: fcf} for the next 5 fiscal years.
    Priority for Year 1 & 2:
      1) mgmt_guidance["next_year_fcff"] (midpoint if a range)
      2) yahoo_consensus["freeCashFlow"][year]
      3) fallback growth of last historical FCF at 5% (or based on initial LLM estimate)

    Years 3-5: grown by an LLM-proposed blended segment rate, with a default.
    """

    def parse_range_to_midpoint(text: str) -> float | None:
        """
        If text = "7.5 to 8.2 billion" or "$7.5B - $8.2B", parse and return midpoint.
        Handles M/B/T suffixes.
        """
        if not text or not isinstance(text, str):
            return None
        text = text.replace(",", "").replace("$", "").lower()
        nums = re.findall(r"(\d+\.?\d*)\s*(million|billion|trillion|m|b|t)?", text)

        values = []
        for num_str, scale_str in nums:
            try:
                val = float(num_str)
                if scale_str in ("million", "m"):
                    val *= 1_000_000
                elif scale_str in ("billion", "b"):
                    val *= 1_000_000_000
                elif scale_str in ("trillion", "t"):
                    val *= 1_000_000_000_000
                values.append(val)
            except ValueError:
                continue

        if not values:
            return None
        elif len(values) >= 2:
            return (values[0] + values[1]) / 2.0
        else: # Single value
            return values[0]

    forecast = {}
    # Year 1 = hist_year + 1
    y1 = hist_year + 1
    fcf_y1 = None

    print(f"[FORECAST DEBUG] Building FCFF forecast from historical FCF: ${last_historical_fcf:,.0f} (Year {hist_year})")

    # a) Management guidance
    fcf_guidance = parse_range_to_midpoint(mgmt_guidance.get("next_year_fcff"))
    if fcf_guidance is not None:
        fcf_y1 = fcf_guidance
        print(f"[INFO] Using management guidance FCFF for {y1}: ${fcf_y1:,.0f}")

    # b) Yahoo consensus
    if fcf_y1 is None:
        # Check if yahoo_consensus has data for y1
        yc_key = str(y1)
        yc_data = yahoo_consensus.get("freeCashFlow", {})
        if yc_key in yc_data:
            fcf_y1 = float(yc_data[yc_key])
            print(f"[INFO] Using Yahoo Finance consensus FCFF for {y1}: ${fcf_y1:,.0f}")
        else:
            print(f"[INFO] No Yahoo Finance consensus for {y1} found.")

    # c) Fallback to initial growth rate based on LLM's assessment of segment trends, or 5%
    initial_growth_rate = 0.05
    if segment_summary_text and segment_summary_text.strip() != "No segment tables found.":
        prompt_initial_growth = (
            f"You are a financial modeler. Based on this segment performance summary:\n\n"
            f"{segment_summary_text}\n\n"
            f"Estimate a reasonable initial (Year 1 to Year 2) annual growth rate for total Free Cash Flow (as a decimal, e.g., 0.065, or 0.05 if no clear indication). "
            f"Consider recent trends and the overall business outlook from the summary. "
            f"Only return the decimal number."
        )
        try:
            raw_initial_rate = deepseek_chat(prompt_initial_growth, max_tokens=10, temperature=0.0) # Very low temp for number
            m = re.search(r"0\.\d+", raw_initial_rate)
            if m:
                initial_growth_rate = float(m.group(0))
                print(f"[INFO] LLM-estimated initial growth rate: {initial_growth_rate*100:.1f}%")
            else:
                print(f"[WARNING] LLM could not provide initial growth rate from segment summary ('{raw_initial_rate.strip()}'); defaulting to 5%.")
        except Exception as e:
            print(f"[WARNING] LLM request for initial growth rate failed: {e}; defaulting to 5%.")
    else:
        print("[WARNING] No segment summary available; using default initial growth rate of 5%.")

    if fcf_y1 is None:
        fcf_y1 = float(last_historical_fcf) * (1.0 + initial_growth_rate)
        print(f"[WARNING] No explicit FCFF for {y1}, fallback to {last_historical_fcf:,.0f} * (1 + {initial_growth_rate:.2f}) = ${fcf_y1:,.0f}")
    forecast[y1] = fcf_y1

    # Year 2 = hist_year + 2
    y2 = hist_year + 2
    fcf_y2 = None

    # a) If management provided a second-year guide (if the prompt for LLM was modified to ask for it)
    # The current prompt doesn't explicitly ask for next_year_fcff_plus1, so this will likely be None
    fcf_guidance2 = parse_range_to_midpoint(mgmt_guidance.get("next_year_fcff_plus1"))
    if fcf_guidance2 is not None:
        fcf_y2 = fcf_guidance2
        print(f"[INFO] Using management guidance FCFF for {y2}: ${fcf_y2:,.0f}")

    # b) Yahoo consensus
    if fcf_y2 is None:
        yc_key2 = str(y2)
        if yc_key2 in yc_data: # Reuse yc_data from above
            fcf_y2 = float(yc_data[yc_key2])
            print(f"[INFO] Using Yahoo Finance consensus FCFF for {y2}: ${fcf_y2:,.0f}")

    # c) Fallback to growing y1 by initial growth rate
    if fcf_y2 is None:
        fcf_y2 = float(forecast[y1]) * (1.0 + initial_growth_rate) # Use the initial growth rate for y2 as well
        print(f"[WARNING] No explicit FCFF for {y2}, fallback to {forecast[y1]:,.0f} * (1 + {initial_growth_rate:.2f}) = ${fcf_y2:,.0f}")
    forecast[y2] = fcf_y2

    # Years 3-5: determine a blended segment growth rate via LLM for the long-term
    blended_rate = 0.03 # Default for long-term if LLM fails or no segments
    if segment_summary_text and segment_summary_text.strip() != "No segment tables found.":
        prompt_long_term_growth = (
            f"You are a financial modeler. Based on this segment performance summary and considering long-term outlook (Years 3-5):\n\n"
            f"{segment_summary_text}\n\n"
            f"Considering that Year 1 and 2 growth might be higher due to specific factors, estimate a realistic, sustainable *long-term* (Year 3 to Year 5) annual growth rate for total Free Cash Flow (as a decimal, e.g., 0.03 for 3%). "
            f"This rate should reflect long-term industry trends and the company's competitive advantages. "
            f"If you cannot confidently estimate a specific growth rate from the text, return a conservative default of 0.03 (3%). "
            f"Only return the decimal number."
        )
        try:
            raw_long_term_rate = deepseek_chat(prompt_long_term_growth, max_tokens=10, temperature=0.0)
            m = re.search(r"0\.\d+", raw_long_term_rate)
            if m:
                blended_rate = float(m.group(0))
                print(f"[INFO] LLM-estimated long-term FCFF growth rate: {blended_rate*100:.1f}%")
                if not (0.01 <= blended_rate <= 0.05): # Sanity check for reasonable long-term growth (1-5%)
                     print(f"⚠️ Warning: LLM-estimated long-term growth rate {blended_rate*100:.1f}% is outside typical 1-5% range; defaulting to 3%.")
                     blended_rate = 0.03
            else:
                print(f"[WARNING] LLM could not provide long-term growth rate from segment summary ('{raw_long_term_rate.strip()}'); defaulting to 3%.")
        except Exception as e:
            print(f"[WARNING] LLM long-term rate request failed: {e}; defaulting to 3%.")
    else:
        print("[WARNING] No segment summary available; using default long-term growth rate of 3% for years 3-5.")

    # Project years 3-5
    for i in range(3, 6): # Project for years hist_year+3, hist_year+4, hist_year+5
        y = hist_year + i
        prev = forecast[y - 1]
        forecast[y] = float(prev) * (1.0 + blended_rate)
        print(f"[DEBUG] Projected FCFF for {y} = {prev:,.0f} * (1+{blended_rate:.3f}) = ${forecast[y]:,.0f}")

    return forecast


# ────────────────────────────────────────────────────────────────────────────────
# 5. Full DCF Agent That Combines All Steps
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
        forecasted_fcff (dict): Dictionary of {year: FCF_value} for explicit forecast period.
        terminal_growth_rate (float): Perpetual growth rate for terminal value (e.g., 0.02).
        wacc (float): Weighted Average Cost of Capital (discount rate, e.g., 0.08).
        net_debt (float): Company's net debt (Total Debt - Cash & Equivalents).
        shares_outstanding (int): Number of shares outstanding.
        current_share_price (float | None): Current market price for context/comparison.

    Returns:
        dict: A dictionary containing intrinsic value per share and detailed breakdown.
    """
    if not forecasted_fcff or not shares_outstanding or wacc <= 0 or terminal_growth_rate >= wacc:
        print("❌ Error: Invalid inputs for DCF calculation.")
        return {"error": "Invalid DCF inputs"}

    # Sort forecasted FCFF by year
    sorted_years = sorted(forecasted_fcff.keys())
    explicit_forecast_period_length = len(sorted_years)

    present_value_of_fcff = 0
    print("\n--- DCF Calculation Details ---")
    print(f"WACC: {wacc:.2%}, Terminal Growth Rate: {terminal_growth_rate:.2%}")

    for i, year in enumerate(sorted_years):
        fcf = forecasted_fcff[year]
        # Discount factor for year 'i+1' (since first year is year 1)
        discount_factor = (1 + wacc)**(i + 1)
        pv_fcf = fcf / discount_factor
        present_value_of_fcff += pv_fcf
        print(f"Year {year}: FCF = ${fcf:,.0f}, Discount Factor = {discount_factor:.2f}, PV(FCF) = ${pv_fcf:,.0f}")

    # Calculate Terminal Value (TV) using the last forecasted FCF
    last_forecast_year = sorted_years[-1]
    last_fcf = forecasted_fcff[last_forecast_year]

    if wacc - terminal_growth_rate <= 0:
        print(f"❌ Error: WACC ({wacc:.2%}) must be greater than Terminal Growth Rate ({terminal_growth_rate:.2%}) for Terminal Value calculation.")
        terminal_value = 0
    else:
        terminal_value = last_fcf * (1 + terminal_growth_rate) / (wacc - terminal_growth_rate)
        # Discount Terminal Value back to present
        tv_discount_factor = (1 + wacc)**explicit_forecast_period_length
        present_value_of_terminal_value = terminal_value / tv_discount_factor
        print(f"Last Forecast FCF ({last_forecast_year}): ${last_fcf:,.0f}")
        print(f"Terminal Value: ${terminal_value:,.0f} (Discounted to PV: ${present_value_of_terminal_value:,.0f})")


    # Enterprise Value = PV of Explicit FCFF + PV of Terminal Value
    enterprise_value = present_value_of_fcff + present_value_of_terminal_value
    print(f"Total Present Value of Forecasted FCFF: ${present_value_of_fcff:,.0f}")
    print(f"Total Present Value of Terminal Value: ${present_value_of_terminal_value:,.0f}")
    print(f"Calculated Enterprise Value: ${enterprise_value:,.0f}")

    # Equity Value = Enterprise Value - Net Debt (or + Net Cash)
    equity_value = enterprise_value - net_debt
    print(f"Net Debt: ${net_debt:,.0f}")
    print(f"Calculated Equity Value: ${equity_value:,.0f}")

    # Intrinsic Value Per Share = Equity Value / Shares Outstanding
    intrinsic_value_per_share = equity_value / shares_outstanding
    print(f"Shares Outstanding: {shares_outstanding:,.0f}")
    print(f"Intrinsic Value Per Share: ${intrinsic_value_per_share:,.2f}")

    result = {
        "forecasted_fcff": forecasted_fcff,
        "present_value_of_explicit_fcff": present_value_of_fcff,
        "terminal_value": terminal_value,
        "present_value_of_terminal_value": present_value_of_terminal_value,
        "enterprise_value": enterprise_value,
        "net_debt": net_debt,
        "equity_value": equity_value,
        "shares_outstanding": shares_outstanding,
        "intrinsic_value_per_share": intrinsic_value_per_share
    }

    if current_share_price is not None:
        print(f"Current Share Price: ${current_share_price:,.2f}")
        deviation = ((intrinsic_value_per_share - current_share_price) / current_share_price) * 100
        print(f"Deviation from current price: {deviation:+.2f}%")
        result["current_share_price"] = current_share_price
        result["deviation_from_current_price_pct"] = deviation

    return result

def run_dcf_model(
    company_name: str,
    assumptions: dict, # Expected to contain 'base_wacc', 'base_terminal_growth_rate', 'bull_multipliers', 'bear_multipliers'
    file_paths: list[str],
    current_share_price: float | None = None # Optional: provide current share price for context
) -> tuple[str, dict]:
    """
    Performs a 3-case (Bull / Base / Bear) FCFF-based DCF valuation for the given company_name.
    1) Uses DeepSeek to convert company_name -> ticker.
    2) Extracts historical financials from uploaded files.
    3) Extracts segment performance and management guidance from the most recent file.
    4) Scrapes analyst consensus from Yahoo Finance (with strong warning about fragility).
    5) Builds a 5-year FCFF forecast using extracted data and LLM assistance.
    6) Calculates intrinsic value per share for Base, Bull, and Bear cases.

    Args:
        company_name (str): The full name of the company (e.g., "RTX Corporation").
        assumptions (dict): Dictionary with DCF assumptions:
            - 'base_wacc': Base case WACC (e.g., 0.08)
            - 'base_terminal_growth_rate': Base case terminal growth (e.g., 0.02)
            - 'bull_growth_multiplier': Multiplier for base growth rate in bull case (e.g., 1.10 for +10% growth)
            - 'bull_wacc_adjust': Adjustment for WACC in bull case (e.g., -0.005 for -0.5% WACC)
            - 'bull_terminal_growth_adjust': Adjustment for terminal growth in bull case (e.g., 0.005 for +0.5% terminal growth)
            - 'bear_growth_multiplier': Multiplier for base growth rate in bear case (e.g., 0.90 for -10% growth)
            - 'bear_wacc_adjust': Adjustment for WACC in bear case (e.g., +0.005 for +0.5% WACC)
            - 'bear_terminal_growth_adjust': Adjustment for terminal growth in bear case (e.g., -0.005 for -0.5% terminal growth)
        file_paths (list[str]): List of paths to uploaded annual reports (PDF/DOCX).
        current_share_price (float | None): Optional, current market price for comparison.

    Returns:
        tuple[str, dict]: A message string (success/error) and a dictionary of DCF results.
    """
    results = {}

    print(f"🚀 Starting DCF analysis for {company_name}...")

    # 1. Convert company name to ticker
    ticker_prompt = f"What is the stock ticker symbol for {company_name}? Reply with only the ticker symbol."
    ticker = deepseek_chat(ticker_prompt, max_tokens=8, temperature=0.0).strip().upper()
    if not ticker:
        return "❌ Error: Could not determine ticker symbol for the company.", {}
    print(f"✅ Determined ticker: {ticker}")

    # Fetch current share price if not provided
    if current_share_price is None:
        current_share_price = get_current_share_price(ticker)
        if current_share_price:
            print(f"✅ Fetched current market price for {ticker}: ${current_share_price:,.2f}")
        else:
            print(f"⚠️ Warning: Could not fetch current market price for {ticker} from yfinance. Will proceed without it for comparison.")
    else:
        print(f"✅ Using provided current market price: ${current_share_price:,.2f}")


    # 2. Extract historical financials
    historical_financials = []
    most_recent_file = None
    most_recent_year = -1
    for f_path in file_paths:
        data = extract_financials_from_file(f_path, ticker=ticker)
        if data and data.get("year") is not None:
            historical_financials.append(data)
            if data["year"] > most_recent_year:
                most_recent_year = data["year"]
                most_recent_file = f_path
        else:
            print(f"⚠️ Warning: Could not extract complete financial data from {f_path}.")

    if not historical_financials:
        return "❌ Error: DCF processing failed: No valid historical financial data extracted from any provided files. Please ensure files are readable and contain standard financial statements.", {}

    # Sort financials by year to ensure we pick the latest
    historical_financials.sort(key=lambda x: x.get("year") or 0, reverse=True)
    latest_financials = historical_financials[0]

    last_historical_fcf = latest_financials.get("free_cash_flow")
    hist_year = latest_financials.get("year")
    net_debt_from_files = (latest_financials.get("total_debt") or 0) - (latest_financials.get("cash_ce") or 0)

    if last_historical_fcf is None or hist_year is None:
        # Fallback for FCF if still missing from documents
        yf_fcf_fallback = get_yfinance_fcf(ticker)
        if yf_fcf_fallback is not None:
            last_historical_fcf = yf_fcf_fallback
            print(f"✅ Using yfinance fallback for latest historical FCF: ${last_historical_fcf:,.0f}")
        if hist_year is None: # Attempt to get latest fiscal year if not found
             try:
                 hist_year = yf.Ticker(ticker).info.get("mostRecentFiscalYearEnd")
                 if hist_year:
                     hist_year = pd.to_datetime(hist_year, unit='s').year
                     print(f"✅ Using yfinance fallback for most recent fiscal year: {hist_year}")
             except Exception:
                 hist_year = 2023 # Hardcoded fallback if all else fails

    if last_historical_fcf is None or hist_year is None:
        return "❌ Error: Could not determine latest historical Free Cash Flow or its corresponding year, which is essential for forecasting.", {}

    # Get shares outstanding and net debt for valuation
    shares_outstanding = get_shares_outstanding(ticker)
    if not shares_outstanding:
        return "❌ Error: Could not retrieve shares outstanding from yfinance.", {}
    print(f"✅ Shares Outstanding: {shares_outstanding:,.0f}")

    # Use net debt from files, but fallback to yfinance if not available or zero from files
    net_debt = net_debt_from_files
    if net_debt == 0 or net_debt is None: # If not found or weird value from docs, try yfinance
        yf_net_debt = get_yfinance_net_debt(ticker)
        if yf_net_debt is not None:
            net_debt = yf_net_debt
            print(f"✅ Using yfinance fallback for Net Debt: ${net_debt:,.0f}")
        else:
            net_debt = 0 # Default to zero if no debt info found
            print(f"⚠️ Warning: Could not retrieve Net Debt from files or yfinance. Assuming Net Debt = $0 for DCF.")
    else:
        print(f"✅ Net Debt from files: ${net_debt:,.0f}")


    # 3. Extract segment performance and management guidance (from most recent file)
    segment_summary_text = ""
    mda_text = ""
    mgmt_guidance = {}

    if most_recent_file and os.path.splitext(most_recent_file.lower())[1] == ".pdf":
        print(f"\n📊 Analyzing segment performance from {most_recent_file}...")
        segment_pages = find_segment_pages(most_recent_file)
        segment_dfs = extract_segment_tables(most_recent_file, segment_pages)
        if segment_dfs:
            segment_summary_text = summarize_segment_performance(segment_dfs, ticker)
            print("✅ Segment performance summary generated.")
        else:
            print("⚠️ Warning: No segment tables extracted. Segment-based growth projection may be less accurate.")

        print(f"\n🗣️ Extracting management guidance from {most_recent_file}...")
        mda_pages = find_mda_pages(most_recent_file)
        mda_text = extract_mda_text(most_recent_file, mda_pages)
        if mda_text:
            mgmt_guidance = summarize_management_guidance(mda_text, ticker)
            print("✅ Management guidance extracted.")
            print(f"   Guidance: {json.dumps(mgmt_guidance, indent=2)}")
        else:
            print("⚠️ Warning: No MD&A text extracted. Management guidance unavailable.")
    else:
        print("⚠️ Warning: Most recent file is not a PDF, skipping segment and guidance extraction from documents.")

    # 4. Scrape analyst consensus
    yahoo_consensus = scrape_yahoo_finance_consensus(ticker)
    if yahoo_consensus["revenue"] or yahoo_consensus["freeCashFlow"]:
        print("✅ Analyst consensus scraped from Yahoo Finance.")
        print(f"   Yahoo Consensus: {json.dumps(yahoo_consensus, indent=2)}")
    else:
        print("⚠️ Warning: No analyst consensus data scraped from Yahoo Finance. This will impact forecast accuracy.")


    # 5. Build FCFF forecast for 5 years
    print("\n📈 Building 5-year FCFF forecast...")
    base_forecast_fcff = build_fcff_forecast(
        hist_year=hist_year,
        last_historical_fcf=last_historical_fcf,
        mgmt_guidance=mgmt_guidance,
        yahoo_consensus=yahoo_consensus,
        segment_summary_text=segment_summary_text
    )
    if not base_forecast_fcff:
        return "❌ Error: Could not build FCFF forecast. Check historical data and parsing.", {}
    print(f"✅ Base FCFF Forecast: {base_forecast_fcff}")

    # 6. Calculate DCF for Base, Bull, Bear cases
    print("\n📊 Calculating DCF valuation for Base, Bull, and Bear cases...")

    # BASE CASE
    base_wacc = assumptions.get('base_wacc', 0.08)
    base_terminal_growth_rate = assumptions.get('base_terminal_growth_rate', 0.02)
    print("\n--- Base Case ---")
    base_dcf_result = calculate_dcf(
        forecasted_fcff=base_forecast_fcff,
        terminal_growth_rate=base_terminal_growth_rate,
        wacc=base_wacc,
        net_debt=net_debt,
        shares_outstanding=shares_outstanding,
        current_share_price=current_share_price
    )
    results["base_case"] = base_dcf_result
    if base_dcf_result.get("error"):
        print(f"❌ Base Case DCF failed: {base_dcf_result['error']}")


    # BULL CASE
    bull_forecast_fcff = {}
    bull_growth_multiplier = assumptions.get('bull_growth_multiplier', 1.10) # 10% higher growth
    for year, fcf_val in base_forecast_fcff.items():
        bull_forecast_fcff[year] = fcf_val * bull_growth_multiplier

    bull_wacc = base_wacc + assumptions.get('bull_wacc_adjust', -0.005) # Lower WACC by 0.5%
    bull_terminal_growth_rate = base_terminal_growth_rate + assumptions.get('bull_terminal_growth_adjust', 0.005) # Higher TG by 0.5%
    print("\n--- Bull Case ---")
    bull_dcf_result = calculate_dcf(
        forecasted_fcff=bull_forecast_fcff,
        terminal_growth_rate=bull_terminal_growth_rate,
        wacc=bull_wacc,
        net_debt=net_debt,
        shares_outstanding=shares_outstanding,
        current_share_price=current_share_price
    )
    results["bull_case"] = bull_dcf_result
    if bull_dcf_result.get("error"):
        print(f"❌ Bull Case DCF failed: {bull_dcf_result['error']}")


    # BEAR CASE
    bear_forecast_fcff = {}
    bear_growth_multiplier = assumptions.get('bear_growth_multiplier', 0.90) # 10% lower growth
    for year, fcf_val in base_forecast_fcff.items():
        bear_forecast_fcff[year] = fcf_val * bear_growth_multiplier

    bear_wacc = base_wacc + assumptions.get('bear_wacc_adjust', 0.005) # Higher WACC by 0.5%
    bear_terminal_growth_rate = base_terminal_growth_rate + assumptions.get('bear_terminal_growth_adjust', -0.005) # Lower TG by 0.5%
    print("\n--- Bear Case ---")
    bear_dcf_result = calculate_dcf(
        forecasted_fcff=bear_forecast_fcff,
        terminal_growth_rate=bear_terminal_growth_rate,
        wacc=bear_wacc,
        net_debt=net_debt,
        shares_outstanding=shares_outstanding,
        current_share_price=current_share_price
    )
    results["bear_case"] = bear_dcf_result
    if bear_dcf_result.get("error"):
        print(f"❌ Bear Case DCF failed: {bear_dcf_result['error']}")


    final_message = f"✅ DCF analysis complete for {company_name} ({ticker})."
    if current_share_price:
        final_message += f"\n   Current Share Price: ${current_share_price:,.2f}"
        if "base_case" in results and "intrinsic_value_per_share" in results["base_case"]:
             final_message += f"\n   Base Case Intrinsic Value: ${results['base_case']['intrinsic_value_per_share']:,.2f}"
             final_message += f" (Deviation: {results['base_case'].get('deviation_from_current_price_pct', 0):+.2f}%)"
        if "bull_case" in results and "intrinsic_value_per_share" in results["bull_case"]:
            final_message += f"\n   Bull Case Intrinsic Value: ${results['bull_case']['intrinsic_value_per_share']:,.2f}"
        if "bear_case" in results and "intrinsic_value_per_share" in results["bear_case"]:
            final_message += f"\n   Bear Case Intrinsic Value: ${results['bear_case']['intrinsic_value_per_share']:,.2f}"

    return final_message, results