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
from openpyxl import Workbook
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
    except Exception:
        pass
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
            return None
        # yfinance columns are dates; first column is most recent
        most_recent_col = cf.columns[0]
        for label in ("Free Cash Flow", "FreeCashFlow"):
            if label in cf.index:
                fcf_val = cf.loc[label, most_recent_col]
                return float(fcf_val)
    except Exception:
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
    except Exception:
        pass
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
    except Exception:
        pass
    return "\n".join(all_text)


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Identify page numbers containing a given keyword in PDF
# ────────────────────────────────────────────────────────────────────────────────
def find_pages_with_keyword(pdf_path: str, keyword: str) -> list[int]:
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
    dfs = []
    if not page_indices:
        return dfs
    pages_str = ",".join(str(i + 1) for i in page_indices)
    try:
        tables = camelot.read_pdf(pdf_path, pages=pages_str, flavor="stream")
        for t in tables:
            dfs.append(t.df)
    except Exception:
        pass
    return dfs


# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Parse Free Cash Flow from a DataFrame with better scale detection
# ────────────────────────────────────────────────────────────────────────────────
def parse_fcf_from_df(df: pd.DataFrame) -> float | None:
    """
    Searches df for a row containing 'Free Cash Flow' (case-insensitive)
    and returns the first numeric value in that row (USD).
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

    for _, row in df.iterrows():
        row_str = " ".join(str(cell) for cell in row.tolist())
        if re.search(r"free\s+cash\s+flow", row_str, re.IGNORECASE):
            for cell in row.tolist()[1:]:
                s = str(cell).replace(",", "").replace("$", "").replace("(", "-").replace(")", "")
                try:
                    val = float(s)
                    return val * scale
                except Exception:
                    continue
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
        return None

    if not full_text.strip():
        print(f"[EXTRACT DEBUG] No text found in {file_path}. Skipping.")
        return None

    # 1) Find year via "Year Ended ... <Year>"
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

    # 2) If PDF, attempt table extraction
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
                fcf = float(val)
                break

        # b) Income Statement → parse revenue, net income
        inc_pages = find_pages_with_keyword(
            file_path,
            r"Consolidated\s+Statements\s+of\s+Income|Consolidated\s+Statements\s+of\s+Comprehensive\s+Income|Income\s+Statement"
        )
        tables_inc = extract_tables_from_pdf(file_path, inc_pages)
        for df in tables_inc:
            r, ni = parse_income_from_df(df)
            if r is not None:
                revenue = float(r)
            if ni is not None:
                net_income = float(ni)
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
            if c is not None:
                cash_ce = float(c)
            if total_debt is not None and cash_ce is not None:
                break

    # 3) Fallback line-by-line search for any missing items
    if fcf is None or revenue is None or net_income is None or total_debt is None or cash_ce is None:
        for line in full_text.splitlines():
            # Free Cash Flow
            if fcf is None and re.search(r"free\s+cash\s+flow", line, re.IGNORECASE):
                m = re.search(r"\$[\s,]*([\d,]+(?:\.\d+)?)", line)
                if m:
                    try:
                        fcf_val = float(m.group(1).replace(",", ""))
                        fcf = float(fcf_val)
                    except:
                        pass

            # Revenue
            if revenue is None and re.search(r"(^|\s)revenue($|\s)", line, re.IGNORECASE):
                m = re.search(r"\$[\s,]*([\d,]+(?:\.\d+)?)", line)
                if m:
                    try:
                        rev_val = float(m.group(1).replace(",", ""))
                        revenue = float(rev_val)
                    except:
                        pass

            # Net Income
            if net_income is None and re.search(r"(^|\s)net\s+income($|\s)", line, re.IGNORECASE):
                m = re.search(r"\$[\s,]*([\d,]+(?:\.\d+)?)", line)
                if m:
                    try:
                        ni_val = float(m.group(1).replace(",", ""))
                        net_income = float(ni_val)
                    except:
                        pass

            # Total Debt
            if total_debt is None and re.search(r"(^|\s)total\s+debt($|\s)", line, re.IGNORECASE):
                m = re.search(r"\$[\s,]*([\d,]+(?:\.\d+)?)", line)
                if m:
                    try:
                        td_val = float(m.group(1).replace(",", ""))
                        total_debt = float(td_val)
                    except:
                        pass

            # Cash & Equivalents
            if cash_ce is None and re.search(r"cash\s+and\s+cash\s+equivalents", line, re.IGNORECASE):
                m = re.search(r"\$[\s,]*([\d,]+(?:\.\d+)?)", line)
                if m:
                    try:
                        c_val = float(m.group(1).replace(",", ""))
                        cash_ce = float(c_val)
                    except:
                        pass

            if fcf is not None and revenue is not None and net_income is not None and total_debt is not None and cash_ce is not None:
                break

    # If FCF still missing, try yfinance as fallback
    if fcf is None:
        yf_fcf = get_yfinance_fcf(os.path.splitext(os.path.basename(file_path))[0])
        if yf_fcf is not None:
            fcf = float(yf_fcf)
            print(f"[INFO] Using yfinance FCF fallback for {file_path}: ${fcf:,.0f}")

    # 4) If nothing found, give up
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
# 1. Segment Extraction & Summarization
# ────────────────────────────────────────────────────────────────────────────────
def find_segment_pages(pdf_path: str) -> list[int]:
    """
    Returns a list of 0-based page indices where 'Segment' or 'MD&A' appears.
    """
    keywords = [
        r"Management’s\s+Discussion\s+and\s+Analysis",
        r"MD&A",
        r"Segments",
        r"Business\s+Segments",
        r"Segment\s+Results"
    ]
    pages = set()
    try:
        reader = PdfReader(pdf_path)
        for idx, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            for kw in keywords:
                if re.search(kw, text, re.IGNORECASE):
                    pages.add(idx)
    except:
        pass
    return sorted(pages)


def extract_segment_tables(pdf_path: str, segment_pages: list[int]) -> list[pd.DataFrame]:
    """
    Uses Camelot to extract every table on the identified Segment/MD&A pages.
    """
    dfs = []
    if not segment_pages:
        return dfs
    pages_str = ",".join(str(i + 1) for i in segment_pages)
    try:
        tables = camelot.read_pdf(pdf_path, pages=pages_str, flavor="stream")
        for tbl in tables:
            dfs.append(tbl.df)
    except:
        pass
    return dfs


def summarize_segment_performance(dfs: list[pd.DataFrame], ticker: str) -> str:
    """
    Given a list of DataFrames (from camelot) that contain segment revenue & margin tables,
    ask DeepSeek to summarize revenue CAGRs and average margins for each segment.
    """
    if not dfs:
        return "No segment tables found."

    table_strs = []
    for df in dfs:
        # Limit to first 5 rows & 5 columns for prompt brevity
        small = df.iloc[:5, :5].to_csv(index=False)
        table_strs.append(small)

    prompt = (
        f"You are a financial data summarization assistant. For ticker {ticker},\n"
        f"here are extracted segment tables (in CSV form). Each table has columns 'Segment', '2021 Revenue', '2022 Revenue', '2023 Revenue', '2021 Operating Profit', etc.\n"
        f"Compute for each segment:\n"
        f"  • Revenue each year (in USD), and 3-year CAGR (2021→2023).\n"
        f"  • Average operating margin over 2021-2023.\n"
        f"Return a bullet-point summary.\n"
        "\n"
        f"TABLES:\n"
        + "\n\n---\n\n".join(table_strs)
    )

    raw = deepseek_chat(prompt, max_tokens=512)
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
    except:
        pass
    return sorted(pages)


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
            page = reader.pages[idx]
            txt = page.extract_text() or ""
            collected.append(txt)
    except:
        pass
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
        return {
            "next_year_overall_revenue": None,
            "next_year_fcff": None,
            "segment_guidance": {
                "Collins Aerospace": None,
                "Pratt & Whitney": None,
                "Raytheon": None
            }
        }

    prompt = (
        f"You are a financial analyst. Below is the MD&A section from the {ticker} 10-K.\n\n"
        f"{mda_text[:8000]}\n\n"
        f"Extract any forward guidance for:\n"
        f"1) Next fiscal year revenue (overall).\n"
        f"2) Next fiscal year FCFF (free cash flow to firm).\n"
        f"3) Any segment-level guidance (e.g. revenue growth %, margin) for Collins, Pratt & Whitney, Raytheon.\n"
        f"Return your answer strictly as valid JSON with keys:\n"
        f"  \"next_year_overall_revenue\", \"next_year_fcff\", and \"segment_guidance\" (which itself is a dict by segment name).\n"
        f"If a particular data point is not mentioned, set it to null.\n"
    )
    raw = deepseek_chat(prompt, max_tokens=1024)
    try:
        guidance = json.loads(raw)
        # Ensure keys exist
        if "next_year_overall_revenue" not in guidance:
            guidance["next_year_overall_revenue"] = None
        if "next_year_fcff" not in guidance:
            guidance["next_year_fcff"] = None
        if "segment_guidance" not in guidance:
            guidance["segment_guidance"] = {
                "Collins Aerospace": None,
                "Pratt & Whitney": None,
                "Raytheon": None
            }
        else:
            # Ensure subkeys exist
            sg = guidance["segment_guidance"]
            for seg in ("Collins Aerospace", "Pratt & Whitney", "Raytheon"):
                if seg not in sg:
                    sg[seg] = None
    except:
        print("[WARNING] Could not parse management guidance JSON. LLM response:")
        print(raw)
        guidance = {
            "next_year_overall_revenue": None,
            "next_year_fcff": None,
            "segment_guidance": {
                "Collins Aerospace": None,
                "Pratt & Whitney": None,
                "Raytheon": None
            }
        }
    return guidance


# ────────────────────────────────────────────────────────────────────────────────
# 3. Analyst Consensus Scraping from Yahoo Finance
# ────────────────────────────────────────────────────────────────────────────────
def scrape_yahoo_finance_consensus(ticker: str) -> dict:
    """
    Basic scrape of Yahoo Finance 'Analysis' page for forward revenue/FCF estimates.
    This example assumes a table with rows 'Revenue Estimate' and 'Free Cash Flow Estimate'.
    """
    consensus = {"revenue": {}, "freeCashFlow": {}}
    url = f"https://finance.yahoo.com/quote/{ticker}/analysis?p={ticker}"
    try:
        resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"})
        soup = BeautifulSoup(resp.text, "html.parser")
        table = soup.find("table", {"class": "W(100%) M(0)"})
        if table:
            for row in table.find_all("tr"):
                cols = [col.get_text(strip=True) for col in row.find_all("td")]
                if not cols:
                    continue
                label = cols[0]
                # Example parsing: actual HTML may differ—adjust selectors accordingly
                if "Revenue Estimate" in label and len(cols) >= 3:
                    try:
                        # Assume cols[1]=2024, cols[2]=2025 (in millions); convert to absolute
                        val24 = float(cols[1].replace(",", "")) * 1e6
                        val25 = float(cols[2].replace(",", "")) * 1e6
                        consensus["revenue"]["2024"] = val24
                        consensus["revenue"]["2025"] = val25
                    except:
                        pass
                if ("Free Cash Flow Estimate" in label or "FCF Estimate" in label) and len(cols) >= 3:
                    try:
                        val24 = float(cols[1].replace(",", "")) * 1e6
                        val25 = float(cols[2].replace(",", "")) * 1e6
                        consensus["freeCashFlow"]["2024"] = val24
                        consensus["freeCashFlow"]["2025"] = val25
                    except:
                        pass
    except Exception as e:
        print(f"[WARNING] Yahoo Finance scrape failed: {e}")
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
    Priority:
      1) mgmt_guidance["next_year_fcff"] (midpoint if a range)
      2) yahoo_consensus["freeCashFlow"][year]
      3) fallback growth of last historical FCF at 5%
      4) Years 3-5: grown by an LLM-proposed blended segment rate
    """

    def parse_range_to_midpoint(text: str) -> float | None:
        """
        If text = "7.5 to 8.2 billion" or "$7.5B - $8.2B", parse and return midpoint.
        """
        if not text or not isinstance(text, str):
            return None
        try:
            nums = re.findall(r"([\d\.]+)\s*(?:billion|B|bn)?", text.replace(",", "").lower())
            if len(nums) >= 2:
                low = float(nums[0]) * 1e9
                high = float(nums[1]) * 1e9
                return (low + high) / 2.0
            elif len(nums) == 1:
                return float(nums[0]) * 1e9
        except:
            pass
        return None

    forecast = {}
    # Year 1 = hist_year + 1
    y1 = hist_year + 1
    fcf_y1 = None

    # a) Management guidance
    fcf_guidance = parse_range_to_midpoint(mgmt_guidance.get("next_year_fcff"))
    if fcf_guidance is not None:
        fcf_y1 = fcf_guidance
        print(f"[INFO] Using management guidance FCFF for {y1}: ${fcf_y1:,.0f}")

    # b) Yahoo consensus
    if fcf_y1 is None:
        yc = yahoo_consensus.get("freeCashFlow", {}).get("2024")
        if yc is not None:
            fcf_y1 = float(yc)
            print(f"[INFO] Using Yahoo Finance consensus FCFF for {y1}: ${fcf_y1:,.0f}")

    # c) Fallback to 5% growth
    if fcf_y1 is None:
        fcf_y1 = float(last_historical_fcf) * 1.05
        print(f"[WARNING] No explicit FCFF for {y1}, fallback to {last_historical_fcf:,.0f} * 1.05 = ${fcf_y1:,.0f}")
    forecast[y1] = fcf_y1

    # Year 2 = hist_year + 2
    y2 = hist_year + 2
    fcf_y2 = None

    # a) If management provided a second-year guide (check key "next_year_fcff_plus1")
    fcf_guidance2 = parse_range_to_midpoint(mgmt_guidance.get("next_year_fcff_plus1"))
    if fcf_guidance2 is not None:
        fcf_y2 = fcf_guidance2
        print(f"[INFO] Using management guidance FCFF for {y2}: ${fcf_y2:,.0f}")

    # b) Yahoo consensus
    if fcf_y2 is None:
        yc2 = yahoo_consensus.get("freeCashFlow", {}).get("2025")
        if yc2 is not None:
            fcf_y2 = float(yc2)
            print(f"[INFO] Using Yahoo Finance consensus FCFF for {y2}: ${fcf_y2:,.0f}")

    # c) Fallback to growing y1 by 5%
    if fcf_y2 is None:
        fcf_y2 = float(forecast[y1]) * 1.05
        print(f"[WARNING] No explicit FCFF for {y2}, fallback to {forecast[y1]:,.0f} * 1.05 = ${fcf_y2:,.0f}")
    forecast[y2] = fcf_y2

    # Years 3-5: determine a blended segment growth rate via LLM
    blended_rate = 0.05
    if segment_summary_text and segment_summary_text.strip() != "No segment tables found.":
        prompt = (
            f"You are a financial modeler. Based on this segment performance summary:\n\n"
            f"{segment_summary_text}\n\n"
            f"Estimate a single blended forward growth rate (as a decimal) for total FCFF over the next 3 years. "
            f"For example, if Collins is growing 5%, Pratt 8%, Raytheon 7% with revenue weights, return something like 0.066. "
            f"If you cannot estimate, return 0.05."
        )
        try:
            raw_rate = deepseek_chat(prompt, max_tokens=64)
            m = re.search(r"0\.\d+", raw_rate)
            if m:
                blended_rate = float(m.group(0))
                print(f"[INFO] Blended segment FCFF growth rate from LLM: {blended_rate*100:.1f}%")
            else:
                print("[WARNING] LLM response did not contain a valid decimal growth rate; defaulting to 5%")
        except Exception as e:
            print(f"[WARNING] LLM blended rate request failed: {e}; defaulting to 5%")
    else:
        print("[WARNING] No segment summary available; using default growth rate of 5% for years 3-5")

    # Project years 3-5
    for i in range(3, 6):
        y = hist_year + i
        prev = forecast[y - 1]
        forecast[y] = float(prev) * (1.0 + blended_rate)
        print(f"[DEBUG] Projected FCFF for {y} = {prev:,.0f} * (1+{blended_rate:.3f}) = ${forecast[y]:,.0f}")

    return forecast


# ────────────────────────────────────────────────────────────────────────────────
# 5. Full DCF Agent That Combines All Steps
# ────────────────────────────────────────────────────────────────────────────────
# In DCFAgent.py, replace the existing run_dcf_model definition with this corrected version:

def run_dcf_model(
    company_name: str,
    assumptions: dict,
    file_paths: list[str]
) -> tuple[str, dict]:
    """
    Performs a 3-case (Bull / Base / Bear) FCFF-based DCF valuation for the given company_name.
    1) Uses DeepSeek to convert company_name → ticker.
    2) Extracts historical financials from uploaded files.
    3) Extracts segment performance and management guidance from the most recent file.
    4) Scrapes Yahoo Finance consensus.
    5) Builds a bottom-up 5-year FCFF forecast.
    6) Fetches net debt via yfinance or fallback to PDF parse.
    7) Projects and discounts FCFF under Bull/Base/Bear scenarios.
    8) Writes an Excel and returns a summary dict.
    """
    # 1) Convert company_name → ticker via DeepSeek
    ticker_prompt = (
        f"You are a stock ticker lookup assistant.\n"
        f"For the company name \"{company_name}\", provide the primary U.S. stock ticker (no extra text)."
    )
    raw_ticker = deepseek_chat(ticker_prompt, max_tokens=32).strip()
    ticker = raw_ticker.upper()

    # 2) Fetch shares out & current price
    shares_outstanding = get_shares_outstanding(ticker)
    current_price = get_current_share_price(ticker)

    # 3) Extract historical financials
    extracted = []
    for path in file_paths:
        fin = extract_financials_from_file(path)
        if fin and fin.get("year") and fin.get("free_cash_flow") is not None:
            extracted.append(fin)
    if not extracted:
        raise ValueError("No valid historical financial data extracted.")
    extracted = sorted(extracted, key=lambda x: x["year"], reverse=True)[:5]
    historical_list = list(reversed(extracted))  # ascending by year

    hist_years = [e["year"] for e in historical_list]
    hist_fcfs = [e["free_cash_flow"] for e in historical_list]
    hist_revs = [e.get("revenue") for e in historical_list]
    hist_nis = [e.get("net_income") for e in historical_list]
    last_hist_year = hist_years[-1]
    last_hist_fcf = hist_fcfs[-1]

    # 4) Extract segment performance from most recent file
    latest_file = file_paths[-1]
    seg_pages = find_segment_pages(latest_file)
    seg_tables = extract_segment_tables(latest_file, seg_pages)
    segment_summary = summarize_segment_performance(seg_tables, ticker)

    # 5) Extract management guidance from MD&A
    mda_pages = find_mda_pages(latest_file)
    mda_text = extract_mda_text(latest_file, mda_pages)
    mgmt_guidance = summarize_management_guidance(mda_text, ticker)

    # 6) Scrape Yahoo Finance consensus
    yahoo_consensus = scrape_yahoo_finance_consensus(ticker)

    # 7) Determine net debt (prefer yfinance)
    yfinance_nd = get_yfinance_net_debt(ticker)
    if yfinance_nd is not None:
        net_debt = float(yfinance_nd)
        print(f"[INFO] Using yfinance net debt for {ticker}: ${net_debt:,.0f}")
    else:
        most_recent = historical_list[-1]
        total_debt = most_recent.get("total_debt") or 0.0
        cash_ce = most_recent.get("cash_ce") or 0.0
        net_debt = float(total_debt) - float(cash_ce)
        print(f"[WARNING] yfinance net debt lookup failed; using PDF parse net debt ({most_recent['year']}): ${net_debt:,.0f}")

    # 8) Build FCFF forecast
    fcf_proj = build_fcff_forecast(
        hist_year=last_hist_year,
        last_historical_fcf=last_hist_fcf,
        mgmt_guidance=mgmt_guidance,
        yahoo_consensus=yahoo_consensus,
        segment_summary_text=segment_summary
    )

    # 9) Scenario assumptions (growth_rates, waccs, terminal_multiples)
    growth_rates = {"bull": None, "base": None, "bear": None}
    if isinstance(assumptions, dict) and "growth_rates" in assumptions:
        gr = assumptions["growth_rates"]
        for s in ("bull", "base", "bear"):
            growth_rates[s] = gr.get(s, None)
    # Fill missing with default
    for s in ("bull", "base", "bear"):
        if growth_rates[s] is None:
            growth_rates[s] = 0.06  # default base growth; user can override

    waccs = {"bull": None, "base": None, "bear": None}
    terminal_mults = {"bull": None, "base": None, "bear": None}
    if isinstance(assumptions, dict) and "waccs" in assumptions and "terminal_multiples" in assumptions:
        wa = assumptions["waccs"]
        tm = assumptions["terminal_multiples"]
        for s in ("bull", "base", "bear"):
            waccs[s] = wa.get(s, None)
            terminal_mults[s] = tm.get(s, None)
    # Fill missing with defaults
    for s in ("bull", "base", "bear"):
        if waccs[s] is None:
            waccs[s] = 0.10  # default WACC
        if terminal_mults[s] is None:
            terminal_mults[s] = 12  # default terminal multiple

    # 10) Run DCF projection for each scenario
    dcf_results = {}
    for scenario in ("bull", "base", "bear"):
        wacc = float(waccs[scenario])
        tm = float(terminal_mults[scenario])

        projections = []
        for i in range(1, 6):
            year_i = last_hist_year + i
            fcf_i = fcf_proj.get(year_i, 0.0)
            projections.append({"year": year_i, "fcf": float(fcf_i)})

        terminal_value = projections[-1]["fcf"] * tm

        ev = sum(
            p["fcf"] / ((1 + wacc) ** idx)
            for idx, p in enumerate(projections, start=1)
        ) + terminal_value / ((1 + wacc) ** 5)

        equity_value = ev - net_debt
        npv_per_share = None
        if shares_outstanding:
            try:
                npv_per_share = float(equity_value) / float(shares_outstanding)
            except:
                npv_per_share = None

        dcf_results[scenario] = {
            "growth_rate": growth_rates[scenario],
            "wacc": wacc,
            "terminal_multiple": tm,
            "projections": projections,
            "terminal_value": terminal_value,
            "enterprise_value": ev,
            "equity_value": equity_value,
            "npv_per_share": npv_per_share
        }

    # 11) Write Excel workbook under ./reports
    wb = Workbook()
    wb.remove(wb.active)
    for scenario in ("bull", "base", "bear"):
        data = dcf_results[scenario]
        ws = wb.create_sheet(f"{scenario.capitalize()} Case")

        ws.append(["Scenario", scenario.capitalize()])
        ws.append(["Assumptions", "Value"])
        ws.append(["WACC", data["wacc"]])
        ws.append(["Terminal Multiple", data["terminal_multiple"]])
        ws.append([])

        ws.append(["Year", "Projected FCFF (USD)"])
        for p in data["projections"]:
            ws.append([p["year"], p["fcf"]])
        ws.append([])

        ws.append(["Terminal Value (Year 5)", data["terminal_value"]])
        ws.append([])
        ws.append(["Enterprise Value (PV of FCFF + Terminal)", data["enterprise_value"]])
        ws.append([])
        ws.append(["Net Debt (Debt - Cash)", net_debt])
        ws.append(["Equity Value", data["equity_value"]])
        ws.append([])
        if data["npv_per_share"] is not None:
            ws.append(["NPV per Share", data["npv_per_share"]])

    reports_dir = os.path.join(os.getcwd(), "reports")
    os.makedirs(reports_dir, exist_ok=True)
    filename = f"{ticker}_EnhancedDCF_{int(time.time())}.xlsx"
    output_path = os.path.join(reports_dir, filename)
    wb.save(output_path)

    # 12) Build summary_dict
    summary = {
        "company_name": company_name,
        "ticker": ticker,
        "current_share_price": current_price,
        "shares_outstanding": shares_outstanding,
        "net_debt": net_debt,
        "historical_years": hist_years,
        "historical_fcfs": hist_fcfs,
        "historical_revenues": hist_revs,
        "historical_net_incomes": hist_nis,
        "segment_summary": segment_summary,
        "management_guidance": mgmt_guidance,
        "yahoo_consensus": yahoo_consensus,
        "fcff_forecast": fcf_proj,
        "scenarios": {}
    }

    # Ensure all three scenarios exist
    for scenario in ("bull", "base", "bear"):
        data = dcf_results.get(scenario, {
            "enterprise_value": None,
            "equity_value": None,
            "npv_per_share": None,
            "terminal_value": None,
            "wacc": None,
            "growth_rate": None,
            "terminal_multiple": None
        })
        summary["scenarios"][scenario] = {
            "enterprise_value": data["enterprise_value"],
            "equity_value": data["equity_value"],
            "npv_per_share": data["npv_per_share"],
            "terminal_value": data["terminal_value"],
            "wacc": data["wacc"],
            "growth_rate": data["growth_rate"],
            "terminal_multiple": data["terminal_multiple"]
        }

    return output_path, summary

