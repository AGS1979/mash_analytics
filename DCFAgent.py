# DCFAgent.py

import os
import json
import time
import re
import tempfile
import requests

import pandas as pd
import numpy as np

from PyPDF2 import PdfReader
import camelot             # for table extraction
from openpyxl import Workbook

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
    Sends `prompt` to DeepSeek Chat and returns the assistant's reply text.
    """
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "deepseek-chat",
        "messages": [ {"role": "user", "content": prompt} ],
        "max_tokens": max_tokens
    }
    resp = requests.post(DEEPSEEK_CHAT_URL, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]

# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Extract raw text from a list of PDFs
# ────────────────────────────────────────────────────────────────────────────────
def extract_text_from_pdfs(file_paths: list[str]) -> str:
    """
    Concatenates all textual content from the given PDF file paths.
    """
    all_text = []
    for path in file_paths:
        try:
            reader = PdfReader(path)
            for page in reader.pages:
                txt = page.extract_text()
                if txt:
                    all_text.append(txt)
        except Exception:
            continue
    return "\n".join(all_text)

# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Identify page numbers containing a given keyword
# ────────────────────────────────────────────────────────────────────────────────
def find_pages_with_keyword(pdf_path: str, keyword: str) -> list[int]:
    """
    Returns a list of 0-based page indices where `keyword` appears in the text.
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
# HELPER: Extract financial table (as pandas DataFrame) from given PDF pages
# ────────────────────────────────────────────────────────────────────────────────
def extract_tables_from_pdf(pdf_path: str, page_indices: list[int]) -> list[pd.DataFrame]:
    """
    Uses Camelot to extract all tables on the specified 1-based pages.
    Camelot requires pages in '1,2,5' format (1-based).
    Returns a list of DataFrames.
    """
    dfs = []
    if not page_indices:
        return dfs

    # Camelot wants pages as a comma-separated string of 1-based indices
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
    Looks for a row containing 'Free Cash Flow' (case-insensitive) in the DataFrame `df`.
    If found, tries to parse the number from the next column.
    Returns a float (in USD) or None if not found/parsable.
    """
    for row_idx in range(len(df)):
        row = df.iloc[row_idx].tolist()
        # Combine all row cells into one string for matching
        row_str = " ".join(str(cell) for cell in row)
        if re.search(r"free\s+cash\s+flow", row_str, re.IGNORECASE):
            # Assume the numeric value is in one of the subsequent columns
            for cell in row[1:]:
                # Remove commas, dollar signs, parentheses
                s = str(cell).replace(",", "").replace("$", "").replace("(", "-").replace(")", "")
                try:
                    return float(s)
                except Exception:
                    continue
    return None

# ────────────────────────────────────────────────────────────────────────────────
# HELPER: Attempt to parse Revenue & Net Income from Income Statement tables
# ────────────────────────────────────────────────────────────────────────────────
def parse_income_from_df(df: pd.DataFrame) -> tuple[float | None, float | None]:
    """
    Looks for 'Revenue' and 'Net Income' rows (case-insensitive) in `df`.
    Returns (revenue, net_income) in USD, or (None, None) if not found.
    """
    rev = None
    ni = None
    for row_idx in range(len(df)):
        row = df.iloc[row_idx].tolist()
        row_str = " ".join(str(cell) for cell in row)
        if re.search(r"(^|\s)revenue($|\s)", row_str, re.IGNORECASE):
            # parse number from next columns
            for cell in row[1:]:
                s = str(cell).replace(",", "").replace("$", "").replace("(", "-").replace(")", "")
                try:
                    rev = float(s)
                    break
                except Exception:
                    continue
        if re.search(r"(^|\s)net\s+income($|\s)", row_str, re.IGNORECASE):
            for cell in row[1:]:
                s = str(cell).replace(",", "").replace("$", "").replace("(", "-").replace(")", "")
                try:
                    ni = float(s)
                    break
                except Exception:
                    continue
        if rev is not None and ni is not None:
            break
    return rev, ni

# ────────────────────────────────────────────────────────────────────────────────
# CORE HELPER: Extract historical financials from a single PDF
# ────────────────────────────────────────────────────────────────────────────────
def extract_financials_from_pdf(pdf_path: str) -> dict | None:
    """
    Attempts to extract {year, revenue, net_income, free_cash_flow} from a 10-K/10-Q PDF.
    Procedure:
      1. Scan the PDF text for a phrase like "For the year ended December 31, 2023"
         to deduce `year`.
      2. Look for pages containing "Consolidated Statements of Cash Flows" → parse tables there → find Free Cash Flow.
      3. Look for pages containing "Consolidated Statements of Income" (or "Income Statement" / "Statement of Operations") → parse those → find revenue & net income.
    Returns a dict {"year": int, "revenue": float, "net_income": float, "free_cash_flow": float}
    or None if it fails.
    """
    try:
        # 1) Extract raw text to find the year
        reader = PdfReader(pdf_path)
        full_text = ""
        for p in reader.pages:
            txt = p.extract_text() or ""
            full_text += "\n" + txt

        # Look for "For the year ended December 31, 2023"
        year_match = re.search(r"For\s+the\s+year\s+ended\s+December\s+31,\s*(\d{4})", full_text, re.IGNORECASE)
        if year_match:
            year = int(year_match.group(1))
        else:
            # Fallback: look for standalone year at top of financial section
            y2 = re.search(r"(\d{4})\s+Consolidated\s+Balance\s+Sheets", full_text, re.IGNORECASE)
            year = int(y2.group(1)) if y2 else None

        if year is None:
            return None

        # 2) Find pages with Cash Flow Statement
        cf_pages = find_pages_with_keyword(pdf_path, "Cash Flows")
        tables_cf = extract_tables_from_pdf(pdf_path, cf_pages)
        fcf = None
        for df in tables_cf:
            val = parse_fcf_from_df(df)
            if val is not None:
                fcf = val
                break

        # 3) Find pages with Income Statement
        inc_pages = find_pages_with_keyword(pdf_path, "Income Statement|Statement of Operations",)
        tables_inc = extract_tables_from_pdf(pdf_path, inc_pages)
        revenue = None
        net_income = None
        for df in tables_inc:
            r, ni = parse_income_from_df(df)
            if r is not None:
                revenue = r
            if ni is not None:
                net_income = ni
            if revenue is not None and net_income is not None:
                break

        if revenue is None and net_income is None and fcf is None:
            return None

        return {
            "year": year,
            "revenue": revenue,
            "net_income": net_income,
            "free_cash_flow": fcf
        }
    except Exception:
        return None

# ────────────────────────────────────────────────────────────────────────────────
# MAIN FUNCTION: run_dcf_model (company_name-based)
# ────────────────────────────────────────────────────────────────────────────────
def run_dcf_model(
    company_name: str,
    assumptions: dict,
    file_paths: list[str]
) -> tuple[str, dict]:
    """
    Performs a 3-case (Bull / Base / Bear) DCF valuation for the given `company_name`.
    1) Uses DeepSeek to convert company_name → ticker.
    2) Extracts historical financials from each uploaded PDF/DOCX via extract_financials_from_pdf.
    3) Aggregates at most 5 most recent years of financials.
    4) Uses user-supplied assumption dict (or defaults) for growth_rates, waccs, terminal_multiples.
    5) Projects FCF for 5 years + terminal value → discounts at scenario WACC → computes NPV.
    6) Writes an Excel under ./reports named "<TICKER>_DCF_<timestamp>.xlsx".
    7) Returns (output_path, summary_dict).

    Input:
      company_name: str, e.g. "Apple Inc." or "Microsoft Corporation"
      assumptions: dict, e.g. {
         "growth_rates": {"bull":0.08, "base":0.05, "bear":0.02},
         "waccs": {"bull":0.09, "base":0.10, "bear":0.11},
         "terminal_multiples": {"bull":14, "base":12, "bear":10}
      }
      file_paths: list of locally saved file paths to user-uploaded annual reports, 10-Ks, etc.

    Output:
      (output_path, summary_dict)
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
    # 2) Extract pandas dicts for each PDF file (year, revenue, net_income, fcf)
    # ────────────────────────────────────────────────────────────────────────────
    extracted = []
    for path in file_paths:
        fin = extract_financials_from_pdf(path)
        if fin is not None and all(k in fin for k in ("year", "free_cash_flow")):
            extracted.append(fin)
    if not extracted:
        # If nothing could be parsed, abort
        raise ValueError("No valid financial data could be extracted from uploaded files.")

    # Keep only 5 most recent years
    extracted = sorted(extracted, key=lambda x: x["year"], reverse=True)[:5]
    historical_list = sorted(extracted, key=lambda x: x["year"])  # ascending

    # Extract arrays
    hist_years = [e["year"] for e in historical_list]
    hist_fcfs = [e["free_cash_flow"] for e in historical_list]
    hist_revs = [e.get("revenue") for e in historical_list]
    hist_nis = [e.get("net_income") for e in historical_list]
    last_year = hist_years[-1]
    last_fcf = hist_fcfs[-1]

    # ────────────────────────────────────────────────────────────────────────────
    # 3) Scenario assumptions: growth_rates, waccs, terminal_multiples
    # ────────────────────────────────────────────────────────────────────────────
    # Fill gaps with DeepSeek prompts if missing

    # 3a) Growth Rates
    if "growth_rates" in assumptions:
        growth_rates = {
            "bull": assumptions["growth_rates"].get("bull"),
            "base": assumptions["growth_rates"].get("base"),
            "bear": assumptions["growth_rates"].get("bear"),
        }
    else:
        # Default fallback: caret sign CAGR of last 3 years of FCF → use DeepSeek to tweak
        gr_prompt = (
            f"For ticker {ticker}, the historical Free Cash Flows (USD) for years "
            f"{json.dumps(historical_list, indent=2)}\n"
            f"Propose forward 5-year annual FCF growth rates under bull, base, bear. "
            f"Reply with JSON {{\"bull\":0.XX,\"base\":0.XX,\"bear\":0.XX}}."
        )
        raw_gr = deepseek_chat(gr_prompt, max_tokens=256)
        try:
            growth_rates = json.loads(raw_gr)
        except Exception:
            # Naive defaults: 5% base, +/- 3% for bull/bear
            growth_rates = {
                "base": 0.05,
                "bull": 0.08,
                "bear": 0.02
            }
    # Fill any missing
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
        wa_prompt = (
            f"For ticker {ticker}, given historical financials {json.dumps(historical_list[-3:], indent=2)},\n"
            f"provide WACC and terminal multiples under bull, base, bear. "
            f"Reply JSON {{\"bull\":{{\"wacc\":0.XX,\"terminal_multiple\":YY}}, "
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
            waccs = {"bull": 0.09, "base": 0.10, "bear": 0.11}
            terminal_mults = {"bull": 14, "base": 12, "bear": 10}

    for scenario in ["bull", "base", "bear"]:
        if waccs.get(scenario) is None:
            waccs[scenario] = waccs.get("base", 0.10)
        if terminal_mults.get(scenario) is None:
            terminal_mults[scenario] = terminal_mults.get("base", 12)

    # ────────────────────────────────────────────────────────────────────────────
    # 4) PROJECT FCF & DISCOUNT for EACH SCENARIO
    # ────────────────────────────────────────────────────────────────────────────
    dcf_results = {}
    for scenario in ["bull", "base", "bear"]:
        gr = growth_rates[scenario]
        wacc = waccs[scenario]
        tm = terminal_mults[scenario]

        # Project 5 years of FCF
        projections = []
        for i in range(1, 6):
            year_i = last_year + i
            fcf_i = last_fcf * ((1 + gr) ** i)
            projections.append({"year": year_i, "fcf": fcf_i})

        # Terminal value at year 5
        terminal_value = projections[-1]["fcf"] * tm

        # Discount each year's FCF + terminal
        npv = sum(
            proj["fcf"] / ((1 + wacc) ** j)
            for j, proj in enumerate(projections, start=1)
        )
        npv += terminal_value / ((1 + wacc) ** 5)

        dcf_results[scenario] = {
            "growth_rate": gr,
            "wacc": wacc,
            "terminal_multiple": tm,
            "projections": projections,
            "terminal_value": terminal_value,
            "npv": npv
        }

    # ────────────────────────────────────────────────────────────────────────────
    # 5) CREATE EXCEL WORKBOOK UNDER ./reports
    # ────────────────────────────────────────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    for scenario in ["bull", "base", "bear"]:
        data = dcf_results[scenario]
        ws = wb.create_sheet(f"{scenario.capitalize()} Case")

        ws.append(["Scenario", scenario.capitalize()])
        ws.append(["Assumptions", "Value"])
        ws.append(["Growth Rate", data["growth_rate"]])
        ws.append(["WACC", data["wacc"]])
        ws.append(["Terminal Multiple", data["terminal_multiple"]])
        ws.append([])

        ws.append(["Year", "Projected FCF (USD)"])
        for proj in data["projections"]:
            ws.append([proj["year"], proj["fcf"]])

        ws.append([])
        ws.append(["Terminal Value (Year 5)", data["terminal_value"]])
        ws.append([])
        ws.append(["NPV of Cash Flows", data["npv"]])

    # Ensure reports directory exists
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
        "historical_years": hist_years,
        "historical_fcfs": hist_fcfs,
        "historical_revenues": hist_revs,
        "historical_net_incomes": hist_nis,
        "scenarios": {
            scenario: {
                "npv": dcf_results[scenario]["npv"],
                "terminal_value": dcf_results[scenario]["terminal_value"],
                "wacc": dcf_results[scenario]["wacc"],
                "growth_rate": dcf_results[scenario]["growth_rate"],
                "terminal_multiple": dcf_results[scenario]["terminal_multiple"]
            }
            for scenario in ["bull", "base", "bear"]
        }
    }

    return output_path, summary_dict
