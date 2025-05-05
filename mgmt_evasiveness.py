# mgmt_evasiveness.py

import os
import re
import io
import json
import requests
import pandas as pd
from datetime import timedelta
from openai import OpenAI
from flask import (
    Flask, request, jsonify, send_file
)

# ─── Configuration ─────────────────────────────────────────────────────────────
# Load keys from environment
FMP_API_KEY      = os.environ["FMP_API_KEY"]
DEEPSEEK_API_KEY = os.environ["DEEPSEEK_API_KEY"]
if not FMP_API_KEY or not DEEPSEEK_API_KEY:
    raise RuntimeError("Must set FMP_API_KEY and DEEPSEEK_API_KEY in the environment")


# ─── DeepSeek & FMP Clients ─────────────────────────────────────────────────────
ds_client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=os.getenv("DEEPSEEK_API_URL", "https://api.deepseek.com/v1"))

def resolve_ticker(company_name: str) -> str:
    prompt = (
        f"What is the FMP-compatible ticker for this company: '{company_name}'? "
        "Return only the ticker symbol."
    )
    resp = ds_client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role":"user","content":prompt}],
        temperature=0
    )
    return resp.choices[0].message.content.strip().upper()

def fetch_transcript(ticker: str, year: int, quarter: int) -> dict:
    url = (
        f"https://financialmodelingprep.com/api/v3/earning_call_transcript/"
        f"{ticker}?year={year}&quarter={quarter}&apikey={FMP_API_KEY}"
    )
    r = requests.get(url)
    data = r.json() if r.status_code==200 else []
    return data[0] if isinstance(data, list) and data else None

def get_price(ticker: str, date_str: str) -> float:
    end = pd.to_datetime(date_str)
    start = end - timedelta(days=7)
    url = (
        f"https://financialmodelingprep.com/api/v3/historical-price-full/"
        f"{ticker}?from={start:%Y-%m-%d}&to={end:%Y-%m-%d}&apikey={FMP_API_KEY}"
    )
    r = requests.get(url)
    hist = (r.json() or {}).get("historical", [])
    if not hist:
        return None
    df = pd.DataFrame(hist)
    df.date = pd.to_datetime(df.date)
    df = df[df.date<=end].sort_values("date",ascending=False)
    return float(df.iloc[0].close) if not df.empty else None

# ─── Transcript Parsing & Evasiveness Analysis ─────────────────────────────────
def parse_management_turns(content: str) -> list[str]:
    lines = [ln.strip() for ln in content.splitlines()
             if ":" in ln and len(ln.strip())>15]
    mgmt = [
        ln for ln in lines
        if not re.search(r'\b(operator|analyst|moderator|host|coordinator|caller)\b',
                         ln.split(":",1)[0], re.IGNORECASE)
    ]
    return mgmt

def analyze_evasiveness(turns: list[str], batch_size: int=20):
    """
    1) Sends batches to DeepSeek with a richer prompt (4 categories).
    2) Extracts (statement, category, reason) triples.
    3) Computes a 0–10 normalized score.
    Returns (df_statements, score: float).
    """
    total = len(turns)
    responses = []

    category_map = {"2":1, "3":2}

    for i in range(0, total, batch_size):
        batch = turns[i:i+batch_size]
        prompt = f"""
        You are a forensic analyst specializing in executive communication.

        Below are statements made by company management during an earnings call. Each statement may respond to a question from an analyst.

        Your task is to identify **only those speaker turns that show evasiveness**. For each speaker turn, carefully read the full text and evaluate whether the speaker is:

        1. Avoiding or refusing to answer a direct question
        2. Providing vague or overly generalized statements without specifics
        3. Shifting blame, deflecting responsibility, or overstating confidence without justification
        4. Contradicting earlier disclosed facts or using ambiguous qualifiers

        For each evasive case, return:
        ---
        Response: The **most relevant 2–3 consecutive sentences** that reflect the evasiveness (not random fragments)
        Category:
        - 2 = Somewhat evasive – vague or generic but at least partially addresses the topic
        - 3 = Clearly evasive – refuses to answer, contradicts info, or uses avoidance tactics
        Reason: A **specific and concise explanation** of why the excerpt is evasive. Mention what information was avoided, what ambiguity exists, or what signals a lack of transparency.

        Output format should be:
        ---
        Response: "..."
        Category: 2
        Reason: ...

        If no part of the speaker turn is evasive, skip it.

        Statements:
        {chr(10).join([f"{j+1}. {line}" for j, line in enumerate(batch)])}
        """
        try:
            resp = ds_client.chat.completions.create(
                model="deepseek-chat",
                messages=[
                    {"role":"system","content":"You are an expert earnings call analyst."},
                    {"role":"user","content":prompt}
                ],
                temperature=0
            )
            text = resp.choices[0].message.content
            matches = re.findall(
                r'Response:\s*"(.+?)"\s*Category:\s*([23])\s*Reason:\s*(.+?)(?=(?:\n\n|$))',
                text, re.DOTALL
            )
            responses += matches
        except Exception:
            # skip failed batch
            continue

    # build DataFrame
    df = pd.DataFrame(responses, columns=["Statement","Category","Reason"])
    if df.empty:
        return df, 0.0

    df.Category = df.Category.astype(int)
    pts = df.Category.map({2:1,3:2}).sum()
    score = round(pts / (total*2) * 10, 2)
    return df, score

def extract_evasiveness_parameters_from_text(q: str):
    """
    Try several natural‐language patterns to pull out
    (company, quarter(int), year(int)).
    """
    # pattern #1: “… for <Company> Q<1-4> <YYYY>”
    m = re.search(r'\bfor\s+(.+?)\s+q([1-4])\s+(\d{4})\b', q, re.I)
    if m:
        return m.group(1).strip(), int(m.group(2)), int(m.group(3))
    
    # pattern #2: “… <Company>’s fourth quarter of 2024 …”
    m = re.search(
      r"([\w &]+?)[’']?s?\s+(?:first|second|third|fourth|1st|2nd|3rd|4th)\s+quarter\s+(?:of\s*)?(\d{4})",
      q, re.I
    )
    if m:
        comp, year = m.group(1).strip(), int(m.group(2))
        ord_map = {'first':1,'1st':1,'second':2,'2nd':2,'third':3,'3rd':3,'fourth':4,'4th':4}
        ord_word = re.search(r'(first|second|third|fourth|1st|2nd|3rd|4th)', q, re.I).group(1).lower()
        return comp, ord_map[ord_word], year

    # pattern #3: “… quarter 4 2024 for <Company>”
    m = re.search(r'quarter\s*([1-4])\s*(\d{4}).+?\bfor\s+(.+)', q, re.I)
    if m:
        return m.group(3).strip(), int(m.group(1)), int(m.group(2))

    return None, None, None


def generate_evasiveness_report(company, year, quarter, outdir):
    """
    Full pipeline: resolve ticker, transcript, parse, analyze,
    fetch price, write an XLSX, return its path.
    """
    ticker = resolve_ticker(company)
    transcript = fetch_transcript(ticker, year, quarter)
    if not transcript:
        raise FileNotFoundError(f"No transcript for {ticker} Q{quarter} {year}")

    turns = parse_management_turns(transcript["content"])
    df_statements, score = analyze_evasiveness(turns)
    price = get_price(ticker, transcript["date"])

    summary = pd.DataFrame([{
      "Company": company,
      "Ticker": ticker,
      "Year": year,
      "Quarter": quarter,
      "TranscriptDate": transcript["date"],
      "EvasivenessScore": score,
      "SharePrice": price
    }])

    os.makedirs(outdir, exist_ok=True)
    fname = f"{ticker}_Q{quarter}_{year}.xlsx"
    path = os.path.join(outdir, fname)
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        summary.to_excel(writer, index=False, sheet_name="Summary")
        if not df_statements.empty:
            df_statements.to_excel(
                writer, index=False, sheet_name="Evasive Statements"
            )
    return path

def merge_reports(xlsx_paths: list[str], out_path: str):
    """
    Given a list of summary-Excel paths, concatenate their Summary sheets
    and write to `out_path`.
    """
    dfs = []
    for p in xlsx_paths:
        df = pd.read_excel(p, sheet_name="Summary")
        dfs.append(df)
    merged = pd.concat(dfs, ignore_index=True)
    merged.sort_values("TranscriptDate", inplace=True)
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        merged.to_excel(writer, index=False, sheet_name="Merged")
    return out_path
