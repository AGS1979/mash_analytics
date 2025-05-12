import os
import requests
import json
import re
from textwrap import wrap
from openai import OpenAI

# Load API keys and client setup
FMP_API_KEY        = os.environ["FMP_API_KEY"]
DEEPSEEK_API_KEY   = os.environ["DEEPSEEK_API_KEY"]
DEEPSEEK_API_BASE  = os.getenv("DEEPSEEK_API_BASE", "https://api.deepseek.com")

# DeepSeek Client
deepseek_client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_API_BASE)


def extract_ticker_and_period(query):
    """Use DeepSeek to extract ticker, quarter, and year."""
    prompt = f"""
You are a financial assistant. Extract the stock ticker, quarter, and year from the user query below. 
If the quarter and year are not explicitly mentioned, infer if the query is asking for the most recent quarter.

Return ONLY this JSON:
{{"ticker": "XXX", "quarter": Q, "year": YYYY}}

Query: {query}
"""

    try:
        response = deepseek_client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}]
        )
        content = response.choices[0].message.content.strip()
        try:
            result = json.loads(content)
            return result["ticker"], int(result["quarter"]), int(result["year"])
        except Exception:
            print("🔎 Raw response from DeepSeek:", content)
            # Fallback: regex parse
            try:
                ticker = re.search(r'"?ticker"?\s*[:=]\s*"?([A-Z.]+)"?', content).group(1)
                quarter = int(re.search(r'"?quarter"?\s*[:=]\s*"?([1-4])"?', content).group(1))
                year = int(re.search(r'"?year"?\s*[:=]\s*"?(\d{4})"?', content).group(1))
                return ticker, quarter, year
            except Exception as e:
                print("❌ Regex fallback failed:", e)
                return None, None, None

    except Exception as e:
        print("❌ DeepSeek parsing failed:", e)
        return None, None, None


def fetch_transcript_from_fmp(ticker, year, quarter):
    url = f"https://financialmodelingprep.com/api/v3/earning_call_transcript/{ticker}?year={year}&quarter={quarter}&apikey={FMP_API_KEY}"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if data and isinstance(data, list) and "content" in data[0]:
            return data[0]["content"]
    return None


def summarize_long_transcript(transcript):
    """Split and summarize transcript in chunks using DeepSeek."""
    chunks = wrap(transcript, 2000)
    summaries = []

    for chunk in chunks:
        prompt = f"Summarize this earnings call excerpt in 2–3 sentences:\n\n{chunk}"
        response = deepseek_client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}]
        )
        summaries.append(response.choices[0].message.content.strip())

    return "\n".join(summaries)


def analyze_guidance_change(transcript):
    """Use DeepSeek to classify the guidance change type."""
    summarized = summarize_long_transcript(transcript)

    prompt = f"""
You are an earnings call analyst. Based on the **summarized transcript** below, determine whether the company upgraded, downgraded, or maintained its full-year guidance (for revenue, EPS, margins, etc.).

Respond with exactly one of:
- "Upgraded full-year guidance"
- "Downgraded full-year guidance"
- "Maintained full-year guidance"
- "Full-year guidance not mentioned"

Then give a 2-3 sentence explanation.

Summarized Transcript:
\"\"\"
{summarized}
\"\"\"
"""
    try:
        response = deepseek_client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"❌ DeepSeek analysis failed: {str(e)}"


def process_guidance_query(query):
    ticker, quarter, year = extract_ticker_and_period(query)
    if not ticker or not year or not quarter:
        return "Unable to extract ticker and time period from your query."

    transcript = fetch_transcript_from_fmp(ticker, year, quarter)
    if not transcript:
        return f"No transcript found for {ticker} in Q{quarter} {year}."

    result = analyze_guidance_change(transcript)
    return result
