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
    prompt = f"""
You are a financial assistant. Given the user query below, extract the stock ticker, quarter, and year.

Only respond with a valid JSON object like this:
{{
  "ticker": "TSLA",
  "quarter": 1,
  "year": 2025
}}

Query: {query}
"""

    try:
        response = deepseek_client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}]
        )
        content = response.choices[0].message.content.strip()

        # ✅ Clean up any markdown/code formatting
        content = content.replace("```json", "").replace("```", "").strip()

        # Parse the JSON
        result = json.loads(content)
        return result["ticker"], int(result["quarter"]), int(result["year"])
    except Exception as e:
        print("❌ DeepSeek parsing failed:", e)
        print("🔎 Raw DeepSeek response:", content)
        return None, None, None


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
