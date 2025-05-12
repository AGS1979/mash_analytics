import os
import requests
import json
from openai import OpenAI

# Load API keys and client setup
FMP_API_KEY        = os.environ["FMP_API_KEY"]
DEEPSEEK_API_KEY   = os.environ["DEEPSEEK_API_KEY"]
DEEPSEEK_API_BASE  = os.getenv("DEEPSEEK_API_BASE", "https://api.deepseek.com")
OPENAI_API_KEY     = os.environ["OPENAI_API_KEY"]
OPENAI_API_BASE    = os.getenv("OPENAI_API_BASE", "https://api.openai.com/v1")

# Clients
deepseek_client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_API_BASE)
client_gpt      = OpenAI(api_key=OPENAI_API_KEY, base_url=OPENAI_API_BASE)


def extract_ticker_and_period(query):
    """Use DeepSeek to extract ticker and quarter/year from natural language query."""
    prompt = f"""
You are a financial assistant. Extract the stock ticker, quarter, and year from the user query below. 
If the quarter and year are not explicitly mentioned, infer if the query is asking for the most recent quarter.

Return a JSON with:
- ticker: string
- quarter: integer (1 to 4)
- year: integer

Query: {query}
"""
    try:
        response = deepseek_client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "user", "content": prompt}]
        )
        content = response.choices[0].message.content.strip()
        result = json.loads(content)
        return result["ticker"], int(result["quarter"]), int(result["year"])
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


def analyze_guidance_change(transcript):
    """Ask GPT if full-year guidance was upgraded, downgraded, or maintained."""
    prompt = f"""
You are an earnings call analyst. Analyze the transcript below and determine whether the company upgraded, downgraded, or maintained its **full-year guidance** (e.g., for revenue, EPS, margins).

Clearly state one of the following outcomes:
- "Upgraded full-year guidance"
- "Downgraded full-year guidance"
- "Maintained full-year guidance"
- Or, "Full-year guidance not mentioned"

Provide a short explanation supporting your conclusion.

Transcript:
\"\"\"
{transcript}
\"\"\"
"""
    try:
        response = client_gpt.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"❌ GPT analysis failed: {str(e)}"


def process_guidance_query(query):
    ticker, quarter, year = extract_ticker_and_period(query)
    if not ticker or not year or not quarter:
        return "Unable to extract ticker and time period from your query."

    transcript = fetch_transcript_from_fmp(ticker, year, quarter)
    if not transcript:
        return f"No transcript found for {ticker} in Q{quarter} {year}."

    result = analyze_guidance_change(transcript)
    return result
