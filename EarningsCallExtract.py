import os
import requests
import re

# API keys
DEEPSEEK_API_KEY = os.environ["DEEPSEEK_API_KEY"]
FMP_API_KEY = os.environ["FMP_API_KEY"]

FMP_URL_TEMPLATE = (
    "https://financialmodelingprep.com/api/v3/earning_call_transcript/{ticker}?year={year}&quarter={quarter}&apikey=" + FMP_API_KEY
)

def extract_ticker_deepseek(query):
    """Uses DeepSeek to extract the stock ticker symbol from query."""
    try:
        response = requests.post(
            "https://api.deepseek.com/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
                "Content-Type": "application/json"
            },
            json={
                "model": "deepseek-chat",
                "messages": [
                    {
                        "role": "system",
                        "content": "Extract the stock ticker symbol (e.g., AAPL, MSFT, TSLA) from this query. Only return the ticker, nothing else."
                    },
                    {"role": "user", "content": query}
                ],
                "temperature": 0
            }
        )
        if response.status_code == 200:
            return response.json()["choices"][0]["message"]["content"].strip().upper()
        else:
            print("DeepSeek API error:", response.text)
            return None
    except Exception as e:
        print("DeepSeek error:", str(e))
        return None

def fetch_earning_call_transcript(ticker, year, quarter):
    url = FMP_URL_TEMPLATE.format(ticker=ticker, year=year, quarter=quarter)
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if isinstance(data, list) and len(data) > 0 and 'content' in data[0]:
            return data[0]['content']
    return ""

def extract_sentences_with_numbers(text):
    """
    Extracts sentences containing both numbers and relevant financial/business keywords.
    """
    keywords = [
        "revenue", "sales", "margin", "profit", "earnings", "guidance", "outlook", "forecast",
        "free cash flow", "cash flow", "EBIT", "operating", "inventory", "backlog", "dividend",
        "return", "cost", "expense", "growth", "performance", "expect", "project", "estimate",
        "competition", "market", "industry", "volume", "price", "demand", "supply"
    ]

    pattern = re.compile(r'\b(?:' + '|'.join(map(re.escape, keywords)) + r')\b', flags=re.IGNORECASE)
    sentences = re.split(r'(?<=[.!?])\s+', text)
    
    return [
        sentence for sentence in sentences
        if re.search(r'\d+', sentence) and pattern.search(sentence)
    ]

def filter_by_keywords(sentences, keywords):
    """Filters the numeric sentences by presence of any keyword."""
    keywords_lower = [k.lower() for k in keywords]
    return [s for s in sentences if any(k in s.lower() for k in keywords_lower)]

def parse_keywords_from_query(query):
    """Parses a comma-separated list of keywords from query text (if present)."""
    match = re.search(r'keywords?\s*[:\-]?\s*(.+)', query, re.IGNORECASE)
    if match:
        return [k.strip() for k in match.group(1).split(',') if k.strip()]
    return []

def get_earnings_call_data(query):
    """Main handler for extracting numeric + keyword-based sentences."""
    query = re.sub(r"^(also|please|can you|kindly|just)\b[,:]?\s*", "", query.strip(), flags=re.IGNORECASE)
    ticker = extract_ticker_deepseek(query)
    if not ticker:
        return {"error": "Could not extract ticker from the query."}
    
    match = re.search(r'\bQ([1-4])\b[\s,:\-]*(\d{4})', query, re.IGNORECASE)
    if not match:
        return {"error": "Could not extract quarter and year."}

    quarter, year = match.groups()
    transcript = fetch_earning_call_transcript(ticker, int(year), int(quarter))
    if not transcript:
        return {"error": "No earnings call transcript available for this request."}

    numeric_sentences = extract_sentences_with_numbers(transcript)
    keywords = parse_keywords_from_query(query)

    if keywords:
        filtered = filter_by_keywords(numeric_sentences, keywords)
        return {
            "company": ticker,
            "quarter": quarter,
            "year": year,
            "keywords": keywords,
            "sentences": filtered
        }
    else:
        return {
            "company": ticker,
            "quarter": quarter,
            "year": year,
            "sentences": numeric_sentences
        }
