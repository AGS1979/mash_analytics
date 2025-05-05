import os
import requests
import re
import openai
import csv

# Load OpenAI API Key from environment
openai.api_key = os.environ["OPENAI_API_KEY"]



# FMP API Key from environment
FMP_API_KEY = os.environ["FMP_API_KEY"]
FMP_URL_TEMPLATE = ("https://financialmodelingprep.com/api/v3/earning_call_transcript/{ticker}?year={year}&quarter={quarter}&apikey=FMP_API_KEY")

def load_ticker_list(file_path="company_tickers.csv"):
    """Loads a dictionary mapping company names to tickers."""
    tickers_dict = {}
    try:
        with open(file_path, mode='r') as file:
            reader = csv.reader(file)
            for row in reader:
                if len(row) < 2 or not row[0].strip() or not row[1].strip():
                    continue
                company_name, ticker = row[:2]
                tickers_dict[company_name.strip().upper()] = ticker.strip().upper()
        return tickers_dict
    except FileNotFoundError:
        return {}

def extract_company_name(query):
    """Extracts the company name from the query using OpenAI."""
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a financial assistant. Extract the company name from the user's query."},
                {"role": "user", "content": query}
            ]
        )
        extracted_text = response['choices'][0]['message']['content'].strip()
        cleaned_name = re.sub(r"(The company name.*is\s*|Company name\s*is\s*)", "", extracted_text)
        return cleaned_name.strip()
    except Exception as e:
        return None

def match_company_to_ticker(company_name, tickers_dict):
    """Matches extracted company name to the correct ticker."""
    company_name = company_name.upper()
    if company_name in tickers_dict:
        return tickers_dict[company_name]
    
    matches = [ticker for name, ticker in tickers_dict.items() if company_name in name.upper()]
    return matches[0] if matches else None

def fetch_earning_call_transcript(ticker, year, quarter):
    """Fetches earnings call transcript from FMP."""
    url = FMP_URL_TEMPLATE.format(ticker=ticker, year=year, quarter=quarter)
    response = requests.get(url)

    if response.status_code == 200:
        data = response.json()
        if isinstance(data, list) and len(data) > 0 and 'content' in data[0]:
            return data[0]['content']
    return ""

def extract_sentences_with_numbers(text):
    """Extracts sentences containing numerical figures from the earnings call."""
    sentences = re.split(r'(?<=[.!?])\s+', text)
    return [sentence for sentence in sentences if re.search(r'\d+', sentence)]

def get_earnings_call_data(query):
    """Main function that extracts figures from an earnings call based on the query."""
    tickers_dict = load_ticker_list("company_tickers.csv")

    company_name = extract_company_name(query)
    if not company_name:
        return {"error": "Could not extract company name from the query."}

    ticker = match_company_to_ticker(company_name, tickers_dict)
    if not ticker:
        return {"error": "No valid ticker found for the given company name."}

    match = re.search(r'Q([1-4])\s+(\d{4})', query, re.IGNORECASE)
    if not match:
        return {"error": "Could not extract quarter and year."}

    quarter, year = match.groups()
    transcript = fetch_earning_call_transcript(ticker, int(year), int(quarter))

    if not transcript:
        return {"error": "No earnings call transcript available for this request."}

    extracted_sentences = extract_sentences_with_numbers(transcript)
    return {"company": ticker, "quarter": quarter, "year": year, "sentences": extracted_sentences}
