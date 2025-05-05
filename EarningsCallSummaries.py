import os
import csv
from flask import send_file
import requests
import pandas as pd
from textblob import TextBlob
import time
import json
import os
import sys
import io
import openai
import re

# Load API keys from environment
api_key = os.environ["FMP_API_KEY"]
openai.api_key = os.environ["OPENAI_API_KEY"]


def extract_company_name(query):
    """Extracts the company name from the query using OpenAI."""
    try:
        print(f"Extracting company name from query: {query}")
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",  
            messages=[
                {"role": "system", "content": "You are a financial assistant. Extract the company name from the user query."},
                {"role": "user", "content": query}
            ]
        )
        extracted_text = response['choices'][0]['message']['content'].strip()
        print(f"Extracted text: {extracted_text}")

        # Clean the extracted text (remove any extra commentary or punctuation)
        cleaned_name = re.sub(
            r"(The company name.*is\s*|Company name\s*is\s*|The company mentioned in your query.*is\s*)",
            "", extracted_text
        )
        cleaned_name = re.sub(r'[^\w\s]', '', cleaned_name).strip()
        print(f"Cleaned company name: {cleaned_name}")

        return cleaned_name
    except Exception as e:
        print(f"Error extracting company name with OpenAI: {e}")
        return None

def get_ticker_for_company(company_name):
    """
    Given a company name, uses OpenAI to return the company's stock ticker symbol.
    The prompt instructs the model to return only the ticker symbol.
    """
    prompt = f"Provide the stock ticker symbol for the following company: {company_name}. Return only the ticker symbol."
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a financial assistant with expertise in stock markets."},
                {"role": "user", "content": prompt}
            ]
        )
        ticker = response['choices'][0]['message']['content'].strip()
        print(f"Ticker for {company_name}: {ticker}")
        return ticker
    except Exception as e:
        print(f"Error getting ticker for {company_name}: {e}")
        return None

def fetch_transcript(ticker, quarter, year, api_key, retries=3):
    """
    Fetches the earnings call transcript from the API.
    Ensures the ticker is valid before making an API call.
    """
    if not ticker.isalnum() or len(ticker) < 2:  # Skip invalid tickers like R, T, X
        print(f"Invalid ticker: {ticker}")
        return None

    url = f'https://financialmodelingprep.com/api/v3/earning_call_transcript/{ticker}?quarter={quarter}&year={year}&apikey={api_key}'
    for attempt in range(retries):
        try:
            response = requests.get(url)
            response.raise_for_status()
            return response.json()
        except requests.RequestException as e:
            print(f"Error fetching data for {ticker}: {e}, attempt {attempt + 1}/{retries}")
            time.sleep(2)  # Retry after a short delay
    return None

def sentiment_analysis(text):
    """Performs sentiment analysis on the provided text using TextBlob."""
    blob = TextBlob(text)
    return blob.sentiment.polarity

def process_keywords_in_transcript(transcripts, keywords):
    """
    Searches for keywords in the transcript content and performs sentiment analysis.
    """
    if not transcripts:
        print("No transcripts fetched!")
        return []

    # Ensure keywords is a list
    if isinstance(keywords, str):
        keywords = [keywords]

    # Clean and validate keywords
    cleaned_keywords = [keyword.strip().lower() for keyword in keywords if keyword.strip()]
    print(f"Cleaned Keywords: {cleaned_keywords}")

    keyword_data = []
    for transcript in transcripts:
        if 'content' not in transcript:
            print("Content missing in the transcript!")
            continue

        content = transcript.get('content', '').lower()
        for keyword in cleaned_keywords:
            if keyword in content:
                sentences = [s.strip() for s in content.split('. ') if keyword in s]
                for sentence in sentences:
                    sentiment = sentiment_analysis(sentence)
                    keyword_data.append({
                        'Keyword': keyword,
                        'Sentence': sentence,
                        'Sentiment': sentiment
                    })

    if not keyword_data:
        print(f"No matches found for keywords: {cleaned_keywords}")
    return keyword_data

def process_query_2(query):
    """
    Processes the user's natural language query for earnings-related queries.
    It extracts the company name, obtains its ticker via OpenAI, and then extracts quarter and year information.
    """
    try:
        # Extract company name from the query using OpenAI
        company_name = extract_company_name(query)
        if not company_name:
            print("Error: Could not extract company name.")
            return None, None, None, None

        print(f"Extracted company name: {company_name}")

        # Use OpenAI to obtain the ticker symbol directly
        ticker = get_ticker_for_company(company_name)
        if not ticker:
            print(f"No valid ticker found for {company_name}.")
            return None, None, None, None

        print(f"Confirmed ticker: {ticker}")

        # Extract quarter and year from the query using a regular expression
        quarter_match = re.search(r"(Q[1-4]) (\d{4})", query, re.IGNORECASE)
        if not quarter_match:
            print("Error: Could not extract quarter or year.")
            return ticker, None, None, None

        quarter = quarter_match.group(1)[1]  # e.g., from "Q2" take "2"
        year = quarter_match.group(2)
        print(f"Extracted quarter: {quarter}, year: {year}")

        # Return ticker, quarter, year; keywords may be handled separately
        return ticker, quarter, year, None

    except Exception as e:
        print(f"Error in process_query_2: {e}")
        return None, None, None, None

def main(ticker, quarter, year, keywords):
    """
    Main function to fetch the transcript, analyze keywords, and generate an Excel report.
    """
    try:
        print(f"Fetching transcript for {ticker}...")

        if not keywords:
            print("Error: Keywords are required.")
            return "No keywords provided."

        transcripts = fetch_transcript(ticker, quarter, year, api_key)
        if not transcripts:
            print(f"No transcripts found for {ticker}.")
            return f"No data found for {ticker} for Q{quarter} {year}."

        print(f"Processing transcripts for {ticker} for Q{quarter} {year}")
        keyword_data = process_keywords_in_transcript(transcripts, keywords)

        if keyword_data:
            df = pd.DataFrame(keyword_data)
            file_path = f'Stock Reports/{ticker}_Q{quarter}_{year}_Keyword_Analysis.xlsx'
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            df.to_excel(file_path, index=False)
            print(f"Excel file created successfully: {file_path}")
            return f"Excel file created successfully: {file_path}"
        else:
            print("No relevant data found.")
            return "No relevant data found for the specified query."
    except Exception as e:
        print(f"Error in main: {e}")
        return "An unexpected error occurred."

# Example usage:
if __name__ == "__main__":
    # Example query (make sure the query includes a quarter and year)
    query = "I need the earnings call transcript for Apple Q3 2022. Can you help?"
    ticker, quarter, year, _ = process_query_2(query)
    
    if ticker and quarter and year:
        # Define the keywords you want to search for in the transcript
        keywords = ["revenue", "profit", "guidance"]
        result = main(ticker, quarter, year, keywords)
        print(result)
    else:
        print("Failed to process the query.")
