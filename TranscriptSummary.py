import os
import json
import openai
import requests
from bs4 import BeautifulSoup
from PyPDF2 import PdfReader

# Load your API keys from the environment
DEEPSEEK_API_KEY = os.environ["DEEPSEEK_API_KEY"]
FMP_API_KEY      = os.environ["FMP_API_KEY"]

# Import the new OpenAI class and create a client instance for DeepSeek
from openai import OpenAI
client = OpenAI(api_key=DEEPSEEK_API_KEY,base_url=os.getenv("DEEPSEEK_API_URL", "https://api.deepseek.com"))

def extract_text_from_pdf(file_stream):
    """Extract text from a PDF file stream."""
    reader = PdfReader(file_stream)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text

def extract_text_from_html(file_stream):
    """Extract text from an HTML file stream."""
    soup = BeautifulSoup(file_stream.read(), "html.parser")
    return soup.get_text(separator="\n")

def summarize_text(text):
    prompt = (
        "Please provide a concise summary of the following earnings call transcript:\n\n"
        f"{text}\n\nSummary:"
    )
    response = client.chat.completions.create(
        model="deepseek-chat",  # DeepSeek's model identifier
        messages=[{"role": "user", "content": prompt}],
        temperature=0.5,
        max_tokens=300
    )
    summary = response.choices[0].message.content.strip()
    return summary

def extract_company_details(user_message):
    prompt = (
        "Extract the company name, correct stock ticker, quarter (as a number), and year from the following request. "
        "If quarter or year are not mentioned, return null for them. "
        "Use correct stock tickers for major public companies. "
        "Respond ONLY in JSON format like this:\n"
        '{"company": "Company Name", "ticker": "TICKER", "quarter": x, "year": "YYYY"}\n\n'
        f"Request: {user_message}\n\nJSON Output:"
    )

    try:
        response = client.chat.completions.create(
            model="deepseek-chat",  # DeepSeek's model identifier
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            max_tokens=100
        )
        raw_response = response.choices[0].message.content.strip()
        print("🔎 Raw DeepSeek Response for company details:", raw_response)
        
        # Remove markdown formatting if present
        if raw_response.startswith("```"):
            first_line_end = raw_response.find("\n")
            if first_line_end != -1:
                raw_response = raw_response[first_line_end+1:]
            if raw_response.endswith("```"):
                raw_response = raw_response[:-3].strip()
        
        details = json.loads(raw_response)
        if isinstance(details.get("quarter"), str) and details["quarter"].startswith("Q"):
            details["quarter"] = int(details["quarter"][1])
        return details
    except Exception as e:
        print(f"🚨 Error in extract_company_details: {e}")
        return {"company": None, "ticker": None, "quarter": None, "year": None}

def get_transcript_from_fmp(ticker, quarter, year):
    """Fetch the earnings call transcript using FMP API for the given ticker, quarter, and year."""
    transcript_url = (
        f"https://financialmodelingprep.com/api/v3/earning_call_transcript/{ticker}"
        f"?quarter={quarter}&year={year}&apikey={FMP_API_KEY}"
    )
    print(f"🔎 Fetching transcript: {transcript_url}")

    r = requests.get(transcript_url)
    if r.status_code == 200:
        data = r.json()
        print("🔎 Raw transcript response:", data)

        if isinstance(data, list) and len(data) > 0:
            return data[0].get("content", "")

    print(f"🚨 No transcript found for {ticker} in {quarter} {year}")
    return None

def generate_earnings_call_summary(user_message):
    """Main function to generate an earnings call summary based on user input."""
    try:
        print("✅ Step 1: Extracting company details from message")
        details = extract_company_details(user_message)  # Uses DeepSeek
        print("🔎 Extracted details:", details)

        company = details.get("company")
        ticker = details.get("ticker")
        quarter = details.get("quarter")
        year = details.get("year")

        if not company or not ticker:
            print("🚨 Error: Missing company name or ticker")
            return "Could not determine company name or ticker from query."

        print(f"✅ Step 2: Fetching transcript for {ticker} - {quarter} {year}")
        transcript = get_transcript_from_fmp(ticker, quarter, year)  # Calls FMP API

        if not transcript:
            print(f"🚨 Error: No transcript found for {ticker} ({company}) for {quarter} {year}.")
            return f"No transcript found for {ticker} ({company}) for {quarter} {year}."

        print("✅ Step 3: Summarizing transcript")
        summary = summarize_text(transcript)  # Uses DeepSeek

        if not summary:
            print("🚨 Error: Failed to generate summary")
            return "Failed to generate summary."

        print("✅ Step 4: Summary complete!")
        return summary

    except Exception as e:
        print(f"🔥 ERROR: Internal Server Error: {str(e)}")
        import traceback
        traceback.print_exc()  # Print the full traceback
        return None
