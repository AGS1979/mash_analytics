import os
import requests
from bs4 import BeautifulSoup
from openai import OpenAI
from flask import jsonify
import re

# Load keys from environment (your structure)
FMP_API_KEY = os.environ.get("FMP_API_KEY")
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY")
DEEPSEEK_API_BASE = os.getenv("DEEPSEEK_API_BASE", "https://api.deepseek.com")

client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_API_BASE)


def get_latest_10k_final_link(ticker):
    url = f"https://financialmodelingprep.com/api/v3/sec_filings/{ticker}?type=10-k&page=0&apikey={FMP_API_KEY}"
    response = requests.get(url)
    if response.status_code != 200:
        return None
    data = response.json()
    for item in data:
        if "finalLink" in item:
            return item["finalLink"]
    return None


def extract_debt_related_text(html):
    soup = BeautifulSoup(html, "html.parser")

    # Convert to plain text fallback if HTML structure fails
    fallback_text = soup.get_text(separator="\n")

    # Keywords to search in section headers
    keywords = [
        "debt", "long-term debt", "credit facility", "covenant",
        "loan agreement", "senior notes", "borrowings", "indenture"
    ]

    # Compile a case-insensitive regex for note headers
    pattern = re.compile(r"(note\s+\d+\s*[-–—]?\s*)?(%s)" % "|".join(keywords), re.IGNORECASE)

    extracted = []

    # Look for bold or strong tags which often contain note headers
    candidates = soup.find_all(['b', 'strong'])

    for tag in candidates:
        header_text = tag.get_text(strip=True)
        if pattern.search(header_text):
            # Try to find the relevant following content
            content_block = []
            next_node = tag.find_next_sibling()
            while next_node and (next_node.name in ['p', 'div', 'span', 'table']):
                content_block.append(next_node.get_text(strip=True))
                next_node = next_node.find_next_sibling()

            if content_block:
                extracted.append(f"<h4>{header_text}</h4><p>{'</p><p>'.join(content_block)}</p>")

    # If no structured HTML block found, fallback to text extraction
    if not extracted:
        return None  # Let upstream handler return "no covenants found"

    return "<br><br>".join(extracted)


def extract_covenants_with_deepseek(debt_text):
    system_msg = (
        "You are a financial assistant. Extract all debt covenant clauses from the input text. "
        "Categorize each clause as either: Financial Covenant, Negative Covenant, or Affirmative Covenant. "
        "Return result as a list of dictionaries with keys: 'Type', 'Description', 'Condition' (if applicable)."
    )
    user_msg = f"Text: {debt_text}"

    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": system_msg},
            {"role": "user", "content": user_msg}
        ],
        temperature=0
    )
    return response.choices[0].message.content


def format_as_html_table(covenants):
    import json
    try:
        rows = json.loads(covenants)
        html = "<table border='1'><tr><th>Type</th><th>Description</th><th>Condition</th></tr>"
        for row in rows:
            html += f"<tr><td>{row.get('Type')}</td><td>{row.get('Description')}</td><td>{row.get('Condition', '')}</td></tr>"
        html += "</table>"
        return html
    except Exception as e:
        return f"<p>Error parsing covenant data: {e}</p><pre>{covenants}</pre>"


def analyze_debt_covenants(ticker):
    link = get_latest_10k_final_link(ticker)
    if not link:
        return "<p>10-K filing not found for ticker.</p>"

    html = requests.get(link).text
    debt_section = extract_debt_related_text(html)
    if not debt_section:
        return "<p>No debt covenant-related text found.</p>"

    covenants_raw = extract_covenants_with_deepseek(debt_section)
    html_output = format_as_html_table(covenants_raw)
    return html_output
