import os
import requests
from bs4 import BeautifulSoup
from openai import OpenAI
import re
import json

# Load API keys
FMP_API_KEY = os.environ.get("FMP_API_KEY")
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY")
DEEPSEEK_API_BASE = os.getenv("DEEPSEEK_API_BASE", "https://api.deepseek.com")

# DeepSeek client
client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_API_BASE)

# SEC-compliant headers
SEC_HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; MaShBot/1.0; https://mash-analytics.onrender.com/chat)",
    "Accept-Encoding": "gzip, deflate",
    "Host": "www.sec.gov",
    "Connection": "keep-alive"
}



def extract_ticker_from_query(query):
    system_msg = "You are a financial assistant. Given a user query, return ONLY the stock ticker symbol for the company mentioned. If no valid ticker can be determined, return 'UNKNOWN'."
    user_msg = f"Query: {query}"

    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": system_msg},
            {"role": "user", "content": user_msg}
        ],
        temperature=0
    )

    raw_output = response.choices[0].message.content.strip()
    print("🧠 DeepSeek raw output:", raw_output)

    if raw_output.upper() == "UNKNOWN":
        return None
    return raw_output.upper()

def get_latest_10k_final_link(ticker):
    ticker = ticker.upper()
    url = f"https://financialmodelingprep.com/api/v3/sec_filings/{ticker}?type=10-k&page=0&apikey={FMP_API_KEY}"
    response = requests.get(url)
    print(f"🔗 URL requested: {url}")
    
    if response.status_code != 200:
        print("❌ Request failed:", response.status_code)
        return None

    data = response.json()
    print(f"📄 Number of filings received: {len(data)}")

    for item in data:
        print(f"➡️ Checking item: {item.get('type')} | FinalLink: {item.get('finalLink')}")
        if item.get("type", "").lower() == "10-k" and item.get("finalLink"):
            print(f"✅ FinalLink found: {item['finalLink']}")
            return item["finalLink"]

    print("❌ No valid 10-K with finalLink found.")
    return None


def extract_debt_related_text(html):
    soup = BeautifulSoup(html, "html.parser")
    pattern = re.compile(r"(note\s+\d+\s*[-–—]?\s*)?(debt|borrowings|credit|covenant|loan)", re.IGNORECASE)
    candidates = soup.find_all(['b', 'strong', 'div', 'span', 'p'])

    extracted = []
    for tag in candidates:
        header = tag.get_text(strip=True)
        if pattern.search(header):
            content_block = []
            next_node = tag.find_next_sibling()
            while next_node and next_node.name in ['p', 'div', 'span', 'table']:
                content_block.append(next_node.get_text(strip=True))
                next_node = next_node.find_next_sibling()
            if content_block:
                extracted.append(f"<h4>{header}</h4><p>{'</p><p>'.join(content_block)}</p>")

    return "<br><br>".join(extracted) if extracted else None


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
    try:
        # Extract only the JSON array using regex
        match = re.search(r"\[\s*{.*?}\s*\]", covenants, re.DOTALL)
        if not match:
            raise ValueError("No JSON array found in the response.")
        rows = json.loads(match.group(0))

        # Wrap table in a scrollable div
        html = "<div style='overflow-x:auto;'>"
        html += "<table border='1' style='width:100%; table-layout:fixed;'>"
        html += "<tr><th>Type</th><th>Description</th><th>Condition</th></tr>"
        for row in rows:
            html += (
                "<tr>"
                f"<td>{row.get('Type')}</td>"
                f"<td>{row.get('Description')}</td>"
                f"<td>{row.get('Condition', '')}</td>"
                "</tr>"
            )
        html += "</table></div>"
        return html
    except Exception as e:
        return f"<p>Error parsing covenant data: {e}</p><pre>{covenants}</pre>"

def analyze_debt_covenants(ticker):
    link = get_latest_10k_final_link(ticker)
    if not link:
        return "<p>10-K filing not found for ticker.</p>"

    print(f"✅ Fetching HTML from: {link}")
    resp = requests.get(link, headers=SEC_HEADERS)
    content_type = resp.headers.get("Content-Type", "")
    print(f"🔎 Content-Type: {content_type}")

    if "html" not in content_type.lower():
        return "<p>10-K filing is not in HTML format. Please try another company or use a different year.</p>"

    html = resp.text
    print("🧪 Preview of HTML:", html[:500])

    debt_section = extract_debt_related_text(html)
    if not debt_section:
        return "<p>No debt covenant-related text found.</p>"

    covenants_raw = extract_covenants_with_deepseek(debt_section)
    html_output = format_as_html_table(covenants_raw)
    return html_output
