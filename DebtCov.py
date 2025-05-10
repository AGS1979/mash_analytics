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
    system_msg = (
        "You are a financial assistant. Given a user query, return ONLY the stock ticker symbol "
        "for the company mentioned. Respond with ONLY the ticker in uppercase. If no ticker can be found, respond with 'UNKNOWN'."
    )
    user_msg = f"Query: {query}"

    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": system_msg},
            {"role": "user", "content": user_msg}
        ],
        temperature=0
    )

    raw = response.choices[0].message.content.strip()
    print("🧠 DeepSeek raw output:", raw)

    # Clean and extract the first uppercase ticker (assumes all caps, 1–5 letters)
    match = re.search(r'\b[A-Z]{1,5}\b', raw)
    if match:
        return match.group(0)
    else:
        return None

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
    all_text = soup.get_text(separator="\n", strip=True)

    # Get paragraphs around keywords (5 lines before & after)
    lines = all_text.splitlines()
    matches = []
    keywords = re.compile(r"\b(covenant|debt|credit agreement|indenture|loan agreement|restrictive covenant)\b", re.IGNORECASE)

    for i, line in enumerate(lines):
        if keywords.search(line):
            context = lines[max(i - 5, 0): min(i + 6, len(lines))]
            matches.append("\n".join(context))

    return "\n\n---\n\n".join(matches) if matches else None


def extract_covenants_with_deepseek(debt_text):
    system_msg = (
        "You are a financial assistant. Extract all clauses related to debt covenants from the given SEC 10-K text. "
        "Debt covenants may appear as terms like 'limitations on indebtedness', 'coverage ratios', 'negative pledges', "
        "'maintenance of insurance', or restrictions under 'indenture agreements'. \n\n"
        "Return each clause as a dictionary with keys:\n"
        "  - 'Type' (Financial Covenant, Negative Covenant, Affirmative Covenant)\n"
        "  - 'Description' (full clause)\n"
        "  - 'Condition' (only if explicitly stated)\n\n"
        "Return a list of these dictionaries. If none are found, return an empty list []."
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
        return (
            "<p><strong>No debt covenants found.</strong> "
            "We reviewed the relevant sections of the latest 10-K filing but did not find any clauses "
            "explicitly related to debt covenants. Try a different company or check if such clauses "
            "exist in another report.</p>"
        )

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

    # ✅ Fallback: try broader regex-based scan if main extraction fails
    if not debt_section:
        print("🔁 No specific debt/covenant section found. Performing broader keyword scan.")
        debt_section = extract_covenant_keywords_fallback(html)

    # ✅ If even fallback fails, return error message
    if not debt_section:
        return (
            "<p><strong>No debt-related section found.</strong> "
            "The latest 10-K does not contain any identifiable sections referring to debt, loans, or covenants. "
            "It’s possible this company does not disclose such details prominently.</p>"
        )


    covenants_raw = extract_covenants_with_deepseek(debt_section)
    html_output = format_as_html_table(covenants_raw)
    return html_output

def extract_covenant_keywords_fallback(html):
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text(separator="\n", strip=True)
    snippets = re.findall(r".{0,250}(covenant[s]?|debt|indenture|credit agreement).{0,250}", text, re.IGNORECASE)
    return "\n\n".join(snippets) if snippets else None