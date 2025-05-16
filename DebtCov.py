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

# SEC headers
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
        messages=[{"role": "system", "content": system_msg}, {"role": "user", "content": user_msg}],
        temperature=0
    )
    raw = response.choices[0].message.content.strip()
    print("🧠 DeepSeek raw output:", raw)
    match = re.search(r'\b[A-Z]{1,5}\b', raw)
    return match.group(0) if match else None

def get_latest_10k_final_link(ticker):
    ticker = ticker.upper()
    url = f"https://financialmodelingprep.com/api/v3/sec_filings/{ticker}?type=10-k&page=0&apikey={FMP_API_KEY}"
    response = requests.get(url)
    print(f"🔗 URL requested: {url}")
    if response.status_code != 200:
        print("❌ Request failed:", response.status_code)
        return None
    data = response.json()
    for item in data:
        if item.get("type", "").lower() == "10-k" and item.get("finalLink"):
            print(f"✅ FinalLink found: {item['finalLink']}")
            return item["finalLink"]
    return None

def extract_debt_related_text(html):
    soup = BeautifulSoup(html, "html.parser")
    all_text = soup.get_text(separator="\n", strip=True)

    paragraphs = all_text.split("\n\n")  # split by paragraph
    matches = []

    # Match paragraphs with any of these terms
    keyword_pattern = re.compile(
        r"\b(covenant|restrictive covenant|credit agreement|loan agreement|revolving credit|debt agreement|compliance with all covenants|limitation on.*debt)\b",
        re.IGNORECASE
    )

    for para in paragraphs:
        if keyword_pattern.search(para):
            matches.append(para.strip())

    return "\n\n".join(matches) if matches else None

def clean_and_shorten(text, max_blocks=7):
    paragraphs = [p.strip() for p in text.split('\n\n') if len(p.strip()) > 100]
    return "\n\n".join(paragraphs[:max_blocks])

def extract_covenants_with_deepseek(debt_text):
    system_msg = (
        "You are an expert financial analyst. You are analyzing the official SEC 10-K filing of a company.\n\n"
        "Your task is to extract only the clauses that describe **debt covenants**, which are typically found in sections labeled: "
        "'Restrictive Covenants', 'Credit Agreements', 'Debt Agreements', 'Indentures', or 'Loan Agreements'.\n\n"
        "These covenants may include:\n"
        "- Leverage ratios or debt-to-equity limits\n"
        "- Interest coverage thresholds\n"
        "- Restrictions on asset transfers, liens, guarantees, or mergers\n"
        "- Requirements to maintain insurance or credit support agreements\n\n"
        "Return your output strictly as a JSON array of dictionaries. Each dictionary must have:\n"
        "  - 'Type': 'Financial Covenant', 'Negative Covenant', or 'Affirmative Covenant'\n"
        "  - 'Description': full paragraph or clause text\n"
        "  - 'Condition': numerical condition (if any, e.g. 'minimum 1.15x coverage')\n\n"
        "**Return only the JSON array. Do not include any explanation or text outside the array.**"
    )
    user_msg = debt_text
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "system", "content": system_msg}, {"role": "user", "content": user_msg}],
        temperature=0
    )
    return response.choices[0].message.content

def format_as_html_table(covenants_raw):
    try:
        match = re.search(r"\[\s*{.*?}\s*\]", covenants_raw, re.DOTALL)
        if not match:
            raise ValueError("No JSON array found.")
        rows = json.loads(match.group(0))
        if not isinstance(rows, list) or not rows:
            raise ValueError("Empty or invalid JSON array.")
        html = "<div style='overflow-x:auto;'><table border='1' style='width:100%; table-layout:fixed;'>"
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
        print("❌ Formatting error:", e)
        return (
            "<p><strong>No debt covenants found.</strong> "
            "We reviewed the relevant sections of the latest 10-K filing but did not find any clauses "
            "explicitly related to debt covenants. Try a different company or check if such clauses exist in another report.</p>"
        )

def extract_covenant_keywords_fallback(html):
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text(separator="\n", strip=True)
    snippets = re.findall(r".{0,250}(covenant[s]?|debt|indenture|credit agreement).{0,250}", text, re.IGNORECASE)
    return "\n\n".join(snippets) if snippets else None

def analyze_debt_covenants(ticker):
    link = get_latest_10k_final_link(ticker)
    if not link:
        return "<p>10-K filing not found for ticker.</p>"
    print(f"✅ Fetching HTML from: {link}")
    resp = requests.get(link, headers=SEC_HEADERS)
    content_type = resp.headers.get("Content-Type", "")
    if "html" not in content_type.lower():
        return "<p>10-K filing is not in HTML format. Please try another company or use a different year.</p>"
    html = resp.text
    debt_section = extract_debt_related_text(html)
    if not debt_section:
        print("🔁 No specific debt/covenant section found. Performing broader keyword scan.")
        debt_section = extract_covenant_keywords_fallback(html)
    if not debt_section:
        return (
            "<p><strong>No debt-related section found.</strong> "
            "The latest 10-K does not contain any identifiable sections referring to debt, loans, or covenants. "
            "It’s possible this company does not disclose such details prominently.</p>"
        )

    debt_text = clean_and_shorten(debt_section)
    print("📄 Extracted Debt Text Sent to DeepSeek:\n")
    print(debt_text[:2000])  # print first 2000 characters
    covenants_raw = extract_covenants_with_deepseek(debt_text)
    print("📤 Raw DeepSeek output:\n", covenants_raw[:1000])  # Optional preview
    return format_as_html_table(covenants_raw)
