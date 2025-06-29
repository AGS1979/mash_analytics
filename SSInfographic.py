import os
import random
import requests
from docx import Document

# === Report Templates TOC ===
REPORT_TEMPLATES = {
    "Spin-Off or Split-Up": """Transaction Overview
ParentCo and SpinCo details
Rationale (regulatory, strategic unlock, valuation arbitrage)
Distribution terms (ratio, eligibility, tax treatment)
ParentCo Post-Spin Outlook
Strategic focus
Financial profile and valuation
SpinCo Investment Case
Business model, growth drivers
Historical and pro forma financials
Independent valuation (e.g., Sum-of-the-Parts)
Risks and Overhangs
Forced selling, low float, governance concerns""",

    "Mergers & Acquisitions": """Deal Summary
Parties involved, consideration (cash/stock), premium
Regulatory/antitrust/board approval status
Target Company Analysis
Valuation vs. offer
Control premium vs. peers
Buyer’s Rationale and Financing
Strategic fit
Synergies and pro forma financials
Deal financing (debt, equity)
Shareholder Vote & Antitrust Risk
Key holders' stance
Timing and likelihood of deal closure
Spread Analysis and Arbitrage Opportunity
Deal spread
IRR scenarios based on timing/risk""",

    "Bankruptcy / Distressed / Restructuring": """Situation Summary
Cause of distress
Filing date, jurisdiction, DIP terms
Capital Structure Analysis
Pre- and post-reorg structure
Seniority waterfall
Creditor classes and recovery potential
Valuation and Recovery Scenarios
Estimated Enterprise Value
Recovery per instrument (bonds, equity, unsecured)
Reorganization Plan and Exit Timeline
Conversion to equity, rights offering, warrants
Exit multiples
Catalysts and Legal Risks
Judge approval, creditor objections, asset sales""",

    "Activist Campaign": """Activist Background
Fund profile, history, prior campaigns
Campaign Details
Demands (board seat, spin, buyback, etc.)
Timeline of engagement
Company's Response and Governance Profile
Management alignment, shareholder defense
Scenario Analysis
Status quo vs. activist success
Proxy fight implications
Valuation Impact
NPV of potential changes (e.g., spin-off value, ROIC uplift)""",

    "Regulatory or Legal Catalyst": """Legal/Regulatory Background
Case/issue summary
Historical legal proceedings
Outcome Scenarios
Win, loss, settlement
Timeline
Financial and Strategic Implications
Fines, product approval, license loss
Revenue/EBITDA impact
Market Reaction History (if any)
Past similar cases""",

    "Asset Sales or Carve-Outs": """Transaction Overview
Buyer, price, structure
Valuation vs. book and peers
Strategic Impact
Focus shift, deleveraging, margin profile
Use of Proceeds
Debt repayment, dividends, buybacks, capex
Re-rating Potential
EBITDA margin uplift, return metrics""",

    "Capital Raising or Buyback Catalyst": """Transaction Mechanics
Size, dilution, instrument type
Capital Structure Post-Deal
Leverage ratios, interest burden
Shareholder Implications
Accretion/dilution
EPS impact
Buyback Analysis (if applicable)
Repurchase pace, valuation support"""
}

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"

FALLBACK_META = [
    ("💼", "blue"),
    ("🏢", "sky"),
    ("🌐", "indigo"),
    ("🧩", "purple"),
    ("📊", "green"),
    ("📈", "emerald"),
    ("👥", "yellow"),
    ("⚠️", "red"),
    ("💡", "pink"),
    ("🧠", "gray"),
]

def extract_sections_from_docx(docx_path):
    doc = Document(docx_path)
    sections = {}
    current_heading = None
    current_text = []

    for para in doc.paragraphs:
        text = para.text.strip()
        style_name = para.style.name.lower()

        if "heading" in style_name:
            if current_heading and current_text:
                sections[current_heading] = "\n".join(current_text).strip()
            current_heading = text
            current_text = []
        else:
            current_text.append(text)

    if current_heading and current_text:
        sections[current_heading] = "\n".join(current_text).strip()

    return sections

def summarize_section_with_deepseek(section_title, section_text):
    prompt = f"""
You are an institutional research analyst preparing a financial infographic.

Summarize the section titled \"{section_title}\" into 3 to 5 concise bullet points.
Each point should be a single sentence, highlighting key insights clearly and professionally.

Section:
\"\"\"{section_text}\"\"\"
"""
    headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}"}
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3
    }
    response = requests.post(DEEPSEEK_URL, headers=headers, json=payload)
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"].strip()

def assign_icon_and_color(index):
    return FALLBACK_META[index % len(FALLBACK_META)]

def build_infographic_html(company_name, sections):
    html = f"""
<!DOCTYPE html>
<html lang=\"en\">
<head>
  <meta charset=\"UTF-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\"/>
  <title>{company_name} – Infographic</title>
  <script src=\"https://cdn.tailwindcss.com\"></script>
  <style>
    body {{
      font-family: 'Inter', sans-serif;
      background-color: #f9fafb;
      color: #1f2937;
    }}
    .section-icon {{
      font-size: 1.4rem;
      margin-right: 0.6rem;
    }}
  </style>
</head>
<body class=\"px-6 py-8 max-w-7xl mx-auto\">
  <h1 class=\"text-3xl font-bold text-center text-blue-800 mb-2\">{company_name} – Investment Memo Infographic</h1>
  <p class=\"text-center text-sm text-gray-500 mb-10\">Generated by Aranca GenAI Platform</p>
  <div class=\"grid grid-cols-1 md:grid-cols-2 gap-6\">
"""
    for idx, (title, section_text) in enumerate(sections.items()):
        icon, color = assign_icon_and_color(idx)
        print(f"\ud83d\udd0d Summarizing section: {title}")
        summary = summarize_section_with_deepseek(title, section_text)

        lines = [line.strip("\u2022- ").strip() for line in summary.split("\n") if line.strip()]
        bullet_items = "\n".join(f"<li>{line}</li>" for line in lines)

        html += f"""
    <div class=\"shadow-md rounded-xl p-5 transition-transform hover:scale-[1.02] duration-200 border-l-4 border-{color}-600 bg-{color}-50\">
      <h2 class=\"text-lg font-semibold text-gray-800 mb-3 flex items-center\">
        <span class=\"section-icon\">{icon}</span>{title}
      </h2>
      <ul class=\"list-disc text-sm text-gray-700 space-y-1 pl-5 leading-relaxed\">
        {bullet_items}
      </ul>
    </div>
"""

    html += """
  </div>
</body>
</html>
"""
    return html

def generate_infographic_html(docx_path, company_name, situation_type, output_path):
    print("\ud83d\udcc4 Extracting memo sections...")
    raw_sections = extract_sections_from_docx(docx_path)

    toc = REPORT_TEMPLATES.get(situation_type)
    if not toc:
        print(f"\u26a0\ufe0f No TOC found for situation type: {situation_type}")
        return

    ordered_titles = [t.strip() for t in toc.strip().splitlines() if t.strip()]
    print(f"\ud83e\udded Using TOC with {len(ordered_titles)} expected sections")

    structured_sections = {}
    for expected_title in ordered_titles:
        matched = None
        for actual_title in raw_sections:
            if expected_title.lower() in actual_title.lower() or actual_title.lower() in expected_title.lower():
                matched = raw_sections[actual_title]
                break
        if matched:
            structured_sections[expected_title] = matched
        else:
            print(f"\u26a0\ufe0f Missing section in DOCX: '{expected_title}' — skipping")

    if not structured_sections:
        print("\u26a0\ufe0f No valid sections found. Check headings or memo formatting.")
        return

    print("\ud83c\udfa8 Generating HTML infographic...")
    html = build_infographic_html(company_name, structured_sections)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\u2705 Infographic saved to: {output_path}")