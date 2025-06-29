import os
import requests
from docx import Document

# === Report Templates TOC ===
# This dictionary maps a situation type to a structured list of expected section headers.
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

# --- API Configuration ---
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"

# --- Styling and Metadata ---
FALLBACK_META = [
    ("💼", "border-blue-600", "bg-blue-50"),
    ("🏢", "border-sky-600", "bg-sky-50"),
    ("🌐", "border-indigo-600", "bg-indigo-50"),
    ("🧩", "border-purple-600", "bg-purple-50"),
    ("📊", "border-green-600", "bg-green-50"),
    ("📈", "border-emerald-600", "bg-emerald-50"),
    ("👥", "border-yellow-600", "bg-yellow-50"),
    ("⚠️", "border-red-600", "bg-red-50"),
    ("💡", "border-pink-600", "bg-pink-50"),
    ("🧠", "border-gray-600", "bg-gray-50"),
]

def extract_sections_from_docx(docx_path, situation_type):
    """
    Parses a .docx file by looking for paragraphs that match the expected section titles
    for the given situation type. This is more robust than checking for "Heading" styles.
    """
    toc = REPORT_TEMPLATES.get(situation_type)
    if not toc:
        return {}
    
    # Create a set of lowercase expected titles for fast lookups
    expected_titles = {t.strip().lower() for t in toc.strip().splitlines() if t.strip()}

    doc = Document(docx_path)
    sections = {}
    current_heading = None
    current_text = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Check if the paragraph text is one of our expected headings
        if text.lower() in expected_titles:
            # If we have a pending heading and text, save it
            if current_heading and current_text:
                # Use the original heading with its original casing as the key
                sections[current_heading] = "\n".join(current_text).strip()
            
            # Start a new section, using the text from the doc as the heading
            current_heading = text
            current_text = []
        elif current_heading:
            # Only append text if we are 'inside' a section
            current_text.append(text)

    # Save the last section after the loop finishes
    if current_heading and current_text:
        sections[current_heading] = "\n".join(current_text).strip()

    return sections

def summarize_section_with_deepseek(section_title, section_text):
    """Summarizes a text section using the DeepSeek API."""
    if not DEEPSEEK_API_KEY:
        raise ValueError("DEEPSEEK_API_KEY environment variable is not set on the server.")

    prompt = f"""
You are an institutional research analyst preparing a financial infographic.
Summarize the section titled \"{section_title}\" into 3 to 5 concise bullet points.
Each point should be a single sentence, highlighting key insights clearly and professionally.
Section:
\"\"\"{section_text}\"\"\"
"""
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3
    }
    
    response = requests.post(DEEPSEEK_URL, headers=headers, json=payload)
    response.raise_for_status()
    
    return response.json()["choices"][0]["message"]["content"].strip()

def assign_meta(index):
    """Assigns an icon and color classes for a section based on its index."""
    return FALLBACK_META[index % len(FALLBACK_META)]

def build_infographic_html(company_name, sections):
    """Constructs the full HTML for the infographic."""
    html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>{company_name} – Infographic</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
  <style>
    body {{ font-family: 'Inter', sans-serif; background-color: #f9fafb; color: #1f2937; }}
    .section-icon {{ font-size: 1.4rem; margin-right: 0.6rem; }}
  </style>
</head>
<body class="px-4 py-8 md:px-6 md:py-10 max-w-7xl mx-auto">
  <header class="text-center mb-12">
    <h1 class="text-3xl md:text-4xl font-bold text-gray-800 mb-2">{company_name} – Investment Memo Infographic</h1>
    <p class="text-sm text-gray-500">Generated by Aranca GenAI Platform</p>
  </header>
  <main class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
"""
    for idx, (title, section_text) in enumerate(sections.items()):
        icon, border_class, bg_class = assign_meta(idx)
        print(f"Summarizing section: {title}")
        
        try:
            summary = summarize_section_with_deepseek(title, section_text)
            lines = [line.lstrip("•*- ").strip() for line in summary.split("\n") if line.strip()]
            bullet_items = "\n".join(f"      <li>{line}</li>" for line in lines)
        except Exception as e:
            print(f"Could not summarize section '{title}': {e}")
            bullet_items = f"<li>Error generating summary: {e}</li>"

        html += f"""
    <div class="shadow-lg rounded-xl p-5 transition-transform hover:scale-[1.02] duration-300 ease-in-out border-l-4 {border_class} {bg_class}">
      <h2 class="text-lg font-semibold text-gray-800 mb-3 flex items-center">
        <span class="section-icon">{icon}</span>{title}
      </h2>
      <ul class="list-disc text-sm text-gray-700 space-y-2 pl-5 leading-relaxed">
{bullet_items}
      </ul>
    </div>
"""

    html += """
  </main>
  <footer class="text-center mt-12">
    <p class="text-xs text-gray-400">This document is for informational purposes only. Not investment advice.</p>
  </footer>
</body>
</html>
"""
    return html

def generate_infographic_entrypoint(docx_path, company_name, situation_type, output_path):
    """Main entrypoint function to orchestrate the infographic generation process."""
    print("--- Starting Infographic Generation ---")
    print("1. Extracting memo sections from DOCX...")
    # Pass situation_type to the new robust extraction function
    raw_sections = extract_sections_from_docx(docx_path, situation_type)

    toc = REPORT_TEMPLATES.get(situation_type)
    if not toc:
        raise ValueError(f"No TOC template found for situation type: '{situation_type}'")

    ordered_titles = [t.strip() for t in toc.strip().splitlines() if t.strip()]
    print(f"2. Using TOC '{situation_type}' with {len(ordered_titles)} expected sections.")

    structured_sections = {}
    # Match the sections we found to the official template titles to ensure order and consistency
    for expected_title in ordered_titles:
        matched_content = None
        for actual_title, actual_content in raw_sections.items():
            # Use a flexible check to match titles
            if expected_title.lower() in actual_title.lower():
                matched_content = actual_content
                break
        
        if matched_content:
            # Use the official template title for the infographic display
            structured_sections[expected_title] = matched_content
            print(f"  - Found section: '{expected_title}'")
        else:
            print(f"  - WARNING: Missing section in DOCX: '{expected_title}' — skipping")

    if not structured_sections:
        # Provide a more helpful error message
        raise FileNotFoundError("No valid sections could be matched from the DOCX. Please ensure the uploaded memo contains section headings that match the standard report template.")

    print("3. Generating HTML infographic...")
    html_content = build_infographic_html(company_name, structured_sections)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"4. Success! Infographic saved to: {output_path}")
    print("--- Generation Complete ---")
