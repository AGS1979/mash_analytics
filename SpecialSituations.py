import os
import re
from typing import List

import pdfplumber
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

# ==========================
# DeepSeek Setup
# ==========================
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")  # Ensure it's set in your .env or environment
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"

# ==========================
# Report Structures
# ==========================
REPORT_TEMPLATES = {
    "Spin-Off or Split-Up": """
Transaction Overview
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
Forced selling, low float, governance concerns
""",
    "Mergers & Acquisitions": """
Deal Summary
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
IRR scenarios based on timing/risk
""",
    "Bankruptcy / Distressed / Restructuring": """
Situation Summary
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
Judge approval, creditor objections, asset sales
""",
    "Activist Campaign": """
Activist Background
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
NPV of potential changes (e.g., spin-off value, ROIC uplift)
""",
    "Regulatory or Legal Catalyst": """
Legal/Regulatory Background
Case/issue summary
Historical legal proceedings
Outcome Scenarios
Win, loss, settlement
Timeline
Financial and Strategic Implications
Fines, product approval, license loss
Revenue/EBITDA impact
Market Reaction History (if any)
Past similar cases
""",
    "Asset Sales or Carve-Outs": """
Transaction Overview
Buyer, price, structure
Valuation vs. book and peers
Strategic Impact
Focus shift, deleveraging, margin profile
Use of Proceeds
Debt repayment, dividends, buybacks, capex
Re-rating Potential
EBITDA margin uplift, return metrics
""",
    "Capital Raising or Buyback Catalyst": """
Transaction Mechanics
Size, dilution, instrument type
Capital Structure Post-Deal
Leverage ratios, interest burden
Shareholder Implications
Accretion/dilution
EPS impact
Buyback Analysis (if applicable)
Repurchase pace, valuation support
"""
}

# ==========================
# Text Extractors
# ==========================

def extract_text_from_pdf(file_path: str) -> str:
    try:
        text = ""
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
        return text.strip()
    except Exception as e:
        return f"[ERROR extracting PDF: {e}]"

def extract_text_from_docx(file_path: str) -> str:
    try:
        doc = Document(file_path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except Exception as e:
        return f"[ERROR extracting DOCX: {e}]"

# ==========================
# Main Orchestration Function
# ==========================

def clean_markdown(text):
    text = re.sub(r'#+\s*', '', text)
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'\*(.*?)\*', r'\1', text)
    text = re.sub(r'`{1,3}(.*?)`{1,3}', r'\1', text)
    text = re.sub(r'!\[.*?\]\(.*?\)', '', text)
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = re.sub(r'^- ', '• ', text, flags=re.MULTILINE)
    return text.strip()

def truncate_safely(text, limit=7000):
    if len(text) <= limit:
        return text
    cutoff = text[:limit].rfind('\n\n')
    return text[:cutoff] if cutoff != -1 else text[:limit]

def generate_special_situation_note(company_name: str, situation_type: str, file_paths: List[str], output_path: str):
    combined_text = ""
    for path in file_paths:
        if path.endswith(".pdf"):
            combined_text += extract_text_from_pdf(path) + "\n"
        elif path.endswith(".docx"):
            combined_text += extract_text_from_docx(path) + "\n"
        else:
            combined_text += f"[Unsupported file: {path}]\n"

    structure = REPORT_TEMPLATES.get(situation_type)
    if not structure:
        raise ValueError(f"Unsupported situation type: {situation_type}")

    prompt = f"""
You are an institutional investment analyst writing a professional memo on a special situation involving {company_name}.
The situation is: **{situation_type}**

Below is the internal company information extracted from various files:

\"\"\"{truncate_safely(combined_text)}\"\"\"

Using the structure below, generate a well-written investment memo. Be factual, insightful, and clear.

Structure:
{structure}
"""

    headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}"}
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3
    }

    print("⏳ Generating memo using DeepSeek...")
    response = requests.post(DEEPSEEK_URL, headers=headers, json=payload)
    response.raise_for_status()
    memo = response.json()["choices"][0]["message"]["content"]

    memo = clean_markdown(memo)
    # Pass the template ('structure') to the splitting function for reliable parsing.
    memo_dict = split_into_sections(memo, structure)
    format_memo_docx(memo_dict, company_name, situation_type, output_path)


# ==========================
# Word Formatting and Section Splitting
# ==========================

def format_memo_docx(memo_dict, company_name, situation_type, output_path):
    doc = Document()

    # Set default style
    style = doc.styles['Normal']
    style.font.name = 'Aptos Display'
    style.font.size = Pt(11)

    # Title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(f"{company_name} – {situation_type} Investment Memo")
    title_run.font.name = 'Aptos Display'
    title_run.font.size = Pt(20)
    title_run.bold = True

    doc.add_paragraph()

    # Fallback if memo is not a dictionary
    if not isinstance(memo_dict, dict) or not memo_dict:
        memo_dict = {"Memo": str(memo_dict)}

    # Apply formatting per section
    for section_title, content in memo_dict.items():
        heading = doc.add_paragraph()
        heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = heading.add_run(section_title)
        run.bold = True
        run.font.size = Pt(14)
        run.font.name = 'Aptos Display'

        # Process paragraphs within the section content
        for para in content.strip().split('\n\n'):
            if para.strip():
                p = doc.add_paragraph(para.strip())
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_after = Pt(10)
                p.paragraph_format.line_spacing = 1.5

        doc.add_paragraph()

    # Set page margins
    section = doc.sections[0]
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)

    doc.save(output_path)

def split_into_sections(text: str, template: str):
    """
    Splits the memo text into a dictionary of sections based on titles from a template.

    This corrected function reliably extracts section titles from the template and uses
    them as delimiters to split the memo text, avoiding the fragility of the previous
    generic pattern matching.

    Args:
        text: The memo text to be split.
        template: The report template string containing the section titles.

    Returns:
        A dictionary with section titles as keys and their content as values.
    """
    sections = {}
    # Extract canonical titles from the template, taking only text before any '('.
    titles = [line.split('(')[0].strip() for line in template.strip().split('\n') if line.strip()]
    if not titles:
        return {"Memo": text.strip()} if text.strip() else {}

    # Build a regex to find any of the titles when they appear on their own line.
    pattern = re.compile(r'^(' + '|'.join(map(re.escape, titles)) + r')\s*$', re.MULTILINE | re.IGNORECASE)

    matches = list(pattern.finditer(text))
    if not matches:
        return {"Memo": text.strip()} if text.strip() else {}

    # Iterate through the found titles to carve out the sections
    for i, match in enumerate(matches):
        title = match.group(1).strip()
        start_of_content = match.end()
        end_of_content = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        content = text[start_of_content:end_of_content].strip()
        
        # Use the canonical title from the template for consistent key names
        canonical_title = next((t for t in titles if t.lower() == title.lower()), title)
        sections[canonical_title] = content

    return sections