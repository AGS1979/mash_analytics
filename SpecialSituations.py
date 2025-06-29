import os
import requests
import pdfplumber
from docx import Document
from typing import List
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re

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
# Main Exported Function
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
    format_memo_docx(memo, company_name, situation_type, output_path)

# ==========================
# Word Formatting Function
# ==========================

def format_memo_docx(memo_text: str, company_name: str, situation_type: str, output_path: str):
    from docx.enum.style import WD_STYLE_TYPE

    doc = Document()

    # === Fonts ===
    style = doc.styles['Normal']
    style.font.name = 'Aptos Display'

    if 'Heading 2' not in doc.styles:
        heading_style = doc.styles.add_style('Heading 2', WD_STYLE_TYPE.PARAGRAPH)
        heading_style.font.size = Pt(14)
        heading_style.font.bold = True
        heading_style.font.name = 'Aptos Display'
    else:
        doc.styles['Heading 2'].font.name = 'Aptos Display'

    # === Title ===
    title = f"{company_name} – {situation_type} Investment Memo"
    title_paragraph = doc.add_paragraph()
    title_run = title_paragraph.add_run(title)
    title_run.bold = True
    title_run.font.size = Pt(18)
    title_run.font.name = 'Aptos Display'
    title_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    doc.add_paragraph("\n")

    # === TOC Sections ===
    toc = REPORT_TEMPLATES.get(situation_type)
    if not toc:
        raise ValueError(f"No TOC found for situation type: {situation_type}")
    expected_sections = [t.strip() for t in toc.strip().splitlines() if t.strip()]

    parsed = {}
    current_title = None
    lines = memo_text.strip().split("\n")

    for line in lines:
        if any(line.strip().lower().startswith(sec.lower()) for sec in expected_sections):
            current_title = next((sec for sec in expected_sections if line.strip().lower().startswith(sec.lower())), line.strip())
            parsed[current_title] = []
        elif current_title:
            parsed[current_title].append(line.strip())

    for i, section in enumerate(expected_sections, 1):
        body_lines = parsed.get(section, [])
        if not body_lines:
            continue

        doc.add_paragraph(f"{i}. {section}", style="Heading 2")

        for para in "\n".join(body_lines).split("\n\n"):
            if para.strip():
                p = doc.add_paragraph(para.strip())
                p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        doc.add_paragraph()

    # === Margins ===
    section = doc.sections[0]
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    doc.save(output_path)
    print(f"✅ Clean, TOC-based memo saved to: {output_path}")