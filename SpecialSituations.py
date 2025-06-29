import os
import requests
import pdfplumber
from docx import Document
from typing import List
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

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
                text += page.extract_text() + "\n"
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

def generate_special_situation_note(company_name: str, situation_type: str, file_paths: List[str], output_path: str):
    # 1. Extract content from files
    combined_text = ""
    for path in file_paths:
        if path.endswith(".pdf"):
            combined_text += extract_text_from_pdf(path) + "\n"
        elif path.endswith(".docx"):
            combined_text += extract_text_from_docx(path) + "\n"
        else:
            combined_text += f"[Unsupported file: {path}]\n"

    # 2. Select report structure
    structure = REPORT_TEMPLATES.get(situation_type)
    if not structure:
        raise ValueError(f"Unsupported situation type: {situation_type}")

    # 3. Build prompt
    prompt = f"""
You are an institutional investment analyst writing a professional memo on a special situation involving {company_name}.
The situation is: **{situation_type}**

Below is the internal company information extracted from various files:

\"\"\"{combined_text[:7000]}\"\"\"

Using the structure below, generate a well-written investment memo. Be factual, insightful, and clear.

Structure:
{structure}
"""

    # 4. Call DeepSeek
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

    # ✅ 5. Format and save to DOCX
    format_memo_docx(memo, company_name, situation_type, output_path)



def format_memo_docx(memo_text: str, company_name: str, situation_type: str, output_path: str):
    doc = Document()

    # Title Page
    title = f"{company_name} {situation_type} Investment Memo"
    title_paragraph = doc.add_paragraph()
    title_run = title_paragraph.add_run(title)
    title_run.bold = True
    title_run.font.size = Pt(18)
    title_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    doc.add_paragraph("\n")

    # Split sections
    sections = memo_text.strip().split("\n\n")

    section_number = 1
    for block in sections:
        if not block.strip():
            continue

        lines = block.strip().split("\n", 1)
        if len(lines) == 2 and lines[0].lower().startswith("investment memo:"):
            # Section Title
            heading = f"{section_number}. {lines[0].replace('Investment Memo: ', '').strip()}"
            doc.add_paragraph(heading, style="Heading 2")
            section_number += 1

            # Body text
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(lines[1].strip())
            run.font.size = Pt(11)
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        else:
            # Fallback: no structured "Investment Memo:" present
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(block.strip())
            run.font.size = Pt(11)
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    doc.save(output_path)
    print(f"✅ Formatted memo saved to: {output_path}")

