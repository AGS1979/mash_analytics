import os
import re
import requests
from docx import Document
from flask import render_template_string
from dotenv import load_dotenv

load_dotenv()
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"

def extract_raw_text(docx_path):
    doc = Document(docx_path)
    return "\n".join(para.text.strip() for para in doc.paragraphs if para.text.strip())

def call_deepseek_summary(text, company_name):
    prompt = f"""
You are an investment analyst tasked with converting the following pre-IPO memo into a concise infographic-ready summary.

Summarize the key points for each of these sections:

1. IPO Offer Details
2. Company Overview
3. Industry Overview and Outlook
4. Business Model
5. Financial Highlights
6. Guidance and Outlook on future financial performance
7. Peer Comparison and Competitors
8. Risks
9. Investment Highlights

Each section should contain 3–5 bullet points maximum. Use crisp, bullet-style formatting (no paragraphs). Keep each bullet point under 30 words.

Company: {company_name}

Memo:
{text}
"""

    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "system", "content": "You are a helpful analyst."},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.3
    }

    response = requests.post(DEEPSEEK_URL, headers=headers, json=payload)
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"]

def parse_deepseek_response(response_text):
    sections = {}
    current_section = None
    content_lines = []
    current_number = 0

    for line in response_text.splitlines():
        line = line.strip()
        if not line:
            continue

        heading_match = re.match(r"^(\d+)\.\s+(.*)", line)
        if heading_match:
            if current_section and content_lines:
                numbered_title = f"{current_number}. {current_section}"
                sections[numbered_title] = content_lines
                content_lines = []
            current_number = int(heading_match.group(1))
            current_section = heading_match.group(2)
        elif line.startswith("-"):
            content_lines.append(line.lstrip("- ").strip())

    if current_section and content_lines:
        numbered_title = f"{current_number}. {current_section}"
        sections[numbered_title] = content_lines

    return sections


def generate_infographic_html(docx_path, company_name):
    raw_text = extract_raw_text(docx_path)
    summary = call_deepseek_summary(raw_text, company_name)
    sections = parse_deepseek_response(summary)

    # ✅ Print section headings to confirm parsing worked
    print("\n🔍 Parsed Section Headings:")
    for k in sections.keys():
        print(f"  - {k}")

    with open("templates/base_infographic.html", "r", encoding="utf-8") as f:
        html_template = f.read()

    html_rendered = render_template_string(
        html_template,
        company_name=company_name,
        sections=sections
    )
    return html_rendered

