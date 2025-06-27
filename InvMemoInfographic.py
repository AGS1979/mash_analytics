import os
import re
import requests
from docx import Document
from flask import render_template_string
from dotenv import load_dotenv
from collections import defaultdict

# === Load environment variable ===
load_dotenv()
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"

# === Extract text from .docx file ===
def extract_raw_text(docx_path):
    doc = Document(docx_path)
    return "\n".join(para.text.strip() for para in doc.paragraphs if para.text.strip())

# === Call DeepSeek to summarize ===
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

# === Format bold labels inside bullet points ===
def bold_labels(text):
    return re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)

# === Parse DeepSeek's markdown-style response into sections ===
def parse_deepseek_response(summary_text):
    sections = defaultdict(list)
    current_section = None

    lines = summary_text.splitlines()
    for line in lines:
        line = line.strip()

        # Markdown-style header: "### 1. IPO Offer Details"
        section_match = re.match(r"^#+\s+\*+\d+\.\s+(.*?)\*+\s*$", line)
        if section_match:
            current_section = section_match.group(1).strip()
            continue

        # Fallback: "### Business Overview"
        simple_header = re.match(r"^#+\s+(.*)", line)
        if simple_header:
            current_section = simple_header.group(1).strip()
            continue

        # Bullet point line: "- key point"
        bullet_match = re.match(r"^- (.+)", line)
        if bullet_match and current_section:
            bullet = bold_labels(bullet_match.group(1).strip())
            sections[current_section].append(bullet)

    return dict(sections)

# === Generate final infographic-ready HTML ===
def generate_infographic_html(docx_path, company_name):
    raw_text = extract_raw_text(docx_path)
    summary = call_deepseek_summary(raw_text, company_name)
    sections = parse_deepseek_response(summary)

    # Debug output
    print("\n🧾 RAW LLM RESPONSE:\n")
    print(summary)

    print("\n🔍 Final Parsed Sections for HTML:")
    for heading, bullets in sections.items():
        print(f"\n{heading}")
        for bullet in bullets:
            print(f" - {bullet}")

    with open("templates/base_infographic.html", "r", encoding="utf-8") as f:
        html_template = f.read()

    html_rendered = render_template_string(
        html_template,
        company_name=company_name,
        sections=sections
    )
    return html_rendered
