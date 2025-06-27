# infographic_generator.py
from docx import Document
from flask import render_template_string
import os
from datetime import datetime

def parse_docx_sections(docx_path):
    doc = Document(docx_path)
    sections = {}
    current_heading = None
    content = []

    # Define known section headers to look for (you can customize this list)
    known_headings = [
        "Executive Summary", "Key Investment Positives", "Key Risks",
        "Valuation Summary", "Company Overview", "Market Opportunity",
        "Financial Summary", "Use of Proceeds", "Management Commentary"
    ]

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Treat paragraph as heading if it matches known titles
        if text in known_headings:
            if current_heading and content:
                sections[current_heading] = "\n".join(content).strip()
                content = []
            current_heading = text
        else:
            content.append(text)

    if current_heading and content:
        sections[current_heading] = "\n".join(content).strip()

    return sections


def generate_infographic_html(docx_path, company_name):
    sections = parse_docx_sections(docx_path)

    with open("templates/base_infographic.html", "r", encoding="utf-8") as f:
        html_template = f.read()

    html_rendered = render_template_string(
        html_template,
        company_name=company_name,
        sections=sections
    )

    return html_rendered

