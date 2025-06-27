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

    for para in doc.paragraphs:
        if para.style.name.startswith('Heading'):
            if current_heading and content:
                sections[current_heading] = "\n".join(content).strip()
                content = []
            current_heading = para.text.strip()
        elif para.text.strip():
            content.append(para.text.strip())

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

