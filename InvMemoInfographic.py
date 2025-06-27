from docx import Document
from flask import render_template_string
import re

# Recognized memo section titles
known_headings = [
    "1. IPO Offer Details", "2. Company Overview", "3. Industry Overview and Outlook",
    "4. Business Model", "5. Financial Highlights", "6. Guidance and Outlook on future financial performance",
    "7. Peer Comparison and Competitors", "8. Risks", "9. Investment Highlights"
]

def parse_docx_sections(docx_path):
    doc = Document(docx_path)
    sections = {}
    current_heading = None
    current_content = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # New section starts
        if text in known_headings:
            if current_heading and current_content:
                sections[current_heading] = format_bullets(current_content)
                current_content = []
            current_heading = text
        else:
            current_content.append((text, para.style.name))

    if current_heading and current_content:
        sections[current_heading] = format_bullets(current_content)

    return sections


def format_bullets(content_blocks):
    """
    Given a list of (text, style) tuples, clean and return bullet points
    """
    bullets = []

    for text, style in content_blocks:
        if style.startswith("List") or re.match(r'^\d+[\.\)]\s+', text) or text.startswith("•"):
            bullets.append(clean_text(text))
        else:
            # Try to split long text into smaller sentences
            split_sentences = re.split(r'[.;]\s+|\n+', text)
            for sent in split_sentences:
                if len(sent.strip()) > 30:
                    bullets.append(clean_text(sent))

    return "\n".join(bullets)


def clean_text(text):
    """
    Clean leading bullets or numbers and return trimmed string.
    """
    return re.sub(r'^(\d+[\.\)]|•|-)\s*', '', text.strip())


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
    