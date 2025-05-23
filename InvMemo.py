import fitz  # PyMuPDF
import requests
import os
from docx import Document
import re
import json
from datetime import datetime
from docx.shared import Pt, Inches

# ========== CONFIG ==========
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
CHUNK_SIZE = 50  # Number of pages per API call


def extract_text_by_page(pdf_path):
    doc = fitz.open(pdf_path)
    return [page.get_text() for page in doc], len(doc)


def get_relevant_pages_chunked(text_by_page, user_query):
    total_pages = len(text_by_page)
    relevant_pages = set()

    for start in range(0, total_pages, CHUNK_SIZE):
        end = min(start + CHUNK_SIZE, total_pages)
        chunk_pages = text_by_page[start:end]

        prompt = "Below are texts extracted from pages of a PDF. Identify only the page numbers (starting from 1) relevant to this query:\n"
        prompt += f"Query: {user_query}\n\n"

        for i, text in enumerate(chunk_pages):
            snippet = text[:1000].replace('\n', ' ')
            prompt += f"\nPage {start + i + 1}: {snippet}\n"

        messages = [
            {"role": "system", "content": "You are an expert document analyst."},
            {"role": "user", "content": prompt},
        ]

        payload = {
            "model": "deepseek-chat",
            "messages": messages
        }

        response = requests.post(
            DEEPSEEK_API_URL,
            headers={"Authorization": f"Bearer {DEEPSEEK_API_KEY}"},
            json=payload
        )
        response.raise_for_status()

        reply = response.json()['choices'][0]['message']['content']
        matches = re.findall(r'\d+', reply)
        for m in matches:
            num = int(m)
            if 1 <= num <= total_pages:
                relevant_pages.add(num)

    return sorted(relevant_pages)


def extract_selected_pages_text(original_path, pages_to_keep):
    doc = fitz.open(original_path)
    combined_text = ""
    for p in pages_to_keep:
        combined_text += doc[p - 1].get_text() + "\n"
    return combined_text.strip()


def extract_company_name(text):
    prompt = (
        "Extract only the legal name of the company from the following IPO or DRHP text. "
        "Return only the company name, nothing else.\n\n"
        f"{text[:3000]}"
    )

    messages = [
        {"role": "system", "content": "You are an expert in IPO documents."},
        {"role": "user", "content": prompt},
    ]

    response = requests.post(
        DEEPSEEK_API_URL,
        headers={"Authorization": f"Bearer {DEEPSEEK_API_KEY}"},
        json={"model": "deepseek-chat", "messages": messages}
    )
    response.raise_for_status()
    return response.json()['choices'][0]['message']['content'].strip()


def generate_memo_sections(filtered_text, custom_notes=""):
    section_titles = [
        "1. IPO Offer Details",
        "2. Company Overview",
        "3. Industry Overview and Outlook",
        "4. Business Model",
        "5. Financial Highlights",
        "6. Guidance and Outlook on future financial performance",
        "7. Peer Comparison and Competitors",
        "8. Risks",
        "9. Investment Highlights"
    ]

    sections = {}
    for title in section_titles:
        prompt = (
            f"Write a 500+ word section for the investment memo titled '{title}'. "
            "Use a natural, analytical, and opinionated tone like an equity research analyst. "
            "Use structured paragraphs. Use ISO currency codes, and show financials in millions with commas.\n\n"
        )
        if custom_notes:
            prompt += f"Also cover: {custom_notes.strip()}\n\n"

        prompt += f"DRHP Text:\n{filtered_text[:16000]}"

        messages = [
            {"role": "system", "content": "You are an expert financial analyst."},
            {"role": "user", "content": prompt},
        ]
        response = requests.post(
            DEEPSEEK_API_URL,
            headers={"Authorization": f"Bearer {DEEPSEEK_API_KEY}"},
            json={"model": "deepseek-chat", "messages": messages}
        )
        response.raise_for_status()
        content = response.json()['choices'][0]['message']['content']
        sections[title] = content.strip()

    return sections


def save_sections_to_word(sections_dict, company_name="Company", output_dir="documents"):
    os.makedirs(output_dir, exist_ok=True)
    filename = f"{company_name.replace(' ', '_')}_PreIPO_Memo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    full_path = os.path.join(output_dir, filename)

    

    doc = Document()
    

    style = doc.styles['Normal']
    style.font.name = 'Aptos Display'
    style.font.size = Pt(11)


    title_para = doc.add_paragraph()
    title_run = title_para.add_run(f"{company_name} Pre-IPO Investment Memo")
    title_run.font.name = 'Aptos Display'
    title_run.font.size = Pt(20)
    title_run.bold = True

    doc.add_paragraph()

    for title, body in sections_dict.items():
        heading = doc.add_paragraph()
        run = heading.add_run(title)
        run.bold = True
        run.font.name = 'Aptos Display'
        run.font.size = Pt(14)

        doc.add_paragraph(body.strip())
        doc.add_paragraph()  # small space between sections

    section = doc.sections[0]
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.gutter = Inches(0)

    doc.save(full_path)
    return full_path



def run_pipeline(pdf_path, custom_focus="", output_dir="documents"):
    text_by_page, total_pages = extract_text_by_page(pdf_path)

    default_query = (
        "Extract only those pages that contain information useful for writing a pre-IPO investment memo. "
        "This includes sections on 'Management’s Discussion and Analysis of Financial Condition and Results of Operations', "
        "'Financial Highlights', 'Risk Factors', 'Business Overview', and 'Industry Overview'. Exclude all other pages."
    )

    pages_to_keep = get_relevant_pages_chunked(text_by_page, default_query)
    if not pages_to_keep:
        raise ValueError("No relevant pages found.")

    filtered_text = extract_selected_pages_text(pdf_path, pages_to_keep)
    if not filtered_text.strip():
        raise ValueError("Filtered text is empty.")


    company_name = extract_company_name(filtered_text)
    sections_dict = generate_memo_sections(filtered_text, custom_focus)
    return save_sections_to_word(sections_dict, company_name=company_name, output_dir=output_dir)