import fitz  # PyMuPDF
import requests
import os
from docx import Document
import re
import json
from datetime import datetime
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

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
        combined_text += doc[p-1].get_text() + "\n"
    return combined_text.strip()


def generate_investment_memo(filtered_text, custom_notes=""):
    base_prompt = (
        "Using the text below extracted from a company's DRHP, generate a professional pre-IPO investment memo. "
        "The memo should include:\n"
        "1. IPO Offer Details\n"
        "2. Company Overview\n"
        "3. Industry Overview and Outlook\n"
        "4. Business Model\n"
        "5. Financial Highlights\n"
        "6. Guidance and Outlook on future financial performance\n"
        "7. Peer Comparison and Competitors\n"
        "8. Risks\n"
        "9. Investment Highlights\n"
        "Write in full sentences forming coherent paragraphs. Use ISO currency codes (INR, USD, etc.). "
        "Use proper headings and bullet points only for lists. Figures should be in millions with comma separators.\n\n"
    )
    if custom_notes:
        base_prompt += f"Special Instructions: {custom_notes}\n\n"

    base_prompt += f"DRHP Text:\n{filtered_text[:8000]}"

    messages = [
        {"role": "system", "content": "You are an expert financial analyst."},
        {"role": "user", "content": base_prompt},
    ]

    response = requests.post(
        DEEPSEEK_API_URL,
        headers={"Authorization": f"Bearer {DEEPSEEK_API_KEY}"},
        json={"model": "deepseek-chat", "messages": messages}
    )
    response.raise_for_status()
    return response.json()['choices'][0]['message']['content']


def save_memo_to_word(memo_text, company_name="Company", output_dir="documents"):
    os.makedirs(output_dir, exist_ok=True)
    filename = f"{company_name.replace(' ', '_')}_PreIPO_Memo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    full_path = os.path.join(output_dir, filename)

    doc = Document()

    # Set default font to Aptos Display (fallback to Calibri)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Aptos Display'
    font.size = Pt(11)

    # Add formatted title
    doc.add_heading(f"{company_name} Pre-IPO Investment Memo", 0)
    doc.add_paragraph()

    # Split into sections
    sections = memo_text.split('\n\n')
    for section in sections:
        section = section.strip()

        # Format headings
            heading_keywords = ['company overview', 'industry overview', 'business model',
                        'financial highlights', 'management', 'key risks', 'investment rationale', 'conclusion']

            for section in sections:
                section = section.strip().replace('#', '').replace('*', '').strip()

                if any(section.lower().startswith(h) for h in heading_keywords):
                    para = doc.add_paragraph()
                    run = para.add_run(section)
                    run.bold = True
                    run.font.size = Pt(14)

                elif re.match(r"^(\*|-|•|\d+\.)\s+", section):
                    # Handle bullets like *, -, •, or numbered list (1., 2., etc.)
                    clean_text = re.sub(r"^(\*|-|•|\d+\.)\s+", "", section)
                    doc.add_paragraph(clean_text, style='List Bullet')

                else:
                    doc.add_paragraph(section)



    doc.save(full_path)
    return full_path



# ========== FINAL RUN PIPELINE ==========
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

    memo_text = generate_investment_memo(filtered_text, custom_focus)
    return save_memo_to_word(memo_text, output_dir=output_dir)
