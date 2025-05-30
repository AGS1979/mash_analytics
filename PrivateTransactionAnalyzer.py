import fitz  # PyMuPDF
import os
import requests
from flask import current_app as app

def extract_text_from_pdf(filepath):
    doc = fitz.open(filepath)
    return "\n".join([page.get_text() for page in doc])

def analyze_transaction_doc(filepath, query="Summarize the key investment insights and risks"):
    text = extract_text_from_pdf(filepath)
    truncated_text = text[:4000]  # Trim for token safety

    prompt = f"""You are a private equity investment analyst.
A potential deal document is provided below.
Respond to the query: {query}

--- DOCUMENT START ---
{truncated_text}
--- DOCUMENT END ---
"""

    headers = {"Authorization": f"Bearer {os.getenv('OPENAI_API_KEY')}"}
    payload = {
        "model": "gpt-4",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3
    }

    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    result = response.json()
    return result["choices"][0]["message"]["content"]
