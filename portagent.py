import os
import faiss
from typing import List
from sentence_transformers import SentenceTransformer
from PyPDF2 import PdfReader
from docx import Document

# === SETUP ===
embedding_model = SentenceTransformer("all-MiniLM-L6-v2")  # 384-dim
company_indexes = {}     # {company: FAISS Index}
company_chunks = {}      # {company: list of text chunks}
company_metadata = {}    # {company: list of metadata entries}

# === FUNCTIONS ===
def extract_text(file_path: str) -> str:
    if file_path.lower().endswith(".pdf"):
        reader = PdfReader(file_path)
        return "\n".join(page.extract_text() or "" for page in reader.pages)

    elif file_path.lower().endswith(".docx"):
        doc = Document(file_path)
        return "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())

    elif file_path.lower().endswith(".txt"):
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()

    else:
        return ""

def chunk_text(text: str, max_tokens: int = 200) -> List[str]:
    paragraphs = text.split("\n")
    chunks, current = [], ""
    for para in paragraphs:
        para = para.strip()
        if not para:
            continue
        if len(current.split()) + len(para.split()) <= max_tokens:
            current += " " + para
        else:
            if current:
                chunks.append(current.strip())
            current = para
    if current:
        chunks.append(current.strip())
    return chunks

def index_pdf(file_path: str, company: str):
    text = extract_text(file_path)
    if not text.strip():
        return
    chunks = chunk_text(text)
    vectors = embedding_model.encode(chunks)

    if company not in company_indexes:
        company_indexes[company] = faiss.IndexFlatL2(384)
        company_chunks[company] = []
        company_metadata[company] = []

    company_indexes[company].add(vectors)
    company_chunks[company].extend(chunks)
    company_metadata[company].extend([file_path] * len(chunks))

def query_gpt_prompt(query: str, selected_companies: List[str], k: int = 5) -> str:
    vec = embedding_model.encode([query])
    prompt = "Use only these excerpts to answer:\n\n"

    for company in selected_companies:
        if company not in company_indexes:
            continue
        D, I = company_indexes[company].search(vec, k)
        for i, idx in enumerate(I[0]):
            chunk = company_chunks[company][idx]
            prompt += f"[{company}-{i+1}] {chunk}\n\n"

    prompt += f"---\nQ: {query}\nA:"
    return prompt
