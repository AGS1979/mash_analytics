import os
import faiss
from typing import List
from sentence_transformers import SentenceTransformer
from PyPDF2 import PdfReader
import os
import openai
from dotenv import load_dotenv
load_dotenv()

# Set OpenAI API Key once
openai.api_key = os.getenv("OPENAI_API_KEY")  # Load from environment or .env

# === SETUP ===
embedding_model = SentenceTransformer("all-MiniLM-L6-v2")  # 384-dim
index = faiss.IndexFlatL2(384)
documents = []
metadata = []

# === FUNCTIONS ===
def extract_text_from_pdf(file_path: str) -> str:
    reader = PdfReader(file_path)
    return "\n".join(page.extract_text() or "" for page in reader.pages)

def chunk_text(text: str, max_tokens: int = 200) -> List[str]:
    paragraphs = text.split("\n")
    chunks, current = [], ""
    for para in paragraphs:
        if len(current.split()) + len(para.split()) <= max_tokens:
            current += " " + para.strip()
        else:
            chunks.append(current.strip())
            current = para.strip()
    if current:
        chunks.append(current.strip())
    return chunks

def index_pdf(file_path: str, doc_id: str):
    global documents, metadata, index
    text = extract_text_from_pdf(file_path)
    chunks = chunk_text(text)
    vectors = embedding_model.encode(chunks)
    index.add(vectors)
    documents.extend(chunks)
    metadata.extend([doc_id] * len(chunks))

def query_gpt_prompt(query: str, k: int = 5) -> str:
    vec = embedding_model.encode([query])
    D, I = index.search(vec, k)
    retrieved = [documents[i] for i in I[0]]
    prompt = "Answer using only these excerpts:\n\n"
    for i, chunk in enumerate(retrieved, 1):
        prompt += f"[{i}] {chunk}\n\n"
    prompt += f"---\n\nQ: {query}\nA:"
    return prompt
