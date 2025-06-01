# deal_analyzer.py

import os
import json
import fitz                     # PyMuPDF
import tiktoken                 # for token counting
import requests                 # for HTTP calls to DeepSeek Chat
import pandas as pd             # for Excel extraction
from docx import Document       # for .docx extraction
from pptx import Presentation   # for .pptx extraction

# ---------------------------------------------------
# 0) FLASK APP + CONFIGURATION
# ---------------------------------------------------

app = Flask(__name__)

# Allowed file extensions (used by allowed_file())
ALLOWED_EXTENSIONS = {"pdf", "docx", "pptx", "xls", "xlsx", "xlsm"}

def allowed_file(filename: str) -> bool:
    """
    Check if the file extension is one of the allowed types.
    """
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# ---------------------------------------------------
# 1) LOAD API KEYS / ENDPOINTS FROM ENVIRONMENT
# ---------------------------------------------------

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_CHAT_URL = os.getenv(
    "DEEPSEEK_CHAT_URL",
    "https://api.deepseek.com/v1/chat/completions"
)

if not DEEPSEEK_API_KEY:
    raise RuntimeError("Please set DEEPSEEK_API_KEY in your environment.")

print(f"[DEBUG] DEEPSEEK_API_KEY found: {'Yes' if DEEPSEEK_API_KEY else 'No'}")
print(f"[DEBUG] Using DEEPSEEK_CHAT_URL: {DEEPSEEK_CHAT_URL}")

# ---------------------------------------------------
# 2) UTILITIES: token counting & chunking
# ---------------------------------------------------

def num_tokens_from_string(string: str) -> int:
    """
    Returns the number of tokens in `string` using a fixed tiktoken encoding ("cl100k_base").
    We cannot rely on encoding_for_model("deepseek-chat"), since that model name is not recognized.
    """
    encoding = tiktoken.get_encoding("cl100k_base")
    token_count = len(encoding.encode(string))
    return token_count

def chunk_text(text: str, max_tokens: int = 1800) -> list[str]:
    """
    Splits `text` into substrings, each containing <= max_tokens tokens.
    Prefers splitting on double-newline (paragraph) boundaries, but
    will fall back to sentences if a paragraph alone is too large.
    """
    paragraphs = text.split("\n\n")
    chunks: list[str] = []
    current_chunk = ""
    current_tokens = 0

    for para in paragraphs:
        para = para.strip()
        if not para:
            continue
        para_tokens = num_tokens_from_string(para)

        if para_tokens > max_tokens:
            # Paragraph itself is too big. Break into sentences.
            sentences = para.split(". ")
            for sentence in sentences:
                sentence = sentence.strip()
                if not sentence:
                    continue
                sent_tokens = num_tokens_from_string(sentence)
                if current_tokens + sent_tokens + 10 > max_tokens:
                    if current_chunk:
                        chunks.append(current_chunk.strip())
                    current_chunk = sentence + ". "
                    current_tokens = sent_tokens
                else:
                    current_chunk += sentence + ". "
                    current_tokens += sent_tokens
        else:
            # Try to append this entire paragraph to current_chunk
            if current_tokens + para_tokens + 20 > max_tokens:
                if current_chunk:
                    chunks.append(current_chunk.strip())
                current_chunk = para + "\n\n"
                current_tokens = para_tokens
            else:
                current_chunk += para + "\n\n"
                current_tokens += para_tokens

    if current_chunk:
        chunks.append(current_chunk.strip())

    print(f"[DEBUG] chunk_text: split into {len(chunks)} chunks (max_tokens={max_tokens})")
    return chunks

# ---------------------------------------------------
# 3) TEXT EXTRACTION FOR MULTIPLE FORMATS
# ---------------------------------------------------

def extract_text_from_pdf(filepath: str) -> str:
    """
    Uses PyMuPDF to extract all text from a PDF file, concatenated page by page.
    """
    print(f"[DEBUG] extract_text_from_pdf: opening {filepath}")
    doc = fitz.open(filepath)
    all_text = []
    for page_number, page in enumerate(doc, start=1):
        text = page.get_text()
        all_text.append(text)
    concatenated = "\n\n".join(all_text)
    print(f"[DEBUG] extract_text_from_pdf: extracted {len(all_text)} pages")
    return concatenated

def extract_text_from_pdf_by_page(filepath: str) -> list[str]:
    """
    Uses PyMuPDF to extract text **page by page** (list of strings, index 0 = page 1).
    """
    print(f"[DEBUG] extract_text_from_pdf_by_page: opening {filepath}")
    doc = fitz.open(filepath)
    page_texts = [page.get_text() for page in doc]
    print(f"[DEBUG] extract_text_from_pdf_by_page: extracted {len(page_texts)} pages")
    return page_texts

def extract_text_from_docx(filepath: str) -> str:
    """
    Uses python-docx to read all paragraphs from a .docx file.
    """
    print(f"[DEBUG] extract_text_from_docx: opening {filepath}")
    doc = Document(filepath)
    all_paragraphs = [para.text for para in doc.paragraphs if para.text.strip() != ""]
    joined = "\n\n".join(all_paragraphs)
    print(f"[DEBUG] extract_text_from_docx: extracted {len(all_paragraphs)} paragraphs")
    return joined

def extract_text_from_pptx(filepath: str) -> str:
    """
    Uses python-pptx to extract all text from every slide in a .pptx file.
    """
    print(f"[DEBUG] extract_text_from_pptx: opening {filepath}")
    prs = Presentation(filepath)
    slide_texts = []
    for slide_idx, slide in enumerate(prs.slides, start=1):
        box_texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text = shape.text.strip()
                if text:
                    box_texts.append(text)
        if box_texts:
            slide_texts.append("\n".join(box_texts))
    concatenated = "\n\n".join(slide_texts)
    print(f"[DEBUG] extract_text_from_pptx: extracted {len(slide_texts)} slides")
    return concatenated

def extract_text_from_excel(filepath: str) -> str:
    """
    Uses pandas/openpyxl to read each sheet in an Excel file and flatten all cells into one big text blob.
    """
    print(f"[DEBUG] extract_text_from_excel: opening {filepath}")
    xlsx = pd.ExcelFile(filepath)
    sheet_texts: list[str] = []
    for sheet in xlsx.sheet_names:
        df = pd.read_excel(xlsx, sheet_name=sheet, dtype=str)
        rows = df.fillna("").apply(lambda r: " | ".join(r.values), axis=1)
        non_empty_rows = rows[rows.str.strip() != ""]
        if not non_empty_rows.empty:
            sheet_block = f"--- SHEET: {sheet} ---\n" + "\n".join(non_empty_rows.tolist())
            sheet_texts.append(sheet_block)
    joined = "\n\n".join(sheet_texts)
    print(f"[DEBUG] extract_text_from_excel: extracted {len(sheet_texts)} sheets")
    return joined

def extract_text(filepath: str) -> str:
    """
    Dispatch by file extension: .pdf, .docx, .pptx, .xls/.xlsx/.xlsm.
    Raises ValueError on unsupported extensions.
    """
    ext = os.path.splitext(filepath)[1].lower()
    print(f"[DEBUG] extract_text: dispatching on extension '{ext}'")
    if ext == ".pdf":
        return extract_text_from_pdf(filepath)
    elif ext == ".docx":
        return extract_text_from_docx(filepath)
    elif ext == ".pptx":
        return extract_text_from_pptx(filepath)
    elif ext in [".xls", ".xlsx", ".xlsm"]:
        return extract_text_from_excel(filepath)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

# ---------------------------------------------------
# 4) LLM (DeepSeek Chat) WRAPPER
# ---------------------------------------------------

def call_deepseek_chat(
    messages: list[dict[str, str]],
    model: str = "deepseek-chat",
    temperature: float = 0.3,
    max_tokens: int = 1500,
    timeout_sec: int = 60
) -> str:
    """
    Calls DeepSeek Chat (OpenAI‐compatible chat.completions). Returns the assistant’s reply string.
    """
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type":  "application/json"
    }
    payload = {
        "model":       model,
        "messages":    messages,
        "temperature": temperature,
        "max_tokens":  max_tokens
    }

    print(f"[DEBUG] call_deepseek_chat: sending request to {DEEPSEEK_CHAT_URL}")
    print(f"[DEBUG] call_deepseek_chat: model={model}, temperature={temperature}, max_tokens={max_tokens}")
    # *** For extra debugging, you could uncomment the next line to see the entire payload:
    # print(f"[DEBUG] call_deepseek_chat payload: {json.dumps(payload)[:1000]}…")

    resp = requests.post(
        DEEPSEEK_CHAT_URL,
        headers=headers,
        json=payload,
        timeout=timeout_sec
    )
    print(f"[DEBUG] call_deepseek_chat: HTTP {resp.status_code}")

    if resp.status_code != 200:
        try:
            detail = resp.json()
        except ValueError:
            detail = resp.text
        raise RuntimeError(f"DeepSeek returned HTTP {resp.status_code}: {detail}")

    data = resp.json()
    # *** For extra debugging, you could uncomment the next line to see the raw response keys:
    # print(f"[DEBUG] call_deepseek_chat response keys: {data.keys()}")
    choice_content = data["choices"][0]["message"]["content"]
    print(f"[DEBUG] call_deepseek_chat: received {len(choice_content)} characters")
    return choice_content

# ---------------------------------------------------
# 5) PAGE IDENTIFICATION FOR ANY USER QUERY
# ---------------------------------------------------

PAGE_BATCH_SIZE = 20  # ask “which pages matter?” in batches of 20 at a time

def get_relevant_pages_chunked(text_by_page: list[str], user_query: str) -> set[int]:
    """
    Loop through `text_by_page` in batches of PAGE_BATCH_SIZE, sending a short prompt:
    “Which page numbers (1-based) are relevant to this query?”
    Returns a set of 0-based page indices.
    """
    total_pages = len(text_by_page)
    relevant_pages = set()
    print(f"[DEBUG] get_relevant_pages_chunked: total_pages={total_pages}, user_query='{user_query}'")

    for start in range(0, total_pages, PAGE_BATCH_SIZE):
        end = min(start + PAGE_BATCH_SIZE, total_pages)
        chunk_pages = text_by_page[start:end]

        prompt = (
            "Below are text snippets from pages of a PDF. Identify ONLY the page\n"
            f"numbers (1-based) that are relevant to this query:\n\nQuery: {user_query}\n\n"
        )
        for i, page_text in enumerate(chunk_pages):
            snippet = page_text[:800].replace("\n", " ")
            prompt += f"Page {start + i + 1}: {snippet}\n\n"

        messages = [
            {"role": "system", "content": "You are an expert document analyst."},
            {"role": "user",   "content": prompt}
        ]

        try:
            resp_text = call_deepseek_chat(messages, model="deepseek-chat", temperature=0.0, max_tokens=200)
        except RuntimeError as e:
            print(f"[WARNING] get_relevant_pages_chunked: batch {start}-{end} failed: {e}")
            continue

        print(f"[DEBUG] get_relevant_pages_chunked: batch {start}-{end} got response length {len(resp_text)}")

        # Parse out all integers in the response; treat them as page numbers.
        for token in resp_text.replace(",", " ").split():
            if token.isdigit():
                page_num = int(token)
                if 1 <= page_num <= total_pages:
                    relevant_pages.add(page_num - 1)  # convert to 0-based

    print(f"[DEBUG] get_relevant_pages_chunked: found relevant_pages={sorted(relevant_pages)}")
    return relevant_pages

# ---------------------------------------------------
# 6) LAYERED SUMMARIZATION + ANALYSIS ROUTINES
# ---------------------------------------------------

def summarize_chunk(
    chunk_text: str,
    custom_prompt: str | None = None
) -> str:
    """
    Summarize one chunk. By default, extract “Key Insights” and “Potential Risks.”
    If custom_prompt is provided, use that verbatim (with {{TEXT}}).
    Output is plain text paragraphs (no markdown).
    """
    if custom_prompt:
        prompt = custom_prompt.replace("{{TEXT}}", chunk_text).strip()
    else:
        prompt = f"""
You are a Private Equity analyst. From the following excerpt, write a concise paragraph
that highlights both Key Insights and Potential Risks. Do NOT use Markdown or bullet points;
just write plain English. Keep each paragraph under 80 words.

--- EXCERPT START ---
{chunk_text}
--- EXCERPT END ---
""".strip()

    messages = [
        {"role": "system", "content": "You are a knowledgeable private equity investment analyst."},
        {"role": "user",   "content": prompt}
    ]
    print(f"[DEBUG] summarize_chunk: sending one chunk (length {len(chunk_text)} chars) to DeepSeek")
    response = call_deepseek_chat(messages)
    print(f"[DEBUG] summarize_chunk: received summary length {len(response)}")
    return response

def aggregate_summaries(
    summaries: list[str],
    custom_prompt: str | None = None
) -> str:
    """
    Given a list of chunk‐level summaries (plain text paragraphs), produce a final JSON
    with keys:
      - Executive_Summary
      - Key_Investment_Insights
      - Risks_and_Red_Flags
      - High_Level_Valuation_Guidance

    If custom_prompt is provided, use it verbatim (with {{SUMMARIES}}).
    Otherwise, explicitly ask for a detailed JSON object (no markdown, no backticks).
    """
    combined_text = "\n\n".join(summaries)
    print(f"[DEBUG] aggregate_summaries: combining {len(summaries)} chunk summaries (combined length {len(combined_text)} chars)")

    if custom_prompt:
        prompt = custom_prompt.replace("{{SUMMARIES}}", combined_text).strip()
    else:
        prompt = f"""
You are a top-tier private equity partner. Below are plain-text paragraph summaries of different
sections of a deal document. Please do all of the following, and return your answer AS A JSON OBJECT
ONLY (no markdown, no backticks, no code fences):

1) Write a one‐paragraph Executive_Summary (under 100 words).
2) Provide a JSON array "Key_Investment_Insights" with 5 concise insights (plain text).
3) Provide a JSON array "Risks_and_Red_Flags" with 5 concise items (plain text).
4) Provide a JSON array "High_Level_Valuation_Guidance" with 2–3 concise bullet points (plain text).

Below are the chunk summaries:

--- CHUNK SUMMARIES START ---
{combined_text}
--- CHUNK SUMMARIES END ---
""".strip()

    messages = [
        {"role": "system", "content": "You are a senior private equity partner, extremely concise and factual."},
        {"role": "user",   "content": prompt}
    ]
    response = call_deepseek_chat(messages)
    print(f"[DEBUG] aggregate_summaries: received aggregate JSON length {len(response)}")
    return response

def deep_dive_section(
    section_name: str,
    combined_text: str,
    custom_prompt: str | None = None
) -> str:
    """
    Perform a detailed deep dive for a given section (e.g. “Market Analysis”).
    If custom_prompt is provided, replace {{SECTION_NAME}} and {{TEXT}}.
    Output should be plain English paragraphs (no Markdown).
    """
    if custom_prompt:
        prompt = (custom_prompt
                  .replace("{{SECTION_NAME}}", section_name)
                  .replace("{{TEXT}}", combined_text)
                  .strip())
    else:
        prompt = f"""
You are a private equity sector specialist. Perform a deeply detailed, multi‐paragraph analysis
for the section: "{section_name}". For each paragraph, do not use bullets or markdown—just plain English
with clear headings (e.g. "Section: {section_name}") if you like. Structure your response as follows:

1) Write a short overview paragraph (2–3 sentences).
2) Then write a detailed multi‐paragraph discussion (4–5 paragraphs) touching on domain insights,
   potential red flags, and implications for deal structure or valuation—whatever is relevant.
3) Conclude with a final paragraph summarizing “Implications for the Deal”.

Below is the combined text (from all chunks). Use it to inform your deep dive.

--- DOCUMENT EXCERPT START ({section_name}) ---
{combined_text}
--- DOCUMENT EXCERPT END ---
""".strip()

    messages = [
        {"role": "system", "content": "You are a domain expert in private equity deals."},
        {"role": "user",   "content": prompt}
    ]
    print(f"[DEBUG] deep_dive_section: running deep dive on '{section_name}' (combined length {len(combined_text)} chars)")
    response = call_deepseek_chat(messages, temperature=0.25, max_tokens=1000)
    print(f"[DEBUG] deep_dive_section: received section '{section_name}' length {len(response)}")
    return response

# ---------------------------------------------------
# 7) MAIN “ANALYZE TRANSACTION” WORKFLOW
# ---------------------------------------------------

def analyze_transaction_doc(
    filepath: str,
    user_query: str,
    run_deep_dives: bool = True,
    chunk_size: int = 1800,
    deep_dive_sections: list[str] | None = None,
    chunk_prompt: str | None = None,
    aggregate_prompt: str | None = None,
    deep_dive_prompts: dict[str, str] | None = None
) -> dict:
    """
    1) Extract the PDF as a list of pages.
    2) Identify which pages matter for `user_query` (using get_relevant_pages_chunked).
    3) Concatenate only those pages (or all if none found).
    4) If the user_query mentions “red flag” or “risk,” do a single‐shot prompt for red flags.
    5) Otherwise, do layered summarization:
         a) Chunk the selected_text into <= chunk_size tokens each.
         b) If the number of chunks is <= 20, summarize each chunk & aggregate.
         c) If > 20 chunks, fall back to a single‐shot aggregated‐prompt.
    6) Optionally, run deep dives on the aggregated summaries.
    """
    print(f"[DEBUG] analyze_transaction_doc: Starting analysis for query='{user_query}' on file '{filepath}'")

    # 1) Extract PDF by page
    text_by_page = extract_text_from_pdf_by_page(filepath)

    # 2) Identify relevant pages
    relevant_pages = get_relevant_pages_chunked(text_by_page, user_query)

    # 3) Build selected_text
    if relevant_pages:
        selected_text = "\n\n".join(text_by_page[i] for i in sorted(relevant_pages))
        pages_scanned = len(text_by_page)
        pages_used = sorted(relevant_pages)
    else:
        selected_text = "\n\n".join(text_by_page)
        pages_scanned = len(text_by_page)
        pages_used = list(range(pages_scanned))

    print(f"[DEBUG] analyze_transaction_doc: pages_scanned={pages_scanned}, pages_used={pages_used}")

    # 4) Single-shot “red flag/risk” branch
    lower_q = user_query.lower()
    if "red flag" in lower_q or "risk" in lower_q:
        prompt = f"""
You are a Private Equity analyst. Identify ALL “Potential Risk” or “Red Flag” statements
in the following document excerpt. Return each as plain English lines (no bullets, no markdown).

Query: {user_query}

--- DOCUMENT EXCERPT START ---
{selected_text}
--- DOCUMENT EXCERPT END ---
""".strip()

        messages = [
            {"role": "system", "content": "You are a knowledgeable private equity investment analyst."},
            {"role": "user",   "content": prompt}
        ]
        print(f"[DEBUG] analyze_transaction_doc: entering single-shot RED FLAG path")
        single_shot_response = call_deepseek_chat(messages, temperature=0.0, max_tokens=1000)
        print(f"[DEBUG] analyze_transaction_doc: single_shot_response length {len(single_shot_response)}")
        return {
            "answer": single_shot_response,
            "pages_scanned": pages_scanned,
            "relevant_pages": pages_used,
            "chunks_used": None
        }

    # 5) Otherwise, layered summarization:

    # 5a) Chunk the selected_text
    chunks = chunk_text(selected_text, max_tokens=chunk_size)
    print(f"[DEBUG] analyze_transaction_doc: total chunks after chunk_text = {len(chunks)}")

    # 5b) If too many chunks (>20), single-shot fallback
    if len(chunks) > 20:
        prompt = f"""
You are a Private Equity analyst. Answer the following query based on this document excerpt:
“{user_query}”

--- DOCUMENT EXCERPT START ---
{selected_text}
--- DOCUMENT EXCERPT END ---

Give a concise answer as plain text (no markdown).
""".strip()

        messages = [
            {"role": "system", "content": "You are a knowledgeable private equity investment analyst."},
            {"role": "user",   "content": prompt}
        ]
        print(f"[DEBUG] analyze_transaction_doc: too many chunks ({len(chunks)}), entering single-shot fallback")
        single_shot_response = call_deepseek_chat(messages, temperature=0.3, max_tokens=1000)
        print(f"[DEBUG] analyze_transaction_doc: single_shot_fallback_response length {len(single_shot_response)}")
        return {
            "answer": single_shot_response,
            "pages_scanned": pages_scanned,
            "relevant_pages": pages_used,
            "chunks_used": len(chunks)
        }

    # 5c) If ≤ 20 chunks: summarize each chunk
    chunk_summaries: list[str] = []
    for idx, chunk in enumerate(chunks, start=1):
        print(f"[DEBUG] analyze_transaction_doc: summarizing chunk {idx}/{len(chunks)}")
        summary = summarize_chunk(chunk, custom_prompt=chunk_prompt)
        chunk_summaries.append(summary)

    # 5d) Aggregate those chunk summaries into final JSON
    aggregate_json_str = aggregate_summaries(chunk_summaries, custom_prompt=aggregate_prompt)
    try:
        aggregate_json = json.loads(aggregate_json_str)
    except json.JSONDecodeError:
        aggregate_json = {"raw_aggregate_text": aggregate_json_str}

    result: dict[str, object] = {
        "aggregate_analysis": aggregate_json,
        "pages_scanned": pages_scanned,
        "relevant_pages": pages_used,
        "chunks_used": len(chunks)
    }

    # 6) (Optional) Run deep dives if requested
    if run_deep_dives:
        if not deep_dive_sections:
            deep_dive_sections = [
                "Market Analysis",
                "Financial Performance",
                "Operational Risks",
                "Exit Strategy"
            ]

        combined_summaries_text = "\n\n".join(chunk_summaries)
        print(f"[DEBUG] analyze_transaction_doc: running deep dives on sections {deep_dive_sections}")
        for section in deep_dive_sections:
            section_key = section.replace(" ", "_")
            custom = None
            if deep_dive_prompts and section_key in deep_dive_prompts:
                custom = deep_dive_prompts[section_key]

            section_md = deep_dive_section(section, combined_summaries_text, custom_prompt=custom)
            result[f"deep_dive_{section_key}"] = section_md

    print(f"[DEBUG] analyze_transaction_doc: completed analysis, returning result")
    return result