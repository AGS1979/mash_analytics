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
# 0) LOAD API KEYS / ENDPOINTS FROM ENVIRONMENT
# ---------------------------------------------------

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_CHAT_URL = os.getenv(
    "DEEPSEEK_CHAT_URL",
    "https://api.deepseek.com/v1/chat/completions"
)

if not DEEPSEEK_API_KEY:
    raise RuntimeError("Please set DEEPSEEK_API_KEY in your environment.")

# ---------------------------------------------------
# 1) UTILITIES: token counting & chunking
# ---------------------------------------------------

def num_tokens_from_string(string: str) -> int:
    """
    Returns the number of tokens in `string` using a fixed tiktoken encoding
    ("cl100k_base"), since "deepseek-chat" is not recognized by tiktoken's encoding_for_model.
    """
    encoding = tiktoken.get_encoding("cl100k_base")
    return len(encoding.encode(string))


def chunk_text(text: str, max_tokens: int = 1800) -> list[str]:
    """
    Splits `text` into substrings, each containing <= max_tokens tokens.
    Prefers splitting on double‐newline (paragraph) boundaries, but
    will fall back to sentences if a paragraph alone is too large.
    """
    paragraphs = text.split("\n\n")
    chunks: list[str] = []
    current_chunk = ""
    current_tokens = 0

    for para in paragraphs:
        para_tokens = num_tokens_from_string(para)

        if para_tokens > max_tokens:
            # Paragraph itself is too big. Break into sentences.
            sentences = para.split(". ")
            for sentence in sentences:
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

    return chunks

# ---------------------------------------------------
# 2) TEXT EXTRACTION FOR MULTIPLE FORMATS
# ---------------------------------------------------

def extract_text_from_pdf(filepath: str) -> str:
    """
    Uses PyMuPDF to extract all text from a PDF file (all pages concatenated).
    """
    doc = fitz.open(filepath)
    all_text = []
    for page in doc:
        all_text.append(page.get_text())
    return "\n\n".join(all_text)

def extract_text_from_pdf_by_page(filepath: str) -> list[str]:
    """
    Uses PyMuPDF to extract text **page by page** (returns a list where
    index 0 is page 1, index 1 is page 2, etc.).
    """
    doc = fitz.open(filepath)
    return [page.get_text() for page in doc]

def extract_text_from_docx(filepath: str) -> str:
    """
    Uses python-docx to read all paragraphs from a .docx file.
    """
    doc = Document(filepath)
    all_paragraphs = [para.text for para in doc.paragraphs if para.text.strip() != ""]
    return "\n\n".join(all_paragraphs)

def extract_text_from_pptx(filepath: str) -> str:
    """
    Uses python-pptx to extract all text from every slide in a .pptx file.
    """
    prs = Presentation(filepath)
    slide_texts = []
    for slide in prs.slides:
        box_texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text = shape.text.strip()
                if text:
                    box_texts.append(text)
        if box_texts:
            slide_texts.append("\n".join(box_texts))
    return "\n\n".join(slide_texts)

def extract_text_from_excel(filepath: str) -> str:
    """
    Uses pandas/openpyxl to read each sheet in an Excel file and
    flatten all cells into one big text blob.
    """
    xlsx = pd.ExcelFile(filepath)
    sheet_texts: list[str] = []
    for sheet in xlsx.sheet_names:
        df = pd.read_excel(xlsx, sheet_name=sheet, dtype=str)
        rows = df.fillna("").apply(lambda r: " | ".join(r.values), axis=1)
        non_empty_rows = rows[rows.str.strip() != ""]
        if not non_empty_rows.empty:
            sheet_block = f"--- SHEET: {sheet} ---\n" + "\n".join(non_empty_rows.tolist())
            sheet_texts.append(sheet_block)
    return "\n\n".join(sheet_texts)

def extract_text(filepath: str) -> str:
    """
    Dispatch by file extension: .pdf, .docx, .pptx, .xls/.xlsx/.xlsm.
    Raises ValueError on unsupported extension.
    """
    ext = os.path.splitext(filepath)[1].lower()
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
# 3) LLM (DeepSeek Chat) WRAPPER
# ---------------------------------------------------

def call_deepseek_chat(
    messages: list[dict[str, str]],
    model: str = "deepseek-chat",
    temperature: float = 0.3,
    max_tokens: int = 1500,
    timeout_sec: int = 60
) -> str:
    """
    Calls DeepSeek Chat (OpenAI‐style / chat.completions endpoint).
    Returns the assistant’s reply string.
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

    try:
        resp = requests.post(
            DEEPSEEK_CHAT_URL,
            headers=headers,
            json=payload,
            timeout=timeout_sec
        )
    except requests.RequestException as e:
        raise RuntimeError(f"Failed to connect to DeepSeek: {e}")

    if resp.status_code != 200:
        try:
            detail = resp.json()
        except ValueError:
            detail = resp.text
        raise RuntimeError(f"DeepSeek returned HTTP {resp.status_code}: {detail}")

    data = resp.json()
    # We assume DeepSeek returns: data["choices"][0]["message"]["content"]
    return data["choices"][0]["message"]["content"]

# ---------------------------------------------------
# 4) PAGE IDENTIFICATION FOR ANY USER QUERY
# ---------------------------------------------------

# How many pages per batch to send to DeepSeek when asking “which pages matter?”
PAGE_BATCH_SIZE = 20

def get_relevant_pages_chunked(text_by_page: list[str], user_query: str) -> set[int]:
    """
    Loop through `text_by_page` in batches of size PAGE_BATCH_SIZE, sending a short DeepSeek prompt
    that asks “Which of these page numbers are relevant to <user_query>?”. Returns a set of 0-based page indices.
    """
    total_pages = len(text_by_page)
    relevant_pages = set()

    for start in range(0, total_pages, PAGE_BATCH_SIZE):
        end = min(start + PAGE_BATCH_SIZE, total_pages)
        chunk_pages = text_by_page[start:end]

        # Build a brief prompt listing page numbers + first ~800 chars of each.
        prompt = (
            "Below are text snippets from pages of a PDF. Identify ONLY the page numbers (1-based) "
            f"that are relevant to this query:\n\nQuery: {user_query}\n\n"
        )
        for i, page_text in enumerate(chunk_pages):
            snippet = page_text[:800].replace("\n", " ")
            prompt += f"Page {start + i + 1}: {snippet}\n\n"

        messages = [
            {"role": "system", "content": "You are an expert document analyst."},
            {"role": "user",   "content": prompt}
        ]

        # Call DeepSeek in “zero-temperature” mode for deterministic page selection
        try:
            resp = call_deepseek_chat(messages, model="deepseek-chat", temperature=0.0, max_tokens=200)
        except RuntimeError as e:
            # If one batch fails, skip it and continue
            continue

        # Parse out all integers in the response; treat them as page numbers.
        found = set()
        for token in resp.replace(",", " ").split():
            if token.isdigit():
                page_num = int(token)
                if 1 <= page_num <= total_pages:
                    found.add(page_num - 1)  # convert to 0-based
        relevant_pages.update(found)

    return relevant_pages

# ---------------------------------------------------
# 5) LAYERED SUMMARIZATION + ANALYSIS ROUTINES
# ---------------------------------------------------

def summarize_chunk(
    chunk_text: str,
    custom_prompt: str | None = None
) -> str:
    """
    Summarize one chunk. By default, this focuses on “Key Insights” and “Potential Risks.”
    If `custom_prompt` is provided, use that verbatim, replacing {{TEXT}} with chunk_text.
    """
    if custom_prompt:
        prompt = custom_prompt.replace("{{TEXT}}", chunk_text)
    else:
        prompt = f"""
You are a Private Equity analyst. From the following excerpt, extract any “Potential Risk” or “Key Insight” statements,
each as a separate bullet. Keep each bullet under 70 words.

--- EXCERPT START ---
{chunk_text}
--- EXCERPT END ---

Respond only with Markdown bullets prefaced with “Key Insight:” or “Potential Risk:”.
""".strip()

    messages = [
        {"role": "system", "content": "You are a knowledgeable private equity investment analyst."},
        {"role": "user",   "content": prompt}
    ]
    return call_deepseek_chat(messages)

def aggregate_summaries(
    summaries: list[str],
    custom_prompt: str | None = None
) -> str:
    """
    Given a list of chunk‐level bullet summaries, produce a final JSON with keys:
      - Executive_Summary
      - Key_Investment_Insights
      - Risks_and_Red_Flags
      - High_Level_Valuation_Guidance

    If `custom_prompt` is provided, use it verbatim, replacing {{SUMMARIES}} with joined bullet summaries.
    """
    combined_text = "\n\n".join(summaries)

    if custom_prompt:
        prompt = custom_prompt.replace("{{SUMMARIES}}", combined_text)
    else:
        prompt = f"""
You are a top-tier private equity partner. Below are bullet-point summaries from different sections of a deal document.
1) Eliminate redundancies
2) Write a cohesive “Executive Summary” (1 paragraph)
3) List “Top 5 Key Investment Insights” (bullets)
4) List “Top 5 Risks and Red Flags” (bullets)
5) Provide “High-Level Valuation Guidance” (2–3 bullets)

--- BULLET SUMMARIES START ---
{combined_text}
--- BULLET SUMMARIES END ---

Respond in JSON format exactly like this example:
{{
  "Executive_Summary": "...",
  "Key_Investment_Insights": ["...", "...", "...", "...", "..."],
  "Risks_and_Red_Flags": ["...", "...", "...", "...", "..."],
  "High_Level_Valuation_Guidance": ["...", "..."]
}}
""".strip()

    messages = [
        {"role": "system", "content": "You are a senior private equity partner, extremely concise and factual."},
        {"role": "user",   "content": prompt}
    ]
    return call_deepseek_chat(messages)

def deep_dive_section(
    section_name: str,
    combined_text: str,
    custom_prompt: str | None = None
) -> str:
    """
    Perform a specialized deep-dive for a given section (e.g. “Market Analysis”).
    If `custom_prompt` is provided, plug in {{SECTION_NAME}} and {{TEXT}}.
    """
    if custom_prompt:
        prompt = (custom_prompt
                  .replace("{{SECTION_NAME}}", section_name)
                  .replace("{{TEXT}}", combined_text))
    else:
        prompt = f"""
You are a PE sector specialist. Perform a detailed **{section_name}** analysis on the following deal document excerpt:

--- DOCUMENT EXCERPT START ---
{combined_text}
--- DOCUMENT EXCERPT END ---

Structure your answer as:
1) <Section Header> (e.g. Market Overview)
2) A short 2-sentence summary
3) A bullet list of 4–6 observations or considerations
4) A final “Implication for Deal” line (< 50 words)

Respond only in Markdown with headings and bullets.
""".strip()

    messages = [
        {"role": "system", "content": "You are a domain expert in private equity deals."},
        {"role": "user",   "content": prompt}
    ]
    return call_deepseek_chat(messages, temperature=0.25, max_tokens=1000)

# ---------------------------------------------------
# 6) MAIN “ANALYZE TRANSACTION” WORKFLOW
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
    1) Extract the PDF by page
    2) Identify which pages matter for `user_query`
    3) If we found relevant pages, only concatenate those; otherwise, use the entire document
    4) Chunk the selected text and summarize each chunk (Layer 1)
    5) Aggregate chunk summaries (Layer 2)
    6) (Optional) Run deep dives on the aggregated summaries (Layer 3)
    7) Return a dict containing:
         - "aggregate_analysis": parsed JSON from aggregate_summaries
         - "deep_dive_<SectionName>" (Markdown) for each requested section
         - "pages_scanned" / "relevant_pages" for debugging
    """

    # 1) Extract PDF into text_by_page
    text_by_page = extract_text_from_pdf_by_page(filepath)

    # 2) Ask DeepSeek which pages are relevant to the user's query
    relevant_pages = get_relevant_pages_chunked(text_by_page, user_query)

    # 3) Decide whether to use only relevant pages or all pages
    if relevant_pages:
        # Concatenate only the relevant pages (in ascending order)
        selected_text = "\n\n".join(text_by_page[i] for i in sorted(relevant_pages))
        pages_scanned = len(text_by_page)
        pages_used_count = len(relevant_pages)
    else:
        # If no pages found, fall back to entire document
        selected_text = "\n\n".join(text_by_page)
        pages_scanned = len(text_by_page)
        pages_used_count = len(text_by_page)

    # 4) Chunk the selected_text
    chunks = chunk_text(selected_text, max_tokens=chunk_size)

    # Summarize each chunk (Layer 1)
    chunk_summaries: list[str] = []
    for idx, chunk in enumerate(chunks):
        summary = summarize_chunk(chunk, custom_prompt=chunk_prompt)
        chunk_summaries.append(summary)

    # 5) Aggregate chunk‐level summaries into final JSON (Layer 2)
    aggregate_json_str = aggregate_summaries(chunk_summaries, custom_prompt=aggregate_prompt)
    try:
        aggregate_json = json.loads(aggregate_json_str)
    except json.JSONDecodeError:
        aggregate_json = {"raw_aggregate_text": aggregate_json_str}

    result: dict[str, object] = {
        "aggregate_analysis": aggregate_json,
        "pages_scanned": pages_scanned,
        "relevant_pages": sorted(relevant_pages or list(range(pages_scanned))),
        "chunks_used": len(chunks)
    }

    # 6) Optional deep dives (Layer 3)
    if run_deep_dives:
        if not deep_dive_sections:
            deep_dive_sections = [
                "Market Analysis",
                "Financial Performance",
                "Operational Risks",
                "Exit Strategy"
            ]

        # Combine all chunk summaries as input to each deep dive
        combined_summaries_text = "\n\n".join(chunk_summaries)

        for section in deep_dive_sections:
            section_key = section.replace(" ", "_")
            custom = None
            if deep_dive_prompts and section_key in deep_dive_prompts:
                custom = deep_dive_prompts[section_key]

            section_md = deep_dive_section(section, combined_summaries_text, custom_prompt=custom)
            result[f"deep_dive_{section_key}"] = section_md

    return result
