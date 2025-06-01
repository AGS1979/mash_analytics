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
    Returns the number of tokens in `string` using a fixed tiktoken encoding ("cl100k_base").
    We cannot rely on encoding_for_model("deepseek-chat") because that model name is not recognized.
    """
    encoding = tiktoken.get_encoding("cl100k_base")
    n = len(encoding.encode(string))
    print(f"[DEBUG] num_tokens_from_string: {n} tokens for text starting '{string[:30].replace(chr(10), ' ')}...'")
    return n

def chunk_text(text: str, max_tokens: int = 1800) -> list[str]:
    """
    Splits `text` into substrings, each containing <= max_tokens tokens.
    Prefers splitting on double-newline (paragraph) boundaries, but
    will fall back to sentences if a paragraph alone is too large.
    """
    print(f"[DEBUG] chunk_text: splitting text of length {len(text)} characters with max_tokens={max_tokens}")
    paragraphs = text.split("\n\n")
    chunks: list[str] = []
    current_chunk = ""
    current_tokens = 0

    for para_index, para in enumerate(paragraphs):
        para_tokens = num_tokens_from_string(para)
        # Print debug of paragraph size
        print(f"[DEBUG] Paragraph {para_index+1}/{len(paragraphs)}: {para_tokens} tokens")

        if para_tokens > max_tokens:
            # If one paragraph is too big, split by sentences
            sentences = para.split(". ")
            for sentence_index, sentence in enumerate(sentences):
                sent = sentence.strip()
                if not sent:
                    continue
                sent_tokens = num_tokens_from_string(sent)
                # If adding this sentence would overflow, flush current_chunk
                if current_tokens + sent_tokens + 10 > max_tokens:
                    if current_chunk:
                        print(f"[DEBUG] chunk_text: flushing chunk (by sentence) with {current_tokens} tokens")
                        chunks.append(current_chunk.strip())
                    current_chunk = sent + ". "
                    current_tokens = sent_tokens
                else:
                    current_chunk += sent + ". "
                    current_tokens += sent_tokens
                print(f"[DEBUG]   -> sentence {sentence_index+1}: {sent_tokens} tokens, current_chunk now {current_tokens} tokens")
        else:
            # Try to append this entire paragraph to current_chunk
            if current_tokens + para_tokens + 20 > max_tokens:
                if current_chunk:
                    print(f"[DEBUG] chunk_text: flushing chunk (by paragraph) with {current_tokens} tokens")
                    chunks.append(current_chunk.strip())
                current_chunk = para + "\n\n"
                current_tokens = para_tokens
                print(f"[DEBUG] Starting new chunk with paragraph {para_index+1}: {para_tokens} tokens")
            else:
                current_chunk += para + "\n\n"
                current_tokens += para_tokens
                print(f"[DEBUG] Appended paragraph {para_index+1} to current_chunk, now {current_tokens} tokens")

    # Flush remainder
    if current_chunk:
        print(f"[DEBUG] chunk_text: flushing final chunk with {current_tokens} tokens")
        chunks.append(current_chunk.strip())

    print(f"[DEBUG] chunk_text: total chunks created: {len(chunks)}")
    return chunks

# ---------------------------------------------------
# 2) TEXT EXTRACTION FOR MULTIPLE FORMATS
# ---------------------------------------------------

def extract_text_from_pdf(filepath: str) -> str:
    """
    Uses PyMuPDF to extract all text from a PDF file, concatenated page by page.
    """
    print(f"[DEBUG] extract_text_from_pdf: opening {filepath}")
    doc = fitz.open(filepath)
    all_text = []
    for page_index, page in enumerate(doc):
        page_text = page.get_text()
        all_text.append(page_text)
        print(f"[DEBUG]   Extracted text from page {page_index+1}, length {len(page_text)} chars")
    combined = "\n\n".join(all_text)
    print(f"[DEBUG] extract_text_from_pdf: total length {len(combined)} chars")
    return combined

def extract_text_from_pdf_by_page(filepath: str) -> list[str]:
    """
    Uses PyMuPDF to extract text **page by page** (returns a list where
    index 0 is page 1, index 1 is page 2, etc.).
    """
    print(f"[DEBUG] extract_text_from_pdf_by_page: opening {filepath}")
    doc = fitz.open(filepath)
    pages = []
    for page_index, page in enumerate(doc):
        text = page.get_text()
        pages.append(text)
        print(f"[DEBUG]   Page {page_index+1}: {len(text)} chars")
    return pages

def extract_text_from_docx(filepath: str) -> str:
    """
    Uses python-docx to read all paragraphs from a .docx file.
    """
    print(f"[DEBUG] extract_text_from_docx: opening {filepath}")
    doc = Document(filepath)
    paragraphs = [para.text for para in doc.paragraphs if para.text.strip() != ""]
    print(f"[DEBUG]   Extracted {len(paragraphs)} paragraphs from DOCX")
    combined = "\n\n".join(paragraphs)
    print(f"[DEBUG] extract_text_from_docx: total length {len(combined)} chars")
    return combined

def extract_text_from_pptx(filepath: str) -> str:
    """
    Uses python-pptx to extract all text from every slide in a .pptx file.
    """
    print(f"[DEBUG] extract_text_from_pptx: opening {filepath}")
    prs = Presentation(filepath)
    slide_texts = []
    for slide_index, slide in enumerate(prs.slides):
        box_texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                t = shape.text.strip()
                if t:
                    box_texts.append(t)
        combined_slide = "\n".join(box_texts)
        slide_texts.append(combined_slide)
        print(f"[DEBUG]   Slide {slide_index+1}: {len(combined_slide)} chars")
    combined = "\n\n".join(slide_texts)
    print(f"[DEBUG] extract_text_from_pptx: total length {len(combined)} chars")
    return combined

def extract_text_from_excel(filepath: str) -> str:
    """
    Uses pandas/openpyxl to read each sheet in an Excel file and
    flatten all cells into one big text blob.
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
            print(f"[DEBUG]   Sheet '{sheet}': {non_empty_rows.shape[0]} non-empty rows")
    combined = "\n\n".join(sheet_texts)
    print(f"[DEBUG] extract_text_from_excel: total length {len(combined)} chars")
    return combined

def extract_text(filepath: str) -> str:
    """
    Dispatch by file extension: .pdf, .docx, .pptx, .xls/.xlsx/.xlsm.
    Raises ValueError on unsupported extension.
    """
    ext = os.path.splitext(filepath)[1].lower()
    print(f"[DEBUG] extract_text: dispatching for extension '{ext}'")
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
    Calls DeepSeek Chat (OpenAI‐compatible chat.completions). Returns the assistant’s reply string.
    """
    print(f"[DEBUG] call_deepseek_chat: sending request to {DEEPSEEK_CHAT_URL}")
    print(f"[DEBUG]   model: {model}, temperature: {temperature}, max_tokens: {max_tokens}")
    # Print truncated messages for debugging
    for idx, msg in enumerate(messages):
        role = msg.get("role", "<no role>")
        content = msg.get("content", "")
        truncated = content[:200].replace("\n", " ") + ("..." if len(content) > 200 else "")
        print(f"[DEBUG]   message[{idx}][{role}]: {truncated}")

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

    print(f"[DEBUG] call_deepseek_chat: HTTP status {resp.status_code}")
    if resp.status_code != 200:
        try:
            detail = resp.json()
        except ValueError:
            detail = resp.text
        raise RuntimeError(f"DeepSeek returned HTTP {resp.status_code}: {detail}")

    data = resp.json()
    # Print truncated response for debugging
    raw_resp = data.get("choices", [{}])[0].get("message", {}).get("content", "")
    print(f"[DEBUG] call_deepseek_chat: received response (first 200 chars): '{raw_resp[:200].replace(chr(10), ' ')}...'")
    return raw_resp

# ---------------------------------------------------
# 4) PAGE IDENTIFICATION FOR ANY USER QUERY
# ---------------------------------------------------

PAGE_BATCH_SIZE = 20  # ask “which pages matter?” in batches of 20 at a time

def get_relevant_pages_chunked(text_by_page: list[str], user_query: str) -> set[int]:
    """
    Loop through `text_by_page` in batches of PAGE_BATCH_SIZE, sending a short
    prompt: “Which page numbers (1-based) are relevant to this query?” 
    Returns a set of 0-based page indices.
    """
    total_pages = len(text_by_page)
    print(f"[DEBUG] get_relevant_pages_chunked: total_pages={total_pages}, user_query='{user_query}'")
    relevant_pages = set()

    for start in range(0, total_pages, PAGE_BATCH_SIZE):
        end = min(start + PAGE_BATCH_SIZE, total_pages)
        chunk_pages = text_by_page[start:end]
        print(f"[DEBUG]   Evaluating pages {start+1} to {end} for relevance")

        # Build a brief prompt listing page numbers + snippet
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
            print(f"[WARNING] get_relevant_pages_chunked: skipping batch {start+1}-{end} due to error: {e}")
            continue

        # Parse out integers from the response to pick page numbers
        print(f"[DEBUG]   get_relevant_pages_chunked: raw response: '{resp_text[:200].replace(chr(10),' ')}...'")
        for token in resp_text.replace(",", " ").split():
            if token.isdigit():
                pnum = int(token)
                if 1 <= pnum <= total_pages:
                    relevant_pages.add(pnum - 1)  # zero‐based
                    print(f"[DEBUG]     Found relevant page: {pnum}")

    print(f"[DEBUG] get_relevant_pages_chunked: total relevant_pages found: {sorted(relevant_pages)}")
    return relevant_pages

# ---------------------------------------------------
# 5) LAYERED SUMMARIZATION + ANALYSIS ROUTINES
# ---------------------------------------------------

def summarize_chunk(
    chunk_text: str,
    custom_prompt: str | None = None
) -> str:
    """
    Summarize one chunk. By default, extract bullets of “Key Insights” and “Potential Risks.”
    If custom_prompt is given, replace {{TEXT}} with chunk_text.
    """
    if custom_prompt:
        prompt = custom_prompt.replace("{{TEXT}}", chunk_text)
        print(f"[DEBUG] summarize_chunk: using custom_prompt, first 200 chars: '{prompt[:200].replace(chr(10),' ')}...'")
    else:
        prompt = f"""
You are a Private Equity analyst. From the following excerpt, extract any “Potential Risk” or “Key Insight”
statements, each as a separate bullet. Keep each bullet under 70 words.

--- EXCERPT START ---
{chunk_text}
--- EXCERPT END ---

Respond only with Markdown bullets prefaced with “Key Insight:” or “Potential Risk:”.
""".strip()
        print(f"[DEBUG] summarize_chunk: using default prompt, chunk length={len(chunk_text)} chars")

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

    If custom_prompt is provided, replace {{SUMMARIES}} with joined bullet text.
    """
    combined_text = "\n\n".join(summaries)
    print(f"[DEBUG] aggregate_summaries: combining {len(summaries)} chunk summaries, total length {len(combined_text)} chars")

    if custom_prompt:
        prompt = custom_prompt.replace("{{SUMMARIES}}", combined_text)
        print(f"[DEBUG] aggregate_summaries: using custom_prompt, first 200 chars: '{prompt[:200].replace(chr(10),' ')}...'")
    else:
        prompt = f"""
You are a top-tier private equity partner. Below are bullet-point summaries from different sections of a deal document.
1) Eliminate redundancies
2) Create a cohesive “Executive Summary” (1 paragraph)
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
        print(f"[DEBUG] aggregate_summaries: using default prompt")

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
    Perform a detailed analysis of a given section_name (e.g. “Market Analysis”).
    If custom_prompt is provided, substitute {{SECTION_NAME}} and {{TEXT}}.
    """
    if custom_prompt:
        prompt = (custom_prompt
                  .replace("{{SECTION_NAME}}", section_name)
                  .replace("{{TEXT}}", combined_text))
        print(f"[DEBUG] deep_dive_section: using custom_prompt for {section_name}, first 200 chars: '{prompt[:200].replace(chr(10),' ')}...'")
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

Respond only in Markdown with headings and bullet points.
""".strip()
        print(f"[DEBUG] deep_dive_section: using default deep-dive for {section_name}")

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
    1) Extract the PDF as a list of pages.
    2) Identify which pages matter for user_query (using get_relevant_pages_chunked).
    3) If relevant_pages is non-empty, concatenate only those pages; otherwise use all pages.
    4) If the user_query explicitly mentions “red flag” or “risk,” do a single‐shot prompt
       over the concatenated text and return the raw answer.
    5) Otherwise, attempt layered summarization:
         a) Chunk the selected_text into ≤ chunk_size tokens each.
         b) If the number of chunks is ≤ 20, summarize each chunk & aggregate.
         c) If > 20 chunks, fall back to a single‐shot aggregated‐prompt (same as “red flag” path).
    6) Optionally, run deep dives on the aggregated bullet summaries.
    """
    print(f"[DEBUG] analyze_transaction_doc: START, filepath={filepath}, user_query='{user_query}'")

    # 1) Extract the PDF into text_by_page
    text_by_page = extract_text_from_pdf_by_page(filepath)
    print(f"[DEBUG] analyze_transaction_doc: extracted {len(text_by_page)} pages from PDF")

    # 2) Ask DeepSeek which pages are relevant to the user's query
    relevant_pages = get_relevant_pages_chunked(text_by_page, user_query)

    # 3) Decide whether to use only relevant pages or all pages
    if relevant_pages:
        selected_text = "\n\n".join(text_by_page[i] for i in sorted(relevant_pages))
        pages_scanned = len(text_by_page)
        pages_used = sorted(relevant_pages)
        print(f"[DEBUG] analyze_transaction_doc: using only {len(pages_used)} relevant pages: {pages_used}")
    else:
        selected_text = "\n\n".join(text_by_page)
        pages_scanned = len(text_by_page)
        pages_used = list(range(pages_scanned))
        print(f"[DEBUG] analyze_transaction_doc: no relevant pages found, using all {pages_scanned} pages")

    # 4) If the user specifically asks for “red flag” or “risk,” do single‐shot
    lower_q = user_query.lower()
    if "red flag" in lower_q or "risk" in lower_q:
        print(f"[DEBUG] analyze_transaction_doc: single-shot path for red-flag/risk")
        prompt = f"""
You are a Private Equity analyst. Identify ALL “Potential Risk” or “Red Flag” statements
in the following document excerpt. Return each risk/red‐flag as a separate bullet.

Query: {user_query}

--- DOCUMENT EXCERPT START ---
{selected_text}
--- DOCUMENT EXCERPT END ---

Respond only with Markdown bullets prefaced with “Potential Risk:” or “Red Flag:”.
""".strip()

        messages = [
            {"role": "system", "content": "You are a knowledgeable private equity investment analyst."},
            {"role": "user",   "content": prompt}
        ]
        single_shot_response = call_deepseek_chat(messages, temperature=0.0, max_tokens=1000)
        print(f"[DEBUG] analyze_transaction_doc: returning single-shot response")
        return {
            "answer": single_shot_response,
            "pages_scanned": pages_scanned,
            "relevant_pages": pages_used,
            "chunks_used": None
        }

    # 5) Otherwise, attempt layered summarization:

    #   5a) Break selected_text into chunks of ≤ chunk_size tokens
    chunks = chunk_text(selected_text, max_tokens=chunk_size)
    print(f"[DEBUG] analyze_transaction_doc: total chunks created: {len(chunks)}")

    #   5b) If too many chunks (> 20), fall back to a single‐shot answer
    if len(chunks) > 20:
        print(f"[DEBUG] analyze_transaction_doc: >20 chunks, falling back to single-shot aggregated summary")
        prompt = f"""
You are a Private Equity analyst. Answer the following query based on this document excerpt:
“{user_query}”

--- DOCUMENT EXCERPT START ---
{selected_text}
--- DOCUMENT EXCERPT END ---

Give a concise answer.
""".strip()

        messages = [
            {"role": "system", "content": "You are a knowledgeable private equity investment analyst."},
            {"role": "user",   "content": prompt}
        ]
        single_shot_response = call_deepseek_chat(messages, temperature=0.3, max_tokens=1000)
        print(f"[DEBUG] analyze_transaction_doc: returning single-shot fallback response")
        return {
            "answer": single_shot_response,
            "pages_scanned": pages_scanned,
            "relevant_pages": pages_used,
            "chunks_used": len(chunks)
        }

    #   5c) We have ≤ 20 chunks: summarize each chunk
    print(f"[DEBUG] analyze_transaction_doc: summarizing each chunk (count={len(chunks)})")
    chunk_summaries: list[str] = []
    for idx, chunk in enumerate(chunks):
        print(f"[DEBUG]   Summarizing chunk {idx+1}/{len(chunks)}, length {len(chunk)} chars")
        summary = summarize_chunk(chunk, custom_prompt=chunk_prompt)
        chunk_summaries.append(summary)

    #   5d) Aggregate those chunk summaries into a JSON
    print(f"[DEBUG] analyze_transaction_doc: aggregating {len(chunk_summaries)} chunk summaries")
    aggregate_json_str = aggregate_summaries(chunk_summaries, custom_prompt=aggregate_prompt)
    try:
        aggregate_json = json.loads(aggregate_json_str)
        print(f"[DEBUG] analyze_transaction_doc: parsed aggregate JSON successfully")
    except json.JSONDecodeError:
        aggregate_json = {"raw_aggregate_text": aggregate_json_str}
        print(f"[WARNING] analyze_transaction_doc: failed to parse aggregate JSON, returning raw text")

    result: dict[str, object] = {
        "aggregate_analysis": aggregate_json,
        "pages_scanned": pages_scanned,
        "relevant_pages": pages_used,
        "chunks_used": len(chunks)
    }

    # 6) (Optional) Run deep dives if requested
    if run_deep_dives:
        print(f"[DEBUG] analyze_transaction_doc: running deep dives on aggregated summaries")
        if not deep_dive_sections:
            deep_dive_sections = [
                "Market Analysis",
                "Financial Performance",
                "Operational Risks",
                "Exit Strategy"
            ]

        combined_summaries_text = "\n\n".join(chunk_summaries)
        for section in deep_dive_sections:
            section_key = section.replace(" ", "_")
            custom = None
            if deep_dive_prompts and section_key in deep_dive_prompts:
                custom = deep_dive_prompts[section_key]
                print(f"[DEBUG]   Using custom prompt for deep dive '{section}'")
            else:
                print(f"[DEBUG]   Using default deep dive for '{section}'")

            section_md = deep_dive_section(section, combined_summaries_text, custom_prompt=custom)
            result[f"deep_dive_{section_key}"] = section_md

    print(f"[DEBUG] analyze_transaction_doc: END, returning result dictionary")
    return result
