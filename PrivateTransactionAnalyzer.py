# deal_analyzer.py

import os
import json
import math
import fitz                     # PyMuPDF
import tiktoken                 # for token counting
import requests                 # for HTTP calls to DeepSeek Chat
import pandas as pd             # for Excel extraction
from docx import Document       # for .docx extraction
from pptx import Presentation   # for .pptx extraction
from typing import List, Dict, Optional, Union, Set

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
    Returns the number of tokens in `string` using the 'cl100k_base' encoding.
    """
    encoding = tiktoken.get_encoding("cl100k_base")
    return len(encoding.encode(string))

def chunk_text(text: str, max_tokens: int = 1800) -> List[str]:
    """
    Splits `text` into substrings, each containing ≤ max_tokens tokens.
    First tries paragraph (double-newline) splits; if a paragraph is itself too big,
    it falls back to splitting on sentences.
    """
    paragraphs = text.split("\n\n")
    chunks: List[str] = []
    current_chunk = ""
    current_tokens = 0

    for para in paragraphs:
        para = para.strip()
        if not para:
            continue
        para_tokens = num_tokens_from_string(para)

        if para_tokens > max_tokens:
            # Paragraph is too large: break into sentences
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
            # Try to append the entire paragraph to current_chunk
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

def reduce_chunks_hierarchically(
    raw_chunks: List[str],
    max_chunks: int,
    chunk_size: int,
    custom_prompt: Optional[str]
) -> List[str]:
    """
    If raw_chunks has more than max_chunks, group them into roughly equal-sized buckets,
    summarize each bucket with summarize_chunk(), and repeat until there are ≤ max_chunks summaries.
    """
    from math import ceil

    summaries = raw_chunks
    iteration = 0

    while len(summaries) > max_chunks:
        iteration += 1
        total = len(summaries)
        bucket_size = ceil(total / max_chunks)
        print(f"[DEBUG] reduce_chunks_hierarchically: iteration {iteration}, {total} chunks → bucket_size={bucket_size}")

        next_round: List[str] = []
        for i in range(0, total, bucket_size):
            bucket = summaries[i : i + bucket_size]
            combined_bucket = "\n\n".join(bucket)
            print(f"[DEBUG] reduce_chunks_hierarchically: summarizing bucket {i // bucket_size + 1} of size {len(bucket)}")
            summary = summarize_chunk(combined_bucket, custom_prompt=custom_prompt)
            next_round.append(summary)

        summaries = next_round
        print(f"[DEBUG] reduce_chunks_hierarchically: reduced to {len(summaries)} intermediate summaries")

    return summaries

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

def extract_text_from_pdf_by_page(filepath: str) -> List[str]:
    """
    Uses PyMuPDF to extract text *page by page* as a list of strings (index 0 = page 1).
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
    all_paragraphs = [para.text for para in doc.paragraphs if para.text.strip()]
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
    Uses pandas/openpyxl to read each sheet in an Excel file and flatten all cells into text.
    """
    print(f"[DEBUG] extract_text_from_excel: opening {filepath}")
    xlsx = pd.ExcelFile(filepath)
    sheet_texts: List[str] = []
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
    Dispatches based on file extension: .pdf, .docx, .pptx, .xls/.xlsx/.xlsm.
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
    messages: List[Dict[str, str]],
    model: str = "deepseek-chat",
    temperature: float = 0.3,
    max_tokens: int = 1500,
    timeout_sec: int = 60
) -> str:
    """
    Calls DeepSeek Chat (OpenAI-compatible chat.completions). Returns the assistant’s reply string.
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
    # For extra debugging, uncomment:
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
    choice_content = data["choices"][0]["message"]["content"]
    print(f"[DEBUG] call_deepseek_chat: received {len(choice_content)} characters")
    return choice_content

# ---------------------------------------------------
# 5) PAGE IDENTIFICATION FOR ANY USER QUERY
# ---------------------------------------------------

PAGE_BATCH_SIZE = 20  # ask “which pages matter?” in batches of 20 at a time

def get_relevant_pages_chunked(
    text_by_page: List[str],
    user_query: str
) -> Set[int]:
    """
    Loops through `text_by_page` in batches of PAGE_BATCH_SIZE, sending a prompt:
    “Which page numbers (1-based) are relevant to this query?”
    Returns a set of 0-based page indices.
    """
    total_pages = len(text_by_page)
    relevant_pages: Set[int] = set()
    print(f"[DEBUG] get_relevant_pages_chunked: total_pages={total_pages}, user_query='{user_query}'")

    for start in range(0, total_pages, PAGE_BATCH_SIZE):
        end = min(start + PAGE_BATCH_SIZE, total_pages)
        chunk_pages = text_by_page[start:end]

        prompt_lines = [
            "Below are text snippets from pages of a PDF. Identify ONLY the page numbers (1-based) that are relevant to this query.",
            f"Query: {user_query}\n"
        ]
        for i, page_text in enumerate(chunk_pages):
            snippet = page_text[:800].replace("\n", " ")
            prompt_lines.append(f"Page {start + i + 1}: {snippet}\n")

        prompt = "\n".join(prompt_lines)

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

        # Parse out integers in the response; treat them as page numbers
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
    custom_prompt: Optional[str] = None
) -> str:
    """
    Summarize one chunk. By default, extract “Key Insights” and “Potential Risks.”
    Output is plain text paragraphs (no Markdown).
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

def summarize_swot(
    combined_text: str
) -> str:
    """
    If the user specifically wants a SWOT analysis, generate a JSON-formatted SWOT analysis
    with exactly these four keys: "Strengths", "Weaknesses", "Opportunities", and "Threats".
    Each key’s value is a list of concise strings. Do NOT return Markdown or backticks.
    """
    prompt = f"""
You are a Private Equity analyst. From the following excerpt, produce a JSON-formatted SWOT analysis with exactly these four keys:
"Strengths", "Weaknesses", "Opportunities", and "Threats". Each key’s value should be a list of concise strings.
Do NOT include any Markdown, backticks, or extra keys—just return a pure JSON object.

--- EXCERPT START ---
{combined_text}
--- EXCERPT END ---
""".strip()

    messages = [
        {"role": "system", "content": "You are a knowledgeable private equity investment analyst."},
        {"role": "user",   "content": prompt}
    ]
    print(f"[DEBUG] summarize_swot: sending combined excerpt (length {len(combined_text)} chars) to DeepSeek for SWOT")
    response = call_deepseek_chat(messages, temperature=0.3, max_tokens=800)
    print(f"[DEBUG] summarize_swot: received SWOT JSON length {len(response)}")
    return response

def aggregate_summaries(
    summaries: List[str],
    custom_prompt: Optional[str] = None
) -> str:
    """
    Given a list of chunk-level summaries (plain text paragraphs), produce a final JSON
    with keys:
      - Executive_Summary
      - Key_Investment_Insights
      - Risks_and_Red_Flags
      - High_Level_Valuation_Guidance

    Always return plain JSON (no Markdown, no backticks, no code fences).
    """
    combined_text = "\n\n".join(summaries)
    print(f"[DEBUG] aggregate_summaries: combining {len(summaries)} chunk summaries (combined length {len(combined_text)} chars)")

    if custom_prompt:
        prompt = custom_prompt.replace("{{SUMMARIES}}", combined_text).strip()
    else:
        prompt = f"""
You are a senior private equity partner, extremely concise and factual. Below are plain-text paragraph summaries
of different sections of a deal document. Please do the following and return your answer AS A JSON OBJECT ONLY
(no Markdown, no backticks, no code fences):

1) "Executive_Summary": one paragraph under 100 words summarizing the entire deal.
2) "Key_Investment_Insights": an array of 6 concise insights (plain strings).
3) "Risks_and_Red_Flags": an array of 6 concise items (plain strings).
4) "High_Level_Valuation_Guidance": an array of 3 concise points (plain strings).

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
    custom_prompt: Optional[str] = None
) -> str:
    """
    Perform a detailed deep dive for a given section (e.g., “Market Analysis”).
    Output plain English paragraphs with clear headings (no Markdown).
    """
    if custom_prompt:
        prompt = (
            custom_prompt
            .replace("{{SECTION_NAME}}", section_name)
            .replace("{{TEXT}}", combined_text)
            .strip()
        )
    else:
        prompt = f"""
You are a private equity sector specialist. Perform a deeply detailed, multi-paragraph analysis
for the section: "{section_name}". Do NOT use bullets or Markdown—only plain English with optional headings.
Structure your response as follows:

1) A short overview paragraph (2–3 sentences).
2) A detailed 4–5 paragraph discussion touching on domain insights, potential red flags, and implications for deal structure or valuation.
3) A concluding paragraph titled "Implications for the Deal".

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
    deep_dive_sections: Optional[List[str]] = None,
    chunk_prompt: Optional[str] = None,
    aggregate_prompt: Optional[str] = None,
    deep_dive_prompts: Optional[Dict[str, str]] = None
) -> Dict[str, Union[str, int, List[int], Dict]]:
    """
    1) Extract the PDF as a list of pages.
    2) Identify which pages matter for `user_query`.
    3) Concatenate only those pages (or all if none found).
    4) Detect if the user wants BOTH valuation and SWOT (split on “ and ”).
       - If “SWOT” is present alongside something else (e.g. “Valuation”),
         treat them as two separate intents.
    5) If only “SWOT” is requested, do the dedicated SWOT prompt.
    6) If only “red flag” or “risk” is requested, do the single‐shot red‐flag prompt.
    7) Otherwise run a **query‐focused** summarization based on `user_query`.
    8) Optionally run all requested deep‐dives.
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

    lower_q = user_query.lower().strip()

    # ---------------------------------------------------
    # 4) “SWOT + Something Else” DETECTION
    # ---------------------------------------------------
    # If the user typed “Valuation and SWOT Analysis” (or used a comma, etc.),
    # we want to run TWO passes: one for “SWOT” and one for “Valuation”.
    #
    # We’ll split on “ and ” first, then fall back to commas if needed.

    faces: List[str] = []
    if "swot" in lower_q and (" and " in lower_q or "," in lower_q):
        # split on " and " first
        if " and " in lower_q:
            faces = [p.strip() for p in lower_q.split(" and ")]
        else:
            faces = [p.strip() for p in lower_q.split(",")]

        # Now faces might be something like ['valuation', 'swot analysis'] or similar
        # Normalize each piece
        faces = [f for f in faces if f]

        # If one of them is 'swot...' and another is something else, keep both
        swot_requested = any("swot" in f for f in faces)
        other_requested = [f for f in faces if "swot" not in f]
        # We’ll handle each in turn below
    else:
        faces = []
        swot_requested = ("swot" in lower_q)
        other_requested = [lower_q] if not swot_requested else []

    do_valuation_first = bool(other_requested)
    do_swot = swot_requested

    # The final payload we’ll return under “answer”
    answer_payload: Dict[str, Union[Dict, str]] = {}
    valuation_chunks_used: Optional[int] = None
    generic_chunks_used: Optional[int] = None

    # ---------------------------------------------------
    # 5) HANDLE ANY NON‐SWOT INTENT (e.g. “valuation” or “Provide key investment highlights”)
    # ---------------------------------------------------
    if do_valuation_first:
        # We assume other_requested might be something like ['valuation', 'swot analysis']
        # Pick everything that DOESN’T contain “swot”
        non_swot_intents = [f for f in other_requested if "swot" not in f]
        # We’ll join them back into one string (e.g. “valuation” or “valuation analysis”)
        combined_non_swot_query = " and ".join(non_swot_intents).strip()

        # If they asked “risk” or “red flag”
        if "red flag" in combined_non_swot_query or "risk" in combined_non_swot_query:
            prompt = f"""
You are a Private Equity analyst. Identify ALL potential “Risk” or “Red Flag” statements
in the following document excerpt. Return each as a plain English line (no Markdown, no bullets).

Query: {combined_non_swot_query}

--- DOCUMENT EXCERPT START ---
{selected_text}
--- DOCUMENT EXCERPT END ---
""".strip()

            messages = [
                {"role": "system", "content": "You are a knowledgeable private equity investment analyst."},
                {"role": "user",   "content": prompt}
            ]
            print(f"[DEBUG] analyze_transaction_doc: entering single-shot RED FLAG path for '{combined_non_swot_query}'")
            single_shot_response = call_deepseek_chat(messages, temperature=0.0, max_tokens=1000)
            print(f"[DEBUG] analyze_transaction_doc: single_shot_response length {len(single_shot_response)}")
            answer_payload["valuation_analysis"] = single_shot_response

        else:
            # Otherwise, run a **query‐focused** summarization based on user_query
            # 6) CHUNK the selected_text
            raw_chunks = chunk_text(selected_text, max_tokens=chunk_size)
            print(f"[DEBUG] analyze_transaction_doc: total raw_chunks after chunk_text = {len(raw_chunks)}")

            # 6b) If too many raw_chunks (>20), reduce hierarchically until ≤ 20 summaries
            if len(raw_chunks) > 20:
                print(f"[DEBUG] analyze_transaction_doc: {len(raw_chunks)} raw_chunks > 20, reducing hierarchically")
                raw_chunks = reduce_chunks_hierarchically(
                    raw_chunks,
                    max_chunks=20,
                    chunk_size=chunk_size,
                    custom_prompt=chunk_prompt
                )
                print(f"[DEBUG] analyze_transaction_doc: reduced to {len(raw_chunks)} summaries")

            # 6c) Summarize each of the ≤ 20 pieces
            chunk_summaries: List[str] = []
            for idx, chunk in enumerate(raw_chunks, start=1):
                print(f"[DEBUG] analyze_transaction_doc: summarizing chunk {idx}/{len(raw_chunks)}")
                summary = summarize_chunk(chunk, custom_prompt=chunk_prompt)
                chunk_summaries.append(summary)

            # 6d) Now send a **query‐focused** prompt to DeepSeek Chat
            combined_chunks = "\n\n".join(chunk_summaries)
            query_prompt = f"""
You are a seasoned private equity investment analyst. From the following excerpts,
{combined_non_swot_query}. Return your answer as a concise list of bullet‐style points.
Use plain English (no Markdown code fences). Focus ONLY on providing the requested highlights.

--- EXCERPT START ---
{combined_chunks}
--- EXCERPT END ---
""".strip()

            messages = [
                {"role": "system", "content": "You are a knowledgeable private equity investment analyst."},
                {"role": "user",   "content": query_prompt}
            ]
            print(f"[DEBUG] analyze_transaction_doc: sending query‐focused summarization for '{combined_non_swot_query}'")
            query_response = call_deepseek_chat(messages, temperature=0.3, max_tokens=800)
            print(f"[DEBUG] analyze_transaction_doc: query_response length {len(query_response)}")
            answer_payload["query_summary"] = query_response

            # Record how many chunks we used in this “valuation” pass
            valuation_chunks_used = len(raw_chunks)

    # ---------------------------------------------------
    # 7) HANDLE “SWOT” INTENT (if requested)
    # ---------------------------------------------------
    if do_swot:
        print(f"[DEBUG] analyze_transaction_doc: detected 'SWOT' in query; running SWOT path")
        swot_json_str = summarize_swot(selected_text)

        # Strip out triple-backtick fences if present
        cleaned = swot_json_str.strip()
        if cleaned.startswith("```"):
            lines = cleaned.splitlines()
            if lines[0].startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].startswith("```"):
                lines = lines[:-1]
            cleaned = "\n".join(lines).strip()

        try:
            swot_json = json.loads(cleaned)
        except json.JSONDecodeError:
            # If JSON fails, fall back to raw text
            swot_json = {"raw_swot_text": swot_json_str}

        answer_payload["swot_analysis"] = swot_json

    # ---------------------------------------------------
    # 8) IF NEITHER “SWOT” NOR ANY “RISK/RED FLAG” NOR ANY OTHER PASS:
    #     run the same **query‐focused** summarization but for the full user_query
    # ---------------------------------------------------
    if not do_valuation_first and not do_swot:
        print(f"[DEBUG] analyze_transaction_doc: No SWOT or red-flag detected; running query‐focused summarization for '{lower_q}'")

        # Chunk everything
        raw_chunks = chunk_text(selected_text, max_tokens=chunk_size)
        print(f"[DEBUG] analyze_transaction_doc: total raw_chunks after chunk_text = {len(raw_chunks)}")

        if len(raw_chunks) > 20:
            print(f"[DEBUG] analyze_transaction_doc: {len(raw_chunks)} raw_chunks > 20, reducing hierarchically")
            raw_chunks = reduce_chunks_hierarchically(
                raw_chunks,
                max_chunks=20,
                chunk_size=chunk_size,
                custom_prompt=chunk_prompt
            )
            print(f"[DEBUG] analyze_transaction_doc: reduced to {len(raw_chunks)} summaries")

        chunk_summaries: List[str] = []
        for idx, chunk in enumerate(raw_chunks, start=1):
            print(f"[DEBUG] analyze_transaction_doc: summarizing chunk {idx}/{len(raw_chunks)}")
            summary = summarize_chunk(chunk, custom_prompt=chunk_prompt)
            chunk_summaries.append(summary)

        combined_chunks = "\n\n".join(chunk_summaries)
        query_prompt = f"""
You are a seasoned private equity investment analyst. From the following excerpts,
{user_query}. Return your answer as a concise list of bullet‐style points.
Use plain English (no Markdown code fences). Focus ONLY on providing the requested highlights.

--- EXCERPT START ---
{combined_chunks}
--- EXCERPT END ---
""".strip()

        messages = [
            {"role": "system", "content": "You are a knowledgeable private equity investment analyst."},
            {"role": "user",   "content": query_prompt}
        ]
        print(f"[DEBUG] analyze_transaction_doc: sending query‐focused summarization for '{user_query}'")
        query_response = call_deepseek_chat(messages, temperature=0.3, max_tokens=800)
        print(f"[DEBUG] analyze_transaction_doc: query_response length {len(query_response)}")
        answer_payload["query_summary"] = query_response

        generic_chunks_used = len(raw_chunks)

    # ---------------------------------------------------
    # 9) BUILD THE RESULT DICT
    # ---------------------------------------------------

    result: Dict[str, Union[str, int, List[int], Dict]] = {
        "answer": answer_payload,
        "pages_scanned": pages_scanned,
        "relevant_pages": pages_used,
        "chunks_used": valuation_chunks_used if do_valuation_first else generic_chunks_used
    }

    # ---------------------------------------------------
    # 10) (Optional) RUN DEEP DIVES
    # ---------------------------------------------------
    if run_deep_dives:
        if not deep_dive_sections:
            deep_dive_sections = [
                "Market Analysis",
                "Financial Performance",
                "Operational Risks",
                "Exit Strategy"
            ]

        # We need combined summaries for deep dives; reuse whichever chunk_summaries was built
        # (Note: we kept chunk_summaries in each branch above.)
        # If do_valuation_first, chunk_summaries already exists; if do_swot only, use chunk_summaries from last pass; otherwise same.
        print(f"[DEBUG] analyze_transaction_doc: running deep dives on sections {deep_dive_sections}")
        for section in deep_dive_sections:
            section_key = section.replace(" ", "_")
            custom = None
            if deep_dive_prompts and section_key in deep_dive_prompts:
                custom = deep_dive_prompts[section_key]

            # Reassemble combined_chunks accordingly
            combined_for_deep = "\n\n".join(chunk_summaries)
            section_text = deep_dive_section(section, combined_for_deep, custom_prompt=custom)
            result[f"deep_dive_{section_key}"] = section_text

    print(f"[DEBUG] analyze_transaction_doc: completed analysis, returning result")
    return result
