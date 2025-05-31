# deal_analyzer.py

import os

import json
import fitz                   # PyMuPDF for PDFs
import tiktoken               # for token counting
import requests               # for HTTP calls to DeepSeek Chat (LLM)
import pandas as pd           # for Excel extraction
from docx import Document     # for .docx extraction
from pptx import Presentation # for .pptx extraction

# ---------------------------------------------------
# 0) LOAD API KEYS / ENDPOINTS FROM ENVIRONMENT
# ---------------------------------------------------

DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_CHAT_URL = os.getenv(
    "DEEPSEEK_CHAT_URL",
    "https://api.deepseek.ai/v1/chat/completions"
)

if not DEEPSEEK_API_KEY:
    raise RuntimeError("Please set DEEPSEEK_API_KEY in your environment.")


# ---------------------------------------------------
# 1) UTILITIES: token counting & chunking
# ---------------------------------------------------

def num_tokens_from_string(string: str, model_name: str="gpt-4") -> int:
    """
    Returns the number of tokens in `string` when encoded with tiktoken.
    """
    encoding = tiktoken.encoding_for_model(model_name)
    return len(encoding.encode(string))


def chunk_text(text: str, max_tokens: int=1800, model_name: str="gpt-4") -> list[str]:
    """
    Splits `text` into a list of substrings, each containing <= max_tokens tokens.
    Tries to split on paragraph boundaries but guarantees token‐safety by falling back to sentences.
    """
    paragraphs = text.split("\n\n")
    chunks: list[str] = []
    current_chunk = ""
    current_tokens = 0

    for para in paragraphs:
        
        para_tokens = num_tokens_from_string(para, model_name)
        
        if para_tokens > max_tokens:
            # If one paragraph alone is too big, split by sentences
            sentences = para.split(". ")
            for sentence in sentences:
                sent_tokens = num_tokens_from_string(sentence, model_name)
                if current_tokens + sent_tokens + 10 > max_tokens:
                    if current_chunk:
                        chunks.append(current_chunk.strip())
                    current_chunk = sentence + ". "
                    current_tokens = sent_tokens
                else:
                    current_chunk += sentence + ". "
                    current_tokens += sent_tokens
        else:
            # Can we append this paragraph to the current chunk?
            if current_tokens + para_tokens + 20 > max_tokens:
                if current_chunk:
                    chunks.append(current_chunk.strip())
                current_chunk = para + "\n\n"
                current_tokens = para_tokens
            else:
                current_chunk += para + "\n\n"
                current_tokens += para_tokens

    # Flush remainder
    if current_chunk:
        chunks.append(current_chunk.strip())

    return chunks


# ---------------------------------------------------
# 2) TEXT EXTRACTION FOR MULTIPLE FORMATS
# ---------------------------------------------------

def extract_text_from_pdf(filepath: str) -> str:
    """
    Uses PyMuPDF to extract all text from a PDF file.
    """
    doc = fitz.open(filepath)
    all_text = []
    for page in doc:
        all_text.append(page.get_text())
    return "\n\n".join(all_text)


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
        # Join each row's cells with a separator, drop empty rows
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
    messages: list[dict[str,str]],
    model: str="gpt-4",
    temperature: float=0.3,
    max_tokens: int=1500
) -> str:
    """
    Calls DeepSeek Chat (compatible with OpenAI-style ChatCompletion).
    Expects:
      - DEEPSEEK_API_KEY  (Bearer)
      - DEEPSEEK_CHAT_URL
    Returns: The assistant's reply (string).
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
    resp = requests.post(DEEPSEEK_CHAT_URL, headers=headers, json=payload)
    resp.raise_for_status()
    data = resp.json()
    # Assume DeepSeek returns the same structure as OpenAI:
    return data["choices"][0]["message"]["content"]


# ---------------------------------------------------
# 4) LAYERED SUMMARIZATION + ANALYSIS ROUTINES
#    (Now all accept optional custom prompts)
# ---------------------------------------------------

def summarize_chunk(
    chunk_text: str,
    custom_prompt: str | None = None
) -> str:
    """
    Summarize one chunk in a concise, 2–3 bullet‐point paragraph,
    highlighting “Key Insights” and “Potential Risks.”

    If `custom_prompt` is provided, use that entire prompt verbatim,
    inserting the chunk text under a marker. Otherwise use the built-in default.
    """
    if custom_prompt:
        # Insert the chunk text in place of a placeholder token like {{TEXT}}
        prompt = custom_prompt.replace("{{TEXT}}", chunk_text)
    else:
        # Default prompt
        prompt = f"""
You are a Private Equity analyst. Summarize the following excerpt from a deal document in 2–3 bullet points,
highlighting any “Key Insights” and “Potential Risks” in each bullet. Keep each bullet to under 70 words.

--- EXCERPT START ---
{chunk_text}
--- EXCERPT END ---

Respond only with Markdown bullets, each prefaced with “Key Insight:” or “Potential Risk:”.
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
    Given a list of chunk‐level “bullet summaries,” produce a final JSON
    with keys: Executive_Summary, Key_Investment_Insights, Risks_and_Red_Flags,
    High_Level_Valuation_Guidance.

    If `custom_prompt` is provided, use it verbatim, replacing {{SUMMARIES}} with the joined bullet summaries.
    """
    combined_text = "\n\n".join(summaries)

    if custom_prompt:
        prompt = custom_prompt.replace("{{SUMMARIES}}", combined_text)
    else:
        prompt = f"""
You are a top-tier PE investment partner. Below are bullet-point summaries (each from different sections of a deal document).
1) Eliminate redundancies
2) Create a final, cohesive “Executive Summary” (single paragraph)
3) List “Top 5 Key Investment Insights” (as bullets)
4) List “Top 5 Risks and Red Flags” (as bullets)
5) Provide “High-Level Valuation Guidance” in 2–3 bullet points.

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
    Returns a Markdown‐formatted answer with headings and bullets.

    If `custom_prompt` is provided, use it verbatim, replacing:
      - {{SECTION_NAME}} with the section_name
      - {{TEXT}} with combined_text
    """
    if custom_prompt:
        prompt = (custom_prompt
                  .replace("{{SECTION_NAME}}", section_name)
                  .replace("{{TEXT}}", combined_text))
    else:
        prompt = f"""
You are a PE sector specialist. Perform a **detailed {section_name}** analysis on the following deal document extract:

--- DOCUMENT EXCERPT START ---
{combined_text}
--- DOCUMENT EXCERPT END ---

Structure your answer like this:
1) <Section Header>  (e.g., Market Overview)
2) A short 2-sentence summary
3) A bullet list of 4–6 observations or considerations
4) A final “Implication for Deal” line (under 50 words)

Respond only in Markdown, with headings and bullet points.
""".strip()

    messages = [
        {"role": "system", "content": "You are a domain expert in private equity deals."},
        {"role": "user",   "content": prompt}
    ]
    return call_deepseek_chat(messages, temperature=0.25, max_tokens=1000)


# ---------------------------------------------------
# 5) MAIN “ANALYZE TRANSACTION” WORKFLOW
# ---------------------------------------------------

def analyze_transaction_doc(
    filepath: str,
    run_deep_dives: bool = True,
    chunk_size: int = 1800,
    deep_dive_sections: list[str] | None = None,
    chunk_prompt: str | None = None,
    aggregate_prompt: str | None = None,
    deep_dive_prompts: dict[str,str] | None = None
) -> dict:
    """
    1) Extract & chunk the document (PDF, DOCX, PPTX, or Excel)
    2) Summarize each chunk (Layer 1) using `chunk_prompt` or default
    3) Aggregate chunk summaries (Layer 2) using `aggregate_prompt` or default
    4) (Optional) Run deep dives on `deep_dive_sections` (Layer 3) using per-section prompts if provided
    5) Return a dict containing:
         - "aggregate_analysis": parsed JSON
         - "deep_dive_<SectionKey>" (Markdown) for each requested section
    """

    # 1) Extract raw text
    raw_text = extract_text(filepath)

    # 2) Chunk up to ~chunk_size tokens each
    chunks = chunk_text(raw_text, max_tokens=chunk_size, model_name="gpt-4")
    print(f">>> Document split into {len(chunks)} chunks (chunk_size={chunk_size})")

    # 3) Summarize each chunk (Layer 1)
    chunk_summaries: list[str] = []
    for idx, chunk in enumerate(chunks):
        print(f"Summarizing chunk {idx+1}/{len(chunks)}...")
        summary = summarize_chunk(chunk, custom_prompt=chunk_prompt)
        chunk_summaries.append(summary)

    # 4) Aggregate chunk summaries (Layer 2)
    print("Aggregating summaries into final JSON...")
    aggregate_json_str = aggregate_summaries(chunk_summaries, custom_prompt=aggregate_prompt)
    try:
        aggregate_json = json.loads(aggregate_json_str)
    except json.JSONDecodeError:
        aggregate_json = {"raw_aggregate_text": aggregate_json_str}

    result: dict[str, object] = {"aggregate_analysis": aggregate_json}

    # 5) Deep dives (Layer 3), if requested
    if run_deep_dives:
        # If no specific deep_dive_sections passed, default back to the four sections
        if not deep_dive_sections:
            deep_dive_sections = [
                "Market Analysis",
                "Financial Performance",
                "Operational Risks",
                "Exit Strategy"
            ]

        # Combine all chunk summaries for passing to each deep dive
        combined_summaries_text = "\n\n".join(chunk_summaries)

        for section in deep_dive_sections:
            section_key = section.replace(" ", "_")  # e.g. "Market Analysis" → "Market_Analysis"
            # See if a custom prompt was provided for this section
            custom = None
            if deep_dive_prompts and section_key in deep_dive_prompts:
                custom = deep_dive_prompts[section_key]

            print(f"Running deep dive on {section} (using custom={bool(custom)})...")
            section_md = deep_dive_section(section, combined_summaries_text, custom_prompt=custom)
            result[f"deep_dive_{section_key}"] = section_md

    return result
