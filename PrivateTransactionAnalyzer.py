import fitz  # PyMuPDF
import os
import requests
import tiktoken
import math
import json

# ---------------------------------------------------
# 1) UTILITIES: token counting & chunking
# ---------------------------------------------------

def num_tokens_from_string(string: str, model_name: str="gpt-4") -> int:
    """
    Returns the number of tokens in `string` when encoded with the given model.
    Uses tiktoken (OpenAI).
    """
    encoding = tiktoken.encoding_for_model(model_name)
    return len(encoding.encode(string))

def chunk_text(text: str, max_tokens: int=1800, model_name: str="gpt-4"):
    """
    Splits `text` into a list of substrings, each containing <= max_tokens tokens.
    Tries to split on paragraph boundaries but guarantees token‐safety.
    """
    paragraphs = text.split("\n\n")
    chunks = []
    current_chunk = ""
    current_tokens = 0

    for para in paragraphs:
        para_tokens = num_tokens_from_string(para, model_name)
        # If single paragraph too big, we forcibly split it by sentences (fallback)
        if para_tokens > max_tokens:
            # break long para into sentences and note tokens
            sentences = para.split(". ")
            for sentence in sentences:
                sent_tokens = num_tokens_from_string(sentence, model_name)
                if current_tokens + sent_tokens + 10 > max_tokens:  # +10 buffer
                    if current_chunk:
                        chunks.append(current_chunk.strip())
                    current_chunk = sentence + ". "
                    current_tokens = sent_tokens
                else:
                    current_chunk += sentence + ". "
                    current_tokens += sent_tokens
        else:
            # can we add this whole paragraph?
            if current_tokens + para_tokens + 20 > max_tokens:  # +20 tokens buffer for prompt overhead
                if current_chunk:
                    chunks.append(current_chunk.strip())
                current_chunk = para + "\n\n"
                current_tokens = para_tokens
            else:
                current_chunk += para + "\n\n"
                current_tokens += para_tokens

    # final flush
    if current_chunk:
        chunks.append(current_chunk.strip())

    return chunks

# ---------------------------------------------------
# 2) TEXT EXTRACTION
# ---------------------------------------------------

def extract_text_from_pdf(filepath: str) -> str:
    """
    Uses PyMuPDF to extract all text from the PDF.
    """
    doc = fitz.open(filepath)
    all_text = []
    for page in doc:
        all_text.append(page.get_text())
    return "\n\n".join(all_text)

# ---------------------------------------------------
# 3) CHAT‐COMPLETION HELPERS
# ---------------------------------------------------

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_URL = "https://api.openai.com/v1/chat/completions"

def call_openai_chat(messages: list, model: str="gpt-4", temperature: float=0.3, max_tokens: int=1500):
    """
    Performs a single ChatCompletion call to OpenAI.
    Returns the assistant's reply string.
    """
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens
    }
    resp = requests.post(OPENAI_URL, headers=headers, json=payload)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]

# ---------------------------------------------------
# 4) LAYERED SUMMARIZATION + ANALYSIS ROUTINES
# ---------------------------------------------------

def summarize_chunk(chunk_text: str) -> str:
    """
    Summarize one chunk in a concise, 200–300 token paragraph.
    """
    prompt = f"""
You are a Private Equity analyst. Summarize the following excerpt from a deal document in 2–3 bullet points,
highlighting any “Key Insights” and “Potential Risks” in each bullet. Keep each bullet to < 70 words.

--- EXCERPT START ---
{chunk_text}
--- EXCERPT END ---

Respond only with Markdown bullets, prefaced with “Key Insight:” or “Potential Risk:”.
"""
    messages = [
        {"role": "system", "content": "You are a knowledgeable private equity investment analyst."},
        {"role": "user", "content": prompt}
    ]
    return call_openai_chat(messages)

def aggregate_summaries(summaries: list) -> str:
    """
    Given a list of chunk‐level summaries, produce a global summary that fuses redundancies and structures.
    Return a cohesive 3–4 paragraph analysis under headings:
      - Executive Summary
      - Combined Key Insights
      - Combined Risks & Red Flags
      - High‐Level Valuation Thoughts
    """
    combined_text = "\n\n".join(summaries)
    prompt = f"""
You are a top‐tier PE investment partner. Below are bullet‐point summaries (each from different sections of a deal document).
1) Eliminate redundancies
2) Create a final, cohesive “Executive Summary” (1 paragraph)
3) List “Top 5 Key Investment Insights” (as bullets)
4) List “Top 5 Risks and Red Flags” (as bullets)
5) Provide “High‐Level Valuation Guidance” in 2–3 bullet points.

--- BULLET SUMMARIES START ---
{combined_text}
--- BULLET SUMMARIES END ---

Respond in JSON format, for example:
{{
  "Executive_Summary": "...",
  "Key_Investment_Insights": [
     "...", "..."
  ],
  "Risks_and_Red_Flags": [
     "...", "..."
  ],
  "High_Level_Valuation_Guidance": [
     "...", "..."
  ]
}}
"""
    messages = [
        {"role": "system", "content": "You are a senior private equity partner, extremely concise and factual."},
        {"role": "user", "content": prompt}
    ]
    return call_openai_chat(messages)

def deep_dive_section(section_name: str, combined_text: str) -> str:
    """
    Runs a specialized analysis on the combined text focused on a particular area:
      - section_name could be "Market Analysis", "Competitive Landscape", "Financial Performance",
        "Operational Risks", "Exit Strategies" etc.
    Returns a detailed bullet‐list or short JSON for that section.
    """
    prompt = f"""
You are a PE sector specialist. Perform a **detailed {section_name}** analysis on the following deal document extract:
--- DOCUMENT EXCERPT START ---
{combined_text}
--- DOCUMENT EXCERPT END ---

Structure your answer as follows:
1) <Section Header>  (e.g., Market Overview)
2) A short 2‐sentence summary
3) A bullet list of 4–6 observations or considerations
4) A final “Implication for Deal” line (< 50 words)

Respond only in Markdown, with headings and bullet points.
"""
    messages = [
        {"role": "system", "content": "You are a domain expert in private equity deals."},
        {"role": "user", "content": prompt}
    ]
    return call_openai_chat(messages, temperature=0.25, max_tokens=1000)

# ---------------------------------------------------
# 5) MAIN “ANALYZE TRANSACTION” WORKFLOW
# ---------------------------------------------------

def analyze_transaction_doc(filepath: str, run_deep_dives: bool=True) -> dict:
    """
    1) Extract & chunk the PDF
    2) Summarize each chunk (LAYER 1)
    3) Aggregate chunk summaries into a master JSON (LAYER 2)
    4) (Optional) Run deep‐dives into specific sections
    5) Return a comprehensive dict containing:
         - aggregated JSON
         - (if run_deep_dives=True) extra sections
    """
    raw_text = extract_text_from_pdf(filepath)
    # 1) Chunk up to ~1800 tokens each
    chunks = chunk_text(raw_text, max_tokens=1800, model_name="gpt-4")
    print(f">>> PDF split into {len(chunks)} chunks")

    # 2) Summarize each chunk
    chunk_summaries = []
    for idx, chunk in enumerate(chunks):
        print(f"Summarizing chunk {idx+1}/{len(chunks)}...")
        summary = summarize_chunk(chunk)
        chunk_summaries.append(summary)

    # 3) Aggregate chunk-level summaries
    print("Aggregating summaries into final JSON...")
    aggregate_json_str = aggregate_summaries(chunk_summaries)
    try:
        aggregate_json = json.loads(aggregate_json_str)
    except json.JSONDecodeError:
        # Fallback: return raw text if JSON fails
        aggregate_json = {
            "raw_aggregate_text": aggregate_json_str
        }

    result = {"aggregate_analysis": aggregate_json}

    # 4) (Optional) Deep dives on key areas
    if run_deep_dives:
        # We can pass the entire raw text or the concatenated chunk_summaries
        combined_summaries_text = "\n\n".join(chunk_summaries)
        for section in ["Market Analysis", "Financial Performance", "Operational Risks", "Exit Strategy"]:
            print(f"Running deep dive on {section}...")
            section_output = deep_dive_section(section, combined_summaries_text)
            result[f"deep_dive_{section.replace(' ', '_')}"] = section_output

    return result
