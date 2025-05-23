import os
import re
from  openai import OpenAI
import PyPDF2
from docx import Document
import concurrent.futures


# Read your DeepSeek key (and optional base URL) from env
DEEPSEEK_API_KEY = os.environ["DEEPSEEK_API_KEY"]
DEEPSEEK_API_BASE = os.getenv("DEEPSEEK_API_BASE", "https://api.deepseek.com")

# Create the DeepSeek client
client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_API_BASE)

# Function to extract text from a PDF
def extract_pdf_text(file_path):
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
    return text

# Function to generate a title using DeepSeek's chat model
def generate_title(text):
    truncated_text = text[:2000]
    messages = [
        {"role": "system", "content": "You are a skilled assistant that generates concise titles for documents."},
        {"role": "user", "content": f"Based on the following content, generate a concise title that accurately reflects the main themes of the document:\n\n{truncated_text}"}
    ]
    response = client.chat.completions.create(
        model="deepseek-chat",  # DeepSeek's model identifier
        messages=messages,
        max_tokens=50,
        temperature=0.7
    )
    return response.choices[0].message.content.strip()

# Function to summarize text using DeepSeek's chat model
def summarize_text(text, target_word_count=100):
    messages = [
        {"role": "system", "content": "You are a highly skilled assistant who condenses information to only the most important details, keeping the summary under the target word count."},
        {"role": "user", "content": f"Please summarize the following text, focusing on the most relevant points, and aim to keep it under {target_word_count} words:\n\n{text}"}
    ]
    response = client.chat.completions.create(
        model="deepseek-chat",  # DeepSeek's model identifier
        messages=messages,
        max_tokens=target_word_count * 2,
        temperature=0.7
    )
    return response.choices[0].message.content.strip()

def summarize_chunks_concurrently(chunks, target_word_count=100):
    with concurrent.futures.ThreadPoolExecutor() as executor:
        # Schedule a summary for each chunk concurrently
        futures = [executor.submit(summarize_text, chunk, target_word_count) for chunk in chunks]
        summaries = [future.result() for future in concurrent.futures.as_completed(futures)]
    return summaries

# Function to split text into smaller chunks
def split_text(text, max_chunk_size=1000):
    chunks = []
    while len(text) > max_chunk_size:
        chunk = text[:max_chunk_size]
        chunks.append(chunk)
        text = text[max_chunk_size:]
    if text:
        chunks.append(text)
    return chunks

# Iterative summarization: Group summaries into batches and summarize each batch until one final summary remains.
def iterative_refine_summary(summaries, target_word_count=250, batch_size=5):
    # Base case: If only one summary remains, return it.
    if len(summaries) == 1:
        return summaries[0]
    new_summaries = []
    # Process summaries in batches.
    for i in range(0, len(summaries), batch_size):
        batch = " ".join(summaries[i:i+batch_size])
        refined = summarize_text(batch, target_word_count=target_word_count)
        new_summaries.append(refined)
    # If more than one summary remains, iterate again.
    return iterative_refine_summary(new_summaries, target_word_count=target_word_count, batch_size=batch_size)

# Function to sanitize title and create a Word document from the summary with a dynamic title
def create_word_document(summary, title, docs_folder="documents"):
    

    sanitized_title = re.sub(r'[<>:"/\\|?*]', '_', title)
    sanitized_title = re.sub(r"'", "_", sanitized_title)
    sanitized_title = sanitized_title.strip() or "Untitled_Document"
    sanitized_title = sanitized_title[:255]  # max filename length

    # Use the Flask `DOCS_FOLDER` via env variable or fallback
    os.makedirs(docs_folder, exist_ok=True)

    full_path = os.path.join(docs_folder, f"{sanitized_title}.docx")

    doc = Document()
    doc.add_heading(sanitized_title, 0)
    doc.add_paragraph('This document summarizes the key insights from the uploaded document.\n')
    doc.add_paragraph(summary)
    doc.save(full_path)

    return full_path  # ✅ return actual path



# Main function to process the PDF
def main():
    pdf_file = input("Please upload a PDF file path for summarization: ")
    extracted_text = extract_pdf_text(pdf_file)
    
    if extracted_text:
        print(f"Extracted text length: {len(extracted_text)}")
        # Generate a custom title using DeepSeek based on the content
        custom_title = generate_title(extracted_text)
        print(f"Generated title: {custom_title}")
        
        # Split the document into smaller chunks
        text_chunks = split_text(extracted_text, max_chunk_size=1000)
        print(f"Number of chunks: {len(text_chunks)}")
        
        # Summarize each chunk concurrently
        summaries = summarize_chunks_concurrently(text_chunks, target_word_count=100)
        print(f"Number of individual summaries: {len(summaries)}")
        
        # Iteratively refine the individual summaries into one final summary (~250 words)
        final_summary = iterative_refine_summary(summaries, target_word_count=250, batch_size=5)
        print(f"Final summary length: {len(final_summary)} characters")
        
        # Create and save the Word document with the dynamic title
        create_word_document(final_summary, custom_title)
        print(f"\nSummary has been saved as '{custom_title}.docx' with 250 words or less!")
    else:
        print("Failed to extract text from the PDF.")

if __name__ == "__main__":
    main()
