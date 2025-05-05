import openai
import PyPDF2
import os

# Read your OpenAI API key from the environment; throws if not set
openai.api_key = os.environ["OPENAI_API_KEY"]

# Function to extract text from a PDF
def extract_pdf_text(file_path):
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
    return text

# Function to summarize text using OpenAI's chat model (correct endpoint)
def summarize_text(text):
    messages = [{"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": f"Summarize the following document into a concise one-page summary:\n\n{text}"}]
    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=messages,
        max_tokens=1000,
        temperature=0.7
    )
    return response['choices'][0]['message']['content'].strip()

# Function to split text into smaller chunks
def split_text(text, max_chunk_size=2000):
    chunks = []
    while len(text) > max_chunk_size:
        chunk = text[:max_chunk_size]
        chunks.append(chunk)
        text = text[max_chunk_size:]
    if text:
        chunks.append(text)
    return chunks

# Ask user to upload a PDF file
def main():
    pdf_file = input("Please upload a PDF file path for summarization: ")
    extracted_text = extract_pdf_text(pdf_file)
    
    if extracted_text:
        # Split the document into smaller chunks
        text_chunks = split_text(extracted_text)
        
        # Summarize each chunk
        summaries = []
        for chunk in text_chunks:
            summary = summarize_text(chunk)
            summaries.append(summary)
        
        # Concatenate all summaries into one final summary
        final_summary = "\n".join(summaries)
        print("\nSummary of the document:\n")
        print(final_summary)
    else:
        print("Failed to extract text from the PDF.")

if __name__ == "__main__":
    main()
