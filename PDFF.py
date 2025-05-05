import os
from openai import OpenAI
import pandas as pd
import pdfkit
import re
import os
import tempfile
import zipfile


# Initialize the DeepSeek client using environment variables
DEEPSEEK_API_KEY = os.environ["DEEPSEEK_API_KEY"]
DEEPSEEK_API_URL = os.getenv("DEEPSEEK_API_URL", "https://api.deepseek.com")
client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_API_URL)

def get_ticker(company_name):
    messages = [
        {
            "role": "system",
            "content": "You are a helpful assistant that provides ticker symbols. Return only the ticker symbol in uppercase letters."
        },
        {
            "role": "user",
            "content": f"Provide the ticker symbol for the following company: {company_name}."
        }
    ]
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=messages,
            stream=False,
            max_tokens=10,
            temperature=1.3
        )
        result = response.choices[0].message.content.strip()
        match = re.search(r'\b([A-Z]{1,5})\b', result)
        if match:
            return match.group(1)
        else:
            return result
    except Exception as e:
        print("Error retrieving ticker from DeepSeek:", e)
        return None

def download_pdf(url, ticker, year):
    output_filename = f"{ticker}_10K_{year}.pdf"
    path_wkhtmltopdf = r'C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
    try:
        pdfkit.from_url(url, output_filename, configuration=config)
        print(f"Downloaded PDF for {year}: {output_filename}")
    except Exception as e:
        print(f"Error converting the {year} filing from URL to PDF: {e}")

def generate_10k_filings(company_name, years):
    """
    Generates 10-K filings for the given company and years.
    Returns the path to a single PDF file or a ZIP archive containing multiple PDFs.
    """
    ticker = get_ticker(company_name)
    if not ticker:
        raise Exception("Ticker could not be identified.")

    csv_file = "10k_links_output_Batch1.csv"
    try:
        df = pd.read_csv(csv_file)
    except Exception as e:
        raise Exception(f"Error reading CSV file '{csv_file}': {e}")

    expected_columns = ['Ticker', '2024', '2023', '2022', '2021', '2020']
    if not all(col in df.columns for col in expected_columns):
        raise Exception(f"CSV file does not contain expected columns: {expected_columns}")

    matched_rows = df[df['Ticker'].str.upper() == ticker.upper()]
    if matched_rows.empty:
        raise Exception("Ticker not found in CSV file.")
    row = matched_rows.iloc[0]

    temp_dir = tempfile.mkdtemp()
    pdf_files = []

    for year in years:
        if year not in df.columns:
            print(f"Year {year} not found in CSV columns.")
            continue
        url = row[year]
        if pd.isna(url) or not url.strip():
            print(f"No URL available for {year}.")
            continue

        download_pdf(url, ticker, year)
        pdf_path = os.path.join(os.getcwd(), f"{ticker}_10K_{year}.pdf")
        if os.path.exists(pdf_path):
            new_path = os.path.join(temp_dir, f"{ticker}_10K_{year}.pdf")
            os.rename(pdf_path, new_path)
            pdf_files.append(new_path)
        else:
            print(f"PDF for {year} was not generated.")

    if not pdf_files:
        raise Exception("No PDFs could be generated.")

    if len(pdf_files) == 1:
        return pdf_files[0]
    else:
        zip_path = os.path.join(temp_dir, f"{ticker}_10K_filings.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for pdf_file in pdf_files:
                zipf.write(pdf_file, os.path.basename(pdf_file))
        return zip_path

# For standalone running (if needed)
def main():
    company_name = input("Enter the company name: ").strip()
    years_input = input("Enter the year or years (comma-separated) between 2020 and 2024: ").strip()
    years = [year.strip() for year in years_input.split(',')]
    output = generate_10k_filings(company_name, years)
    print("Output file:", output)

if __name__ == "__main__":
    main()
