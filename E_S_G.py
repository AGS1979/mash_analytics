import fitz  # PyMuPDF for PDF extraction
import requests
import json
import re
import os
from datetime import datetime

# Try to import OCR libraries if available
try:
    import pytesseract
    from PIL import Image
    import io
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# DeepSeek API Settings from environment
DEEPSEEK_API_KEY = os.environ["DEEPSEEK_API_KEY"]
DEEPSEEK_API_URL = os.getenv(
    "DEEPSEEK_API_URL",
    "https://api.deepseek.com/v1/chat/completions"
)

# Logging setup
def setup_logging():
    log_dir = "esg_analysis_logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    log_file = os.path.join(log_dir, f"esg_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    return log_file

LOG_FILE = setup_logging()

def log_message(message):
    """Log messages to file and print to console"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}\n"
    print(message)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(log_entry)

def extract_text_from_pdf(pdf_path):
    """Enhanced PDF text extraction with better error handling"""
    log_message(f"📂 Loading PDF: {pdf_path}")
    try:
        doc = fitz.open(pdf_path)
        text = []
        max_pages = 250  # Limit for very large documents
        
        for page_num, page in enumerate(doc):
            if page_num >= max_pages:
                break
            try:
                page_text = page.get_text("text")
                if page_text.strip():
                    text.append(page_text)
            except Exception as e:
                log_message(f"⚠️ Error reading page {page_num + 1}: {e}")
        
        full_text = "\n\n".join(text)
        if not full_text.strip():
            log_message("❌ Warning: No text found in PDF. Is this a scanned document?")
        return full_text
    except Exception as e:
        log_message(f"❌ Error reading PDF file: {e}")
        return ""

def analyze_esg_with_deepseek(text):
    """Improved DeepSeek analysis with better prompting and error handling"""
    if not text.strip():
        log_message("❌ Error: Cannot send empty text to DeepSeek API!")
        return "DeepSeek API Error: No text provided."
    
    log_message("🔄 Sending request to DeepSeek API for ESG analysis...")
    
    prompt = f"""
You are an expert ESG analyst. Carefully read the following ESG disclosure and generate a detailed analysis. Be specific and data-driven.

Provide the analysis in these sections:

1. 🌍 **Environmental (E)**:
   - Give **10 detailed insights** about energy use, emissions, renewable energy adoption, waste reduction, water conservation, climate initiatives, biodiversity actions, etc.
   - Use **quantitative data**, clear targets, and named programs or initiatives.
   - Mention **year-over-year improvements** or regressions if applicable.
   - Avoid vague statements; elaborate where necessary.

2. 🏢 **Social (S)**:
   - Give **10 detailed insights** covering labor practices, diversity & inclusion, community engagement, training programs, health & safety, etc.
   - Include **figures**, **employee stats**, and **notable case studies** if present.
   - Highlight notable changes over time and any certifications or recognitions.

3. 🏛 **Governance (G)**:
   - Provide **10 robust insights** on board structure, executive compensation, risk management, ethics programs, whistleblower mechanisms, and audit independence.
   - Use **board diversity numbers**, policy names, or governance frameworks where mentioned.

4. 🎤 **Key Management Remarks**:
   - Extract **5–10 strong quotes** from executive leadership, especially forward-looking or strategic statements.
   - Attribute each quote to a named executive or title if mentioned.

5. 🎯 **ESG Sentiment Score**:
   - Rate from 1–10 (10 = exceptional ESG commitment and execution).
   - Justify score briefly in 1–2 lines by considering specificity, tone, and depth of ESG strategy.

Return only the output in this structured format:
    ```
    Environmental:
    1. Insight 1...
    2. Insight 2...
    ...
    10. Insight 10...

    Social:
    1. Insight 1...
    ...
    10. Insight 10...

    Governance:
    1. Insight 1...
    ...
    10. Insight 1...

    Key Remarks:
    1. "Quote 1..." - [Title]
    2. "Quote 2..." - [Title]
    ...

    ESG Sentiment Score: X/10
    ```

    DOCUMENT TEXT:
    {text[:10000]}  
    """
    
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.5,
        "max_tokens": 3000
    }
    
    try:
        response = requests.post(DEEPSEEK_API_URL, headers=headers, json=payload, timeout=60)
        if response.status_code != 200:
            log_message(f"❌ API Error: {response.status_code}, Response: {response.text}")
            return f"DeepSeek API Error: {response.status_code}"
        
        response_data = response.json()
        if "choices" in response_data:
            result = response_data["choices"][0]["message"]["content"]
            log_message("📥 Successfully received analysis from DeepSeek")
            return result
        else:
            log_message("❌ Unexpected API response format")
            return "DeepSeek API Error: No insights generated."
    except Exception as e:
        log_message(f"❌ DeepSeek API Request Failed: {e}")
        return f"DeepSeek API Error: {str(e)}"

def parse_esg_data(api_response):
    """Enhanced parsing with better error handling"""
    esg_data = {
        "environment": [], 
        "social": [], 
        "governance": [], 
        "management_remarks": [], 
        "sentiment_score": "N/A"
    }
    
    try:
        # Extract Environmental insights
        env_match = re.search(r'Environmental:\s*(.*?)(?=\n\s*Social:|$)', api_response, re.DOTALL)
        if env_match:
            env_insights = [i.strip() for i in env_match.group(1).split('\n') if i.strip()]
            esg_data["environment"] = [re.sub(r'^\d+\.\s*', '', i) for i in env_insights[:10]]
        
        # Extract Social insights
        soc_match = re.search(r'Social:\s*(.*?)(?=\n\s*Governance:|$)', api_response, re.DOTALL)
        if soc_match:
            soc_insights = [i.strip() for i in soc_match.group(1).split('\n') if i.strip()]
            esg_data["social"] = [re.sub(r'^\d+\.\s*', '', i) for i in soc_insights[:10]]
        
        # Extract Governance insights
        gov_match = re.search(r'Governance:\s*(.*?)(?=\n\s*Key Remarks:|$)', api_response, re.DOTALL)
        if gov_match:
            gov_insights = [i.strip() for i in gov_match.group(1).split('\n') if i.strip()]
            esg_data["governance"] = [re.sub(r'^\d+\.\s*', '', i) for i in gov_insights[:10]]
        
        # Extract Management Remarks
        mgmt_match = re.search(r'Key Remarks:\s*(.*?)(?=\n\s*ESG Sentiment Score:|$)', api_response, re.DOTALL)
        if mgmt_match:
            remarks = [i.strip() for i in mgmt_match.group(1).split('\n') if i.strip()]
            esg_data["management_remarks"] = [re.sub(r'^\d+\.\s*', '', i) for i in remarks[:10]]
        
        # Extract Sentiment Score
        sentiment_match = re.search(r'ESG Sentiment Score:\s*(\d+\.?\d*)\s*/\s*10', api_response)
        if sentiment_match:
            esg_data["sentiment_score"] = sentiment_match.group(1)
    
    except Exception as e:
        log_message(f"⚠️ Error parsing ESG data: {e}")
    
    return esg_data

def generate_html_report(esg_data, company_name):
    """
    Creates an interactive HTML report with company name only
    Report name: ESG_Insights_<Company Name>.html
    Report title: <Company Name> ESG Insights Report
    """
    # Clean company name for filename
    safe_company_name = re.sub(r'[^\w\-_]', '_', company_name)[:50]
    output_file = f"ESG_Insights_{safe_company_name}.html"
    
    # Format current date
    current_date = datetime.now().strftime("%B %d, %Y")
    
    def generate_section(title, icon, insights):
        if not insights:
            return ""
        section_html = f"""
            <h2><span class="category-icon">{icon}</span>{title}</h2>
            <table>
                <thead>
                    <tr>
                        <th width="5%">#</th>
                        <th>Insight</th>
                    </tr>
                </thead>
                <tbody>
        """
        for idx, insight in enumerate(insights, 1):
            section_html += f"""
                    <tr>
                        <td>{idx}</td>
                        <td>{insight}</td>
                    </tr>
            """
        section_html += """
                </tbody>
            </table>
        """
        return section_html
    
    # Build the complete HTML content
    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{company_name} ESG Insights Report</title>
        <style>
            body {{ 
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                line-height: 1.6;
                color: #333;
                background-color: #f9f9f9;
                padding: 0;
                margin: 0;
            }}
            .container {{
                max-width: 1000px;
                margin: 20px auto;
                background: white;
                padding: 30px;
                border-radius: 8px;
                box-shadow: 0 0 20px rgba(0,0,0,0.1);
            }}
            header {{
                border-bottom: 2px solid #2196F3;
                padding-bottom: 20px;
                margin-bottom: 30px;
            }}
            h1, h2, h3 {{
                color: #2c3e50;
            }}
            h1 {{
                margin-top: 0;
                font-size: 2.2em;
            }}
            h2 {{
                border-bottom: 1px solid #eee;
                padding-bottom: 8px;
                margin-top: 30px;
                font-size: 1.5em;
                color: #2196F3;
            }}
            h3.subtitle {{
                color: #7f8c8d;
                font-weight: normal;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin: 20px 0;
                font-size: 0.95em;
            }}
            th, td {{
                border: 1px solid #ddd;
                padding: 12px 15px;
                text-align: left;
            }}
            th {{
                background-color: #2196F3;
                color: white;
                font-weight: bold;
            }}
            tr:nth-child(even) {{
                background-color: #f2f2f2;
            }}
            tr:hover {{
                background-color: #e3f2fd;
            }}
            .sentiment {{
                font-size: 1.2em;
                padding: 10px 15px;
                background-color: #e8f5e9;
                border-radius: 4px;
                display: inline-block;
                margin: 10px 0;
            }}
            footer {{
                margin-top: 40px;
                text-align: center;
                font-size: 0.9em;
                color: #7f8c8d;
                border-top: 1px solid #eee;
                padding-top: 20px;
            }}
            .category-icon {{
                font-size: 1.2em;
                margin-right: 8px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <header>
                <h1>{company_name} ESG Insights Report</h1>
                <h3 class="subtitle">Generated on: {current_date}</h3>
                <div class="sentiment">
                    <strong>ESG Sentiment Score:</strong> {esg_data['sentiment_score']}/10
                </div>
            </header>
    """
    
    # Add sections
    html_content += generate_section("Environmental Insights", "🌍", esg_data["environment"])
    html_content += generate_section("Social Insights", "🏢", esg_data["social"])
    html_content += generate_section("Governance Insights", "🏛", esg_data["governance"])
    
    # Add management remarks if available
    if esg_data["management_remarks"]:
        html_content += """
            <h2>🎤 Key Remarks by Management</h2>
            <table>
                <thead>
                    <tr>
                        <th width="5%">#</th>
                        <th>Remark</th>
                    </tr>
                </thead>
                <tbody>
        """
        for idx, remark in enumerate(esg_data["management_remarks"], 1):
            html_content += f"""
                    <tr>
                        <td>{idx}</td>
                        <td>{remark}</td>
                    </tr>
            """
        html_content += """
                </tbody>
            </table>
        """
    
    # Footer
    html_content += f"""
            <footer>
                ESG Report Analysis generated on {current_date}<br>
                </footer>
        </div>
    </body>
    </html>
    """
    
    # Save file
    with open(output_file, "w", encoding="utf-8") as file:
        file.write(html_content)
    
    log_message(f"✅ ESG Report generated: {output_file}")
    return output_file

# In E_S_G.py

def generate_esg_report(pdf_path, company_name):
    """
    Main function to generate ESG report with enhanced error handling
    Returns the output HTML file path on success, or None on failure.
    """
    try:
        log_message(f"\n{'='*50}")
        log_message(f"Starting ESG analysis for: {pdf_path}")

        # Step 1: extract text…
        pdf_text = extract_text_from_pdf(pdf_path)
        if not pdf_text.strip():
            log_message("❌ Error: No text extracted from PDF")
            return None

        # Step 2: call DeepSeek…
        esg_analysis = analyze_esg_with_deepseek(pdf_text)
        if esg_analysis.startswith("DeepSeek API Error"):
            log_message(f"❌ Analysis failed: {esg_analysis}")
            return None

        # Step 3: parse…
        esg_data = parse_esg_data(esg_analysis)

        # Step 4: generate HTML
        report_file = generate_html_report(esg_data, company_name)

        log_message(f"\n🎉 Successfully generated report: {report_file}")
        log_message(f"{'='*50}\n")

        # **RETURN the path** so the caller can move or serve it
        return report_file

    except Exception as e:
        log_message(f"❌ Fatal error in report generation: {e}")
        return None

if __name__ == "__main__":
    print("📊 ESG Report Analysis Tool")
    print("="*50)
    pdf_file = input("📂 Enter the PDF file name (or path): ").strip()
    
    if not os.path.exists(pdf_file):
        print(f"❌ Error: File '{pdf_file}' not found!")
        exit(1)
    
    company_name = input("🏢 Enter the company name: ").strip()
    if not company_name:
        print("❌ Error: Company name cannot be empty!")
        exit(1)
    
    success = generate_esg_report(pdf_file, company_name)
    
    if success:
        print("\n✅ Report generated successfully! Check the output HTML file.")
    else:
        print("\n❌ Report generation failed. Check the log file for details.")
    
    print(f"\nLog file: {LOG_FILE}")