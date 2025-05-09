from StockReport import get_fmp_json
from dotenv import load_dotenv
load_dotenv()   # reads .env into os.environ
import tempfile
import zipfile
import pandas as pd
import traceback  # Add this at the top
import os
from flask import Flask, request, jsonify, render_template, send_from_directory, url_for, session, redirect, send_file
from StockReport import create_stock_report, process_query_1
from EarningsCallSummaries import main, process_query_2
from flask_cors import CORS
from MktPerf import (
generate_full_market_performance,
generate_emerging_market_performance,
generate_developed_market_performance,
process_query_3
)
import sys
from datetime import datetime, timedelta
from NewsUpdates import get_latest_news_summary  # Import the new module
from EarningsCallExtract import get_earnings_call_data
from DocumentShortSummary import (
    extract_pdf_text, generate_title, split_text,
    summarize_text, iterative_refine_summary, create_word_document
)
from DocumentLongSummary import extract_pdf_text as extract_pdf_text_long, summarize_text as summarize_text_long, split_text as split_text_long
#import ESG
import openpyxl
import json
from TranscriptSummary import generate_earnings_call_summary
from PDFF import get_ticker, download_pdf, generate_10k_filings
import E_S_G
import ESGComp
import time
import shutil
import json

from werkzeug.utils import secure_filename
from mgmt_evasiveness import (
    resolve_ticker,
    fetch_transcript,
    get_price,
    parse_management_turns,
    analyze_evasiveness,
    generate_evasiveness_report,
    merge_reports, extract_evasiveness_parameters_from_text
)
from DebtCov import analyze_debt_covenants


# Define your email whitelist here
WHITELISTED_EMAILS = {
    "avinashg.singh@aranca.com",
    "ujjal.roy@aranca.com",
    "rohit.dhawan@aranca.com",
    "avi104@yahoo.co.in",
    "vishal.kumar@aranca.com"
}

# Add the current directory to the sys.path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

app = Flask(__name__, static_folder='static')
CORS(app)

# Set a secret key for session management (use a secure random key in production)
app.secret_key = "YOUR_SECRET_KEY"

# --- NEW: Define folders for file uploads and generated documents ---
UPLOAD_FOLDER = os.path.join(app.root_path, 'uploads')
DOCS_FOLDER = os.path.join(app.root_path, 'documents')
REPORTS_DIR = os.path.join(app.root_path, 'Stock Reports')

CHAT_HISTORY_FILE = os.path.join(app.root_path, "chat_history.json")

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOCS_FOLDER'] = DOCS_FOLDER
app.config['REPORTS_DIR'] = REPORTS_DIR

# Ensure folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOCS_FOLDER, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)


# ─── 1) End‐point to generate & store the XLSX ────────────────────────────────
@app.route('/analyze-evasiveness', methods=['POST'])
def analyze_evasiveness_route():
    payload = request.get_json() or {}
    company = payload.get('company')
    quarter = payload.get('quarter')
    year    = payload.get('year')

    # if any param is missing, try to parse free-text `query`
    if not (company and quarter and year):
        q = payload.get('query','')
        c, qtr, yr = extract_evasiveness_parameters_from_text(q)
        if not (c and qtr and yr):
            return jsonify({
              "error": "Missing company, quarter or year; "
                       "please supply as fields or in a natural‐language `query`."
            }), 400
        company, quarter, year = c, qtr, yr

    try:
        # this will throw FileNotFoundError if no transcript, or other on failure
        filepath = generate_evasiveness_report(
            company, int(year), int(quarter),
            app.config['UPLOAD_FOLDER']
        )
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except Exception as e:
        return jsonify({"error": f"Analysis failed: {e}"}), 500

    download_url = url_for('download_report',
                           filename=os.path.basename(filepath),
                           _external=True)
    return jsonify({
        "message": "Evasiveness report ready",
        "download_url": download_url
    }), 200


# ─── 2) Endpoint to serve the generated file ─────────────────────────────────
@app.route('/download/<filename>', methods=['GET'])
def download_report(filename):
    """
    Serves the XLSX previously generated.
    """
    full_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(full_path):
        return jsonify({"error": "File not found"}), 404
    return send_file(full_path, as_attachment=True)


# ─── (Optional) 3) Pure JSON summary route ───────────────────────────────────
@app.route('/evasiveness-summary', methods=['POST'])
def evasiveness_summary_route():
    payload = request.get_json() or {}
    company = payload.get('company')
    quarter = payload.get('quarter')
    year    = payload.get('year')

    if not (company and quarter and year):
        q = payload.get('query','')
        c, qtr, yr = extract_evasiveness_parameters_from_text(q)
        if not (c and qtr and yr):
            return jsonify({"error":"Missing company/quarter/year"}),400
        company, quarter, year = c, qtr, yr

    # 1) resolve ticker & transcript
    ticker = resolve_ticker(company)
    transcript = fetch_transcript(ticker, int(year), int(quarter))
    if not transcript:
        return jsonify({"error": "Transcript not found"}), 404

    # 2) parse/manage & analyze
    mgmt_turns = parse_management_turns(transcript["content"])
    df_statements, score = analyze_evasiveness(mgmt_turns)
    price = get_price(ticker, transcript["date"])

    summary = {
        "Company":          company,
        "Ticker":           ticker,
        "Year":             year,
        "Quarter":          quarter,
        "TranscriptDate":   transcript["date"],
        "EvasivenessScore": score,
        "SharePrice":       price
    }
    return jsonify(summary), 200

@app.route('/load_chat', methods=['GET'])
def load_chat():
    username = request.args.get('username')

    if not username:
        return jsonify({"error": "Username is required"}), 400

    try:
        if os.path.exists(CHAT_HISTORY_FILE):
            with open(CHAT_HISTORY_FILE, "r") as f:
                all_chats = json.load(f)
        else:
            all_chats = {}

        # Retrieve chat history for the specific user
        user_chat = all_chats.get(username, [])

        return jsonify({"chat": user_chat})

    except Exception as e:
        print(f"❌ Error loading chat history: {e}")
        return jsonify({"error": str(e)}), 500

def save_chat_history(history):
    with open(CHAT_HISTORY_FILE, "w") as f:
        json.dump(history, f)

# Excel file for storing users — absolute path
EXCEL_FILE = os.path.join(app.root_path, "users.xlsx")

def create_excel_if_not_exists():
    """Create the Excel file with a Users sheet if it doesn't exist."""
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Users"
        ws.append(["Username", "Password"])  # Headers
        wb.save(EXCEL_FILE)

def get_user_password(username):
    """Return the stored password for the given username, or None if not found."""
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["Users"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == username:
            return row[1]
    return None

def user_exists(username):
    return get_user_password(username) is not None

def add_user(username, password):
    """Append a new user to the Excel file (storing the password in plain text)."""
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["Users"]
    ws.append([username, password])
    wb.save(EXCEL_FILE)



#############################
# Authentication Routes
#############################

@app.route('/login', methods=['GET', 'POST'])
def login():
    create_excel_if_not_exists()  # Ensure Excel exists
    error = None
    if request.method == 'POST' and request.form.get("form_type") == "login":
        username = request.form.get('username')
        password = request.form.get('password')
        # For demo, you can continue to use the hard-coded admin credentials...
        # If admin logs in, let him in.
        if username == "admin" and password == "password":
            session['logged_in'] = True
            session['username'] = username
            return redirect(url_for('chat'))
        else:
            # Otherwise, check the Excel file
            stored_password = get_user_password(username)
            if stored_password and stored_password == password:
                session['logged_in'] = True
                session['username'] = username
                return redirect(url_for('chat'))
            else:
                error = "Invalid username or password. Please try again."
    return render_template('login.html', error=error)

@app.route('/signup', methods=['POST'])
def signup():
    create_excel_if_not_exists()
    error = None
    username = request.form.get('username')
    password = request.form.get('password')

    if not username or not password:
        error = "Please provide both email and password."
    elif username not in WHITELISTED_EMAILS:
        error = "This email is not authorized to sign up. Please contact admin."
    elif user_exists(username):
        error = "Email already registered. Please log in."
    else:
        add_user(username, password)
        return redirect(url_for('login'))

    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    session.pop('username', None)
    return redirect(url_for('login'))

#############################
# Main Chat UI Route
#############################

@app.route('/chat')
@app.route('/chat')
def chat():
    # Redirect to login if not authenticated or username is missing.
    if not session.get('logged_in') or not session.get('username'):
        return redirect(url_for('login'))
    # Return a rendered template with the username passed in.
    return render_template('index.html', username=session["username"])


# Redirect root to /chat if logged in, else to /login
@app.route('/')
def index():
    if session.get('logged_in'):
        return redirect(url_for('chat'))
    return redirect(url_for('login'))

@app.route('/generate-report', methods=['POST'])
def generate_report():
    """Handles stock report generation requests."""
    try:
        if not request.is_json:
            return jsonify({"error": "Request must be in JSON format."}), 400
        
        data = request.get_json()
        query = data.get("query")

        if not query:
            return jsonify({"error": "Query parameter is missing."}), 400

        print(f"✅ Received query: {query}")

        # Extract company and ticker
        ticker = process_query_1(query)
        company_name = None
        print("📥 Query:", query)
        print("🔍 Extracted company name:", company_name)
        print("🔍 Extracted ticker:", ticker)

        if not ticker:
            print("❌ Could not extract ticker.")
            return jsonify({"error": "Could not extract ticker from the query. Please try again."}), 400

        print(f"✅ Extracted Ticker: {ticker}")

        # Generate stock report
        file_path, wb, ws = create_stock_report(ticker)

        if not file_path:
            return jsonify({"error": f"Failed to generate stock report for {ticker}."}), 400
        file_name = os.path.basename(file_path)
        file_url = url_for('send_report', filename=file_name, _external=True)


        print(f"✅ Report successfully generated: {file_url}")

        return jsonify({
            "message": "Stock report generated successfully!",
            "file_url": file_url
        }), 200

    except Exception as e:
        print(f"🔥 Error in /generate-report: {str(e)}")
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500


@app.route("/generate-debt-covenants", methods=["POST"])
def generate_debt_covenants():
    data = request.get_json()
    query = data.get("query", "")

    # Basic ticker extraction — you likely already have this
    ticker = process_query_1(query)

    if not ticker:
        return jsonify({"response": "Could not identify a valid ticker."})

    html_table = analyze_debt_covenants(ticker)
    return jsonify({"response": html_table})


# Earnings Call Summary Route - This should ask for keywords if needed
@app.route('/generate-earnings-report', methods=['POST'])
def generate_earnings_report_route():
    """Handles keyword-based earnings report extraction."""
    try:
        if not request.is_json:
            return jsonify({"error": "Request must be in JSON format."}), 400

        data = request.get_json()
        query = data.get("query")
        keywords = data.get("keywords")  # Explicitly fetch keywords from payload

        if not query:
            return jsonify({"error": "Query parameter is missing."}), 400

        print(f"✅ Received earnings report query: {query}")

        # Extract company, quarter, and year
        ticker, quarter, year, _ = process_query_2(query)

        if not ticker or not quarter or not year:
            return jsonify({"error": "Could not extract all required parameters from the query."}), 400

        # If keywords are missing, ask the user for them
        if not keywords:
            print("🔎 Keywords missing. Asking user for keywords.")
            return jsonify({
                "message": "Please enter keywords in the format: growth, outlook, etc.",
                "request_keywords": True,
                "data": {"ticker": ticker, "quarter": quarter, "year": year}
            }), 200

        print(f"📊 Processing with keywords: {keywords}")
        response = main(ticker, quarter, year, keywords)

        if response and isinstance(response, str):
            if response.startswith("Excel file"):
                return jsonify({"message": response}), 200
            else:
                return jsonify({"error": response}), 400
        else:
            return jsonify({"error": "Unexpected error while processing the report."}), 500

    except Exception as e:
        print(f"🔥 Error in generate-earnings-report: {str(e)}")
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500


@app.route('/generate-earnings-call-summary', methods=['POST'])
def generate_earnings_call_summary_route():
    try:
        print("✅ Received request at /generate-earnings-call-summary")

        if not request.is_json:
            print("🚨 Error: Request is not in JSON format")
            return jsonify({"error": "Request must be in JSON format"}), 400

        data = request.get_json()
        print("✅ Received payload:", data)

        if "message" not in data:
            print("🚨 Error: 'message' key missing in payload")
            return jsonify({"error": "Missing 'message' in request"}), 400

        user_message = data["message"]
        print(f"✅ Processing earnings call summary for query: {user_message}")

        # Call function to process the earnings call summary
        summary = generate_earnings_call_summary(user_message)

        if not summary:
            print("🚨 Error: No summary generated")
            return jsonify({"error": "Failed to generate summary"}), 500

        print("✅ Earnings call summary successfully generated")
        return jsonify({"summary": summary})

    except Exception as e:
        print(f"🔥 ERROR: Internal Server Error: {str(e)}")
        return jsonify({"error": f"Internal Server Error: {str(e)}"}), 500

@app.route('/generate-market-performance', methods=['POST'])
def generate_market_performance():
    """This route handles market performance report generation based on the user query."""
    try:
        if request.is_json:
            data = request.get_json()
            query = data.get("query")

            if query:
                print(f"Received query: {query}")
                
                # Use the process_query_3 function to determine the type of market performance
                query_type = process_query_3(query)
                print(f"Query type determined by process_query_3: {query_type}")  # Debug print
                if not query_type:
                    return jsonify({"error": "Query type not recognized. Please specify global, emerging, or developed."}), 400

                # Define date range
                start_date = (datetime.today() - timedelta(days=5*365)).strftime('%Y-%m-%d')
                end_date = datetime.today().strftime('%Y-%m-%d')
                print(f"Start Date: {start_date}, End Date: {end_date}")
                # Trigger the appropriate report generation
                if query_type == "full":
                    file_path = generate_full_market_performance(start_date, end_date)
                elif query_type == "emerging":
                    file_path = generate_emerging_market_performance(start_date, end_date)
                elif query_type == "developed":
                    file_path = generate_developed_market_performance(start_date, end_date)
                else:
                    return jsonify({"error": "Query type not recognized."}), 400

                # Generate file URL to allow user to download
                file_name = os.path.basename(file_path)
                file_url = url_for('send_report', filename=file_name)

                return jsonify({
                    "message": "Market performance generated.",
                    "file_url": file_url
                }), 200
            else:
                return jsonify({"error": "Query parameter is missing"}), 400
        else:
            return jsonify({"error": "Request must be in JSON format"}), 400
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500


@app.route('/fetch-news-updates', methods=['POST'])
def fetch_news_updates():
    """API to fetch latest news updates."""
    try:
        if request.is_json:
            data = request.get_json()
            query = data.get("query")

            if not query:
                return jsonify({"error": "Query parameter is missing"}), 400

            print(f"Fetching news updates for: {query}")
            news_summary = get_latest_news_summary(query)

            # Ensure 'news' is always an array
            if isinstance(news_summary, dict) and "error" in news_summary:
                return jsonify({"news": []}), 200  # Return empty array if no news is found
            
            return jsonify({"news": news_summary}), 200
        else:
            return jsonify({"error": "Request must be in JSON format"}), 400
    except Exception as e:
        print(f"Error in fetching news: {e}")
        return jsonify({"error": str(e)}), 500

from openpyxl import Workbook

@app.route('/extract-earnings-call-data', methods=['POST'])
def extract_earnings_call_data():
    """Handles earnings call figure extraction."""
    try:
        if not request.is_json:
            return jsonify({"error": "Request must be in JSON format."}), 400

        data = request.get_json()
        query = data.get("query")

        if not query or not isinstance(query, str) or query.strip() == "":
            return jsonify({"error": "Query parameter is missing or invalid."}), 400

        print(f"📢 Extracting earnings figures for query: {query}")

        # Block summary-type queries
        if "summary" in query.lower() or "keywords" in query.lower():
            return jsonify({"error": "Use the earnings summary route for this request."}), 400

        # Run main extraction logic
        result = get_earnings_call_data(query)
        if "error" in result:
            return jsonify(result), 400

        # ✅ Generate Excel file
        company = result["company"]
        quarter = result["quarter"]
        year    = result["year"]
        sentences = result["sentences"]

        wb = Workbook()
        ws = wb.active
        ws.title = "Extracted Figures"
        ws.append(["Extracted Sentences"])

        for sentence in sentences:
            ws.append([sentence])

        # Define filename
        filename = f"{company}_{year}_Q{quarter}_figures.xlsx"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        wb.save(filepath)

        # Return JSON with download URL
        download_url = url_for('download_report', filename=filename, _external=True)

        result["download_url"] = download_url
        return jsonify(result), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/document-short-summary', methods=['POST'])
def document_short_summary():
    # At the start of your route function
    os.environ["DOCS_FOLDER"] = os.path.abspath(app.config['DOCS_FOLDER'])

    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file part in the request"}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400

        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        print("File saved to:", file_path)

        extracted_text = extract_pdf_text(file_path)
        if not extracted_text:
            print("Error: No text extracted")
            return jsonify({"error": "Failed to extract text from PDF"}), 400
        print("Extracted text length:", len(extracted_text))

        custom_title = generate_title(extracted_text)
        print("Custom title generated:", custom_title)

        # Split the text into smaller chunks (using a max chunk size of 1000 characters)
        text_chunks = split_text(extracted_text, max_chunk_size=1000)
        print("Number of text chunks:", len(text_chunks))

        # Summarize each chunk using a target of ~100 words per chunk
        summaries = []
        for chunk in text_chunks:
            summary = summarize_text(chunk, target_word_count=100)
            summaries.append(summary)
        print("Number of individual summaries:", len(summaries))

        # Iteratively refine the summaries into one final summary (targeting ~250 words)
        final_summary = iterative_refine_summary(summaries, target_word_count=250, batch_size=5)
        print("Final summary length:", len(final_summary))

        # Create and save the Word document with the dynamic title
        doc_path = create_word_document(final_summary, custom_title, docs_folder=app.config['DOCS_FOLDER'])
        doc_filename = os.path.basename(doc_path)

        return jsonify({
            "message": "Short document summary generated successfully!",
            "title": custom_title,
            "summary": final_summary,
            "document_url": url_for('download_doc', filename=doc_filename)
        }), 200


    except Exception as e:
        print(f"Error in document_short_summary: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/document-long-summary', methods=['POST'])
def document_long_summary():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file part in the request"}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400

        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        extracted_text = extract_pdf_text_long(file_path)
        if not extracted_text:
            return jsonify({"error": "Failed to extract text from PDF"}), 400

        text_chunks = split_text_long(extracted_text)
        summaries = [summarize_text_long(chunk) for chunk in text_chunks]
        final_summary = "\n".join(summaries)

        return jsonify({
            "message": "Long document summary generated successfully!",
            "summary": final_summary
        }), 200

    except Exception as e:
        print(f"Error in document_long_summary: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/document-esg', methods=['POST'])
def document_esg_summary():
    """
    This route processes an uploaded sustainability report PDF,
    extracts ESG data, and returns the summarized ESG information.
    """
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file part in the request"}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400

        # Extract text from the PDF using ESG's helper function
        pdf_text = ESG.extract_text_from_pdf(file)
        if not pdf_text:
            return jsonify({"error": "Failed to extract text from the PDF"}), 400

        # Process the text to extract ESG data
        esg_summary = ESG.extract_esg_data(pdf_text)
        
        return jsonify({
            "message": "ESG summary generated successfully!",
            "summary": esg_summary
        }), 200

    except Exception as e:
        print(f"Error in document_esg_summary: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/get10k', methods=['POST'])
def get_10k():
    try:
        data = request.get_json()
        company_name = data.get("company_name")
        years_input = data.get("years")
        if not company_name or not years_input:
            return jsonify({"error": "Missing 'company_name' or 'years' parameter."}), 400

        if isinstance(years_input, list):
            years = [str(year).strip() for year in years_input]
        else:
            years = [year.strip() for year in years_input.split(',')]

        # Call the refactored function from PDFF.py
        output_file = generate_10k_filings(company_name, years)
        return send_file(output_file, as_attachment=True)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route('/save_chat', methods=['POST'])
def save_chat():
    data = request.json
    username = data.get('username')
    chat_messages = data.get('chat', [])

    if not username:
        return jsonify({"error": "Username is required"}), 400

    try:
        if os.path.exists(CHAT_HISTORY_FILE):
            with open(CHAT_HISTORY_FILE, "r") as f:
                all_chats = json.load(f)
        else:
            all_chats = {}

        # ✅ Fix: Prevent duplicate messages and update history correctly
        if username in all_chats:
            existing_messages = {(msg['type'], msg['message']) for msg in all_chats[username]}  # Store existing messages as (type, message) tuples
            new_messages = [msg for msg in chat_messages if (msg['type'], msg['message']) not in existing_messages]  # Remove duplicates
            all_chats[username].extend(new_messages)
        else:
            all_chats[username] = chat_messages  # First-time user

        # Save updated chat history
        with open(CHAT_HISTORY_FILE, "w") as f:
            json.dump(all_chats, f, indent=4)

        return jsonify({"message": "Chat history saved successfully"}), 200

    except Exception as e:
        print(f"❌ Error saving chat history: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/process_report', methods=['POST'])
@app.route('/process_report', methods=['POST'])
def process_report():
    # 1) Validate inputs
    if 'file' not in request.files:
        return jsonify({'error': 'No PDF uploaded.'}), 400
    query = request.form.get('query', '').strip()
    if not query:
        return jsonify({'error': 'Please include your query (e.g. the company name) in "query".'}), 400

    # 2) Ensure output folders exist
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['DOCS_FOLDER'],  exist_ok=True)

    # 3) Save the uploaded PDF
    uploaded   = request.files['file']
    safe_name  = secure_filename(uploaded.filename)
    pdf_path   = os.path.join(app.config['UPLOAD_FOLDER'], safe_name)
    uploaded.save(pdf_path)

    # 4) Use the raw query as the company name (or parse out the company if you like)
    ticker = process_query_1(query)
    company_profile = get_fmp_json(f"profile/{ticker}?")
    company_name = company_profile[0]['companyName'] if company_profile and isinstance(company_profile, list) else ticker


    try:
        # 5) Call your modified generate_esg_report → returns the HTML path
        html_path = E_S_G.generate_esg_report(pdf_path, company_name)
        if not html_path:
            raise RuntimeError("ESG module did not return a report path")

        # 6) Verify it exists
        if not os.path.exists(html_path):
            raise FileNotFoundError(f"ESG report not found at {html_path}")

        # 7) Move it into your documents folder under a unique name
        unique_name = f"{int(time.time())}_{os.path.basename(html_path)}"
        final_path  = os.path.join(app.config['DOCS_FOLDER'], unique_name)
        shutil.move(html_path, final_path)

        # 8) Return a download URL
        download_url = url_for('download_doc', filename=unique_name, _external=True)
        return jsonify({
            'result':       "ESG Report generated successfully!",
            'download_url': download_url
        }), 200

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/docs/<filename>')
def download_doc(filename):
    """Serve the ESG HTML file from the docs folder."""
    return send_from_directory(
        app.config['DOCS_FOLDER'],
        filename,
        as_attachment=True  # triggers file download
    )

@app.route('/compare_insights', methods=['POST'])
def compare_insights_route():
    files = request.files.getlist('files')
    query = request.form.get('query')

    if not files:
        return jsonify({'error': 'No files provided.'}), 400
    if not query:
        return jsonify({'error': 'Query is required.'}), 400
    if len(files) > 5:
        return jsonify({'error': 'A maximum of 5 HTML files is allowed.'}), 400

    file_paths = []
    for file in files:
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        file_paths.append(file_path)

    try:
        # Generate comparison report and save to file
        output_file = os.path.join(app.config['DOCS_FOLDER'], "ESG_Comparison.html")
        ESGComp.generate_comparison_html([ESGComp.extract_esg_data(fp) for fp in file_paths], output_file)

        # Provide a direct download link
        download_url = url_for('download_doc', filename="ESG_Comparison.html", _external=True)

    except Exception as e:
        return jsonify({'error': str(e)}), 500

    return jsonify({
        'message': "ESG Comparison Report generated successfully!",
        'download_url': download_url
    })

@app.route('/Stock Reports/<filename>')
def send_report(filename):
    return send_from_directory(REPORTS_DIR, filename)

if __name__ == "__main__":
    app.run(debug=True)