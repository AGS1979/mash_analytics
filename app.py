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
from DebtCov import analyze_debt_covenants, extract_ticker_from_query
from openpyxl import Workbook
from redflaganalysis import get_red_flag_sentences
from Guidance import process_guidance_query  # 👈 Import the new module
from PyPDF2 import PdfReader
from InvMemo import run_pipeline  # ← your modularized memo logic
from FactorOptimizer import run_factor_optimizer_csv
from InvMemo import PDFQueryEngine  # Import the class we modularized earlier
from PrivateTransactionAnalyzer import analyze_transaction_doc
from DCF import (
    get_fmp_ticker, get_fmp_data, get_current_price, extract_text_from_files,
    generate_dcf_logic, clean_and_format_dcf_output, save_excel, extract_kpi_drivers
)
from InvMemoInfographic import generate_infographic_html
from SpecialSituations import generate_special_situation_note  # Your core function
from SSInfographic import generate_infographic_entrypoint  # Your earlier function
from portagent import index_pdf, query_gpt_prompt  # 🔁 import your functions
from openai import OpenAI

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "infographics"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# Define your email whitelist here
WHITELISTED_EMAILS = {
    "avinashg.singh@aranca.com",
    "ujjal.roy@aranca.com",
    "rohit.dhawan@aranca.com",
    "avi104@yahoo.co.in",
    "vishal.kumar@aranca.com"
}


users_db = {}

ALLOWED_EXTENSIONS = {"pdf", "docx", "pptx", "xlsx", "xlsm"}

# Add the current directory to the sys.path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

app = Flask(__name__, static_folder='static')
CORS(app)
# Allow up to 100 MB uploads
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024



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
        ws.append(["Username", "Password", "FirstName", "Company"])
        wb.save(EXCEL_FILE)

def get_user_info(username):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["Users"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == username:
            return {
                "username": row[0],
                "password": row[1],
                "first_name": row[2] if len(row) > 2 else "",
                "company_name": row[3] if len(row) > 3 else ""
            }
    return None

def user_exists(username):
    return get_user_info(username) is not None

def add_user(username, password, first_name, company_name):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["Users"]
    ws.append([username, password, first_name, company_name])
    wb.save(EXCEL_FILE)



#############################
# Authentication Routes
#############################

@app.route('/login', methods=['GET', 'POST'])
@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST' and request.form.get("form_type") == "login":
        username = request.form.get('username')
        password = request.form.get('password')

        user = users_db.get(username)
        if user and user["password"] == password:
            session['logged_in'] = True
            session['username'] = username
            session['first_name'] = user["first_name"]
            session['company_name'] = user["company_name"]
            return redirect(url_for('index'))  # ✅ redirect to your index.html
        else:
            error = "Invalid username or password."

    return render_template('login.html', error=error)


# ---------------------------------------------------
# 6) If using Flask, expose an endpoint
# ---------------------------------------------------


def allowed_file(filename: str) -> bool:
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    return ext in ALLOWED_EXTENSIONS

@app.route("/analyze_deal", methods=["POST"])
def analyze_deal_endpoint():
    """
    Expects a multipart-form POST with:
      • file (under key "file", must be one of PDF/DOCX/PPTX/Excel)
      • query: the user’s free-form query (e.g. “red flags” or “SWOT analysis”)
      • optional form fields:
          - chunk_size: integer (e.g. 1500). Defaults to 1800 if omitted/invalid.
          - run_deep_dives: "true" or "false" (case-insensitive). Defaults to true.
          - deep_dive_sections: comma-separated list of section names 
              (e.g. "Market Analysis,Financial Performance,Exit Strategy").
          - chunk_prompt: (optional) full-text prompt containing "{{TEXT}}"
          - aggregate_prompt: (optional) full-text prompt containing "{{SUMMARIES}}"
          - prompt_<SectionKey>: for each deep-dive section, custom prompt containing
              "{{SECTION_NAME}}" and "{{TEXT}}".
    Returns JSON:
      {
        "aggregate_analysis": { … },
        "deep_dive_Market_Analysis": "...",
        "deep_dive_Financial_Performance": "...",
        ...
        "pages_scanned": 50,
        "relevant_pages": [2, 5, 7],
        "chunks_used": 3
      }
    """
    print("[DEBUG] /analyze_deal: request received")

    # 1) Validate that a file was uploaded
    if "file" not in request.files:
        print("[DEBUG] /analyze_deal: no file part")
        return jsonify({"error": "No file part"}), 400

    file = request.files["file"]
    if file.filename == "":
        print("[DEBUG] /analyze_deal: no selected file")
        return jsonify({"error": "No selected file"}), 400

    # 2) Check extension
    if not allowed_file(file.filename):
        print(f"[DEBUG] /analyze_deal: unsupported file type: {file.filename}")
        return jsonify({"error": "Unsupported file type"}), 400

    # 3) Save the file temporarily
    filename = secure_filename(file.filename)
    save_dir = "/tmp"
    os.makedirs(save_dir, exist_ok=True)
    save_path = os.path.join(save_dir, filename)
    file.save(save_path)
    print(f"[DEBUG] /analyze_deal: saved file to {save_path}")

    # 4) Read the user’s free-form query
    user_query = request.form.get("query", "").strip()
    if not user_query:
        print("[DEBUG] /analyze_deal: no query provided")
        try:
            os.remove(save_path)
        except OSError:
            pass
        return jsonify({"error": "Please provide a ‘query’ field describing what you want (e.g. 'red flags')."}), 400

    # 5a) chunk_size (default to 1800)
    raw_chunk_size = request.form.get("chunk_size", "").strip()
    try:
        chunk_size = int(raw_chunk_size) if raw_chunk_size else 1800
        if chunk_size <= 0:
            chunk_size = 1800
    except (ValueError, TypeError):
        chunk_size = 1800

    # 5b) run_deep_dives (“true”/“false” → bool)
    raw_run_deep = request.form.get("run_deep_dives", "true").strip().lower()
    run_deep_dives = (raw_run_deep == "true")

    # 5c) deep_dive_sections (comma-separated list or default)
    raw_sections = request.form.get("deep_dive_sections", "").strip()
    if raw_sections:
        deep_dive_sections = [sec.strip() for sec in raw_sections.split(",") if sec.strip()]
    else:
        deep_dive_sections = [
            "Market Analysis",
            "Financial Performance",
            "Operational Risks",
            "Exit Strategy"
        ]

    # 5d) custom prompts for chunk-level/aggregate-level
    chunk_prompt = request.form.get("chunk_prompt", None)
    if chunk_prompt is not None and not chunk_prompt.strip():
        chunk_prompt = None

    aggregate_prompt = request.form.get("aggregate_prompt", None)
    if aggregate_prompt is not None and not aggregate_prompt.strip():
        aggregate_prompt = None

    # 5e) for each deep-dive section, look for “prompt_<SectionKey>”
    deep_dive_prompts: dict[str, str] = {}
    for section in deep_dive_sections:
        key = section.replace(" ", "_")
        form_name = f"prompt_{key}"
        custom = request.form.get(form_name, None)
        if custom and custom.strip():
            deep_dive_prompts[key] = custom

    if not deep_dive_prompts:
        deep_dive_prompts = None

    print(f"[DEBUG] /analyze_deal: user_query='{user_query}', chunk_size={chunk_size}, run_deep_dives={run_deep_dives}")
    print(f"[DEBUG] /analyze_deal: deep_dive_sections={deep_dive_sections}, deep_dive_prompts keys={list(deep_dive_prompts or [])}")
    print(f"[DEBUG] /analyze_deal: chunk_prompt provided? {'Yes' if chunk_prompt else 'No'}")
    print(f"[DEBUG] /analyze_deal: aggregate_prompt provided? {'Yes' if aggregate_prompt else 'No'}")

    # 6) Call the analyzer
    try:
        analysis = analyze_transaction_doc(
            filepath=save_path,
            user_query=user_query,
            run_deep_dives=run_deep_dives,
            chunk_size=chunk_size,
            deep_dive_sections=deep_dive_sections,
            chunk_prompt=chunk_prompt,
            aggregate_prompt=aggregate_prompt,
            deep_dive_prompts=deep_dive_prompts
        )
    except Exception as e:
        print(f"[ERROR] /analyze_deal: analysis failed: {e}")
        try:
            os.remove(save_path)
        except OSError:
            pass
        return jsonify({"error": str(e)}), 500

    # 7) Cleanup temp file
    try:
        os.remove(save_path)
    except OSError:
        pass

    print(f"[DEBUG] /analyze_deal: returning analysis JSON")
    return jsonify(analysis)




@app.route('/signup', methods=['POST'])
def signup():
    username = request.form.get('username')
    password = request.form.get('password')
    first_name = request.form.get('first_name')
    company_name = request.form.get('company_name')

    if not username or not password:
        return render_template('login.html', error="Please provide both username and password.")
    elif username not in WHITELISTED_EMAILS:
        return render_template('login.html', error="This email is not authorized to sign up.")

    if username in users_db:
        return render_template('login.html', error="User already exists. Please log in.")

    # ✅ Save user in memory
    users_db[username] = {
        "password": password,
        "first_name": first_name,
        "company_name": company_name
    }

    # ✅ Set session and redirect
    session['logged_in'] = True
    session['username'] = username
    session['first_name'] = first_name
    session['company_name'] = company_name

    return redirect(url_for('chat'))  # or index, depending on your design




UPLOAD_DIR = "uploads"
OUTPUT_DIR = "memos"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.route("/generate-special-situation-memo", methods=["POST"])
def generate_special_situation_memo():
    try:
        company_name = request.form.get('company_name', '').strip()
        situation_type = request.form.get('situation_type', '').strip()
        uploaded_files = request.files.getlist('files')

        if not company_name or not situation_type:
            return jsonify({"error": "Missing company_name or situation_type"}), 400

        if not uploaded_files:
            return jsonify({"error": "No files uploaded"}), 400

        # Save uploaded files to UPLOAD_FOLDER
        saved_paths = []
        for file in uploaded_files:
            if not file.filename.lower().endswith(('.pdf', '.docx')):
                continue
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            saved_paths.append(file_path)

        if not saved_paths:
            return jsonify({"error": "No valid files saved."}), 400

        # Save memo to DOCS_FOLDER (used by download_doc route)
        raw_name       = f"{company_name} – {situation_type} Memo.docx"
        output_filename = secure_filename(raw_name)
        output_path     = os.path.join(app.config['DOCS_FOLDER'], output_filename)

        # Call generator
        generate_special_situation_note(company_name, situation_type, saved_paths, output_path)

        if not os.path.exists(output_path):
            return jsonify({"error": "Memo generation failed."}), 500

        # ✅ Use external URL for full download link
        download_url = url_for('download_doc', filename=output_filename, _external=True)

        return jsonify({
            "message": "Memo generated successfully!",
            "download_url": download_url
        }), 200

    except Exception as e:
        print(f"🔥 Error in /generate-special-situation-memo: {e}")
        return jsonify({"error": f"❌ Error generating memo: {str(e)}"}), 500


@app.route("/generate-infographic", methods=["POST"])
def generate_infographic():
    if 'company_name' not in request.form:
        return jsonify({"error": "Missing company name"}), 400
    if 'situation_type' not in request.form:
        return jsonify({"error": "Missing situation type"}), 400
    if 'memo_file' not in request.files:
        return jsonify({"error": "Missing uploaded memo"}), 400

    company_name = request.form['company_name'].strip()
    situation_type = request.form['situation_type'].strip()
    memo_file = request.files['memo_file']

    if memo_file.filename == '':
        return jsonify({"error": "No file selected"}), 400
    if not memo_file.filename.lower().endswith('.docx'):
        return jsonify({"error": "Only .docx files are supported"}), 400

    memo_filename = secure_filename(memo_file.filename)
    memo_path = os.path.join(UPLOAD_DIR, memo_filename)
    memo_file.save(memo_path)

    html_filename = f"{company_name.replace(' ', '_')}_Infographic.html"
    output_path = os.path.join(OUTPUT_DIR, html_filename)

    try:
        generate_infographic_entrypoint(memo_path, company_name, situation_type, output_path)
    except Exception as e:
        return jsonify({"error": f"❌ Error generating infographic: {str(e)}"}), 500

    return send_file(output_path, as_attachment=True, download_name=html_filename, mimetype="text/html")


@app.route("/analyze-portfolio-company", methods=["POST"])
def analyze_portfolio_company():
    try:
        company_names = request.form.get('company_name', '').split(",")
        company_names = [c.strip() for c in company_names if c.strip()]
        user_query = request.form.get('query', '').strip()
        uploaded_files = request.files.getlist('files')

        if not company_names or not user_query:
            return jsonify({"error": "Missing company name(s) or query"}), 400
        if not uploaded_files:
            return jsonify({"error": "No files uploaded"}), 400

        saved_paths = []
        for file in uploaded_files:
            if not file.filename.lower().endswith(('.pdf', '.docx', '.txt')):
                continue
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            saved_paths.append(file_path)

        if not saved_paths:
            return jsonify({"error": "No valid files uploaded. Only .pdf, .docx, and .txt are accepted."}), 400

        for file_path in saved_paths:
            for company in company_names:
                index_pdf(file_path, company)

        prompt = query_gpt_prompt(user_query, company_names)

        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You're an AI portfolio analysis assistant."},
                {"role": "user", "content": prompt}
            ]
        )

        answer = response.choices[0].message.content.strip()
        return jsonify({"result": answer}), 200

    except Exception as e:
        print(f"🔥 Error in /analyze-portfolio-company: {e}")
        return jsonify({"error": f"❌ {str(e)}"}), 500



@app.route('/custom-agents-data')
def get_custom_agents():
    agents = [
        {
            "id": "PE_analyzer",
            "name": "Private Transaction Analyzer",
            "category": "Private Equity",
            "description": "Analyze CIMs, models, and memos to extract key deal insights..",
            "output": "Sample signals file"
        },
        {
            "id": "factor_opt",
            "name": "Factor Exposure Optimizer",
            "category": "Portfolio Optimization",
            "description": "Rebalances portfolio to maintain desired exposure to style factors like value and momentum.",
            "output": "Optimization result file"
        },
        {
            "id": "Special_Situations_Analyzer",
            "name": "Special Situations Analyzer",
            "category": "Investment Notes",
            "description": "Creates an investment note based on a special situation",
            "output": "Regime classification file"
        },
        {
            "id": "Pre-IPO_Investment_Memo",
            "name": "Pre-IPO Investment Memo",
            "category": "IPOs",
            "description": "Create pre-ipo investment memos using publicly available DRHPs.",
            "output": "Sample sector heatmap output..."
        },
        {
            "id": "dcf_analyzer",
            "name": "DCF Analyzer",
            "category": "Valuation",
            "description": "Perform a discounted cash flow (DCF) valuation for a given stock ticker.",
            "output": "Excel/JSON DCF report"
        },
        {
            "id": "Portfolio_Analyzer",
            "name": "Portfolio Agent",
            "category": "Private Equity",
            "description": "Ask questions about uploaded portfolio company PDFs (investment memos, updates, etc.)",
            "output": "GPT-based analysis answer"
        },
        {
            "id": "new_agent_1",
            "name": "Sector Heatmap Analyzer",
            "category": "Sector Insights",
            "description": "Visualizes sector performance across multiple dimensions like momentum and volatility.",
            "output": "Sample sector heatmap output..."
        },
        {
            "id": "new_agent_1",
            "name": "Sector Heatmap Analyzer",
            "category": "Sector Insights",
            "description": "Visualizes sector performance across multiple dimensions like momentum and volatility.",
            "output": "Sample sector heatmap output..."
        },
        {
            "id": "new_agent_1",
            "name": "Sector Heatmap Analyzer",
            "category": "Sector Insights",
            "description": "Visualizes sector performance across multiple dimensions like momentum and volatility.",
            "output": "Sample sector heatmap output..."
        }
    ]
    return jsonify(agents)


@app.route('/generate-preipo-memo', methods=['POST'])
def generate_preipo_memo():
    try:
        file = request.files.get('file')
        notes = request.form.get('notes', '')  # optional user input
        if not file or not file.filename.endswith('.pdf'):
            return jsonify({'error': 'Please upload a valid PDF file.'}), 400

        file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        file.save(file_path)

        output_path = run_pipeline(file_path, notes)  # your DRHP processing logic here

        if not output_path or not os.path.exists(output_path):
            return jsonify({'error': 'Failed to generate memo.'}), 500

        download_url = url_for('download_doc', filename=os.path.basename(output_path), _external=True)
        return jsonify({
            "message": "Memo generated successfully!",
            "download_url": download_url
        }), 200

    except Exception as e:
        print(f"🔥 Error in /generate-preipo-memo: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/generate-preipo-infographic', methods=['POST'])
def generate_preipo_infographic():
    try:
        file = request.files.get('file')
        if not file or not file.filename.endswith('.docx'):
            return jsonify({'error': 'Upload a valid .docx memo.'}), 400

        # Read uploaded file into memory
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp:
            file.save(temp.name)
            company_name = file.filename.split('_')[0]
            html_content = generate_infographic_html(temp.name, company_name)

        return jsonify({
            "message": "Infographic ready!",
            "html": html_content
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

        

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    session.pop('username', None)
    return redirect(url_for('login'))

#############################
# Main Chat UI Route
#############################

@app.route('/chat')
def chat():
    if not session.get('logged_in') or not session.get('username'):
        return redirect(url_for('login'))

    return render_template(
        'index.html',  # ✅ you're rendering login.html after login
        username=session.get("username"),  # 👈 important
        first_name=session.get("first_name", ""),
        company_name=session.get("company_name", "")
    )


@app.route("/prepare-dcf-financials", methods=["POST"])
def prepare_dcf_financials():
    try:
        company_name = request.form.get("company_name")
        if not company_name:
            return jsonify({"error": "Missing company name."}), 400

        ticker = get_fmp_ticker(company_name)
        current_price = get_current_price(ticker)
        financials_df = get_fmp_data(ticker)

        excel_buffer = save_excel(financials_df)
        timestamp = int(time.time())
        excel_filename = f"{ticker}_financials_{timestamp}.xlsx"
        excel_path = os.path.join(app.config["UPLOAD_FOLDER"], excel_filename)
        with open(excel_path, "wb") as f:
            f.write(excel_buffer.getbuffer())

        return jsonify({
            "message": f"Financials prepared for {ticker}.",
            "ticker": ticker,
            "cmp": current_price,
            "excel_url": url_for("download_report", filename=excel_filename, _external=True)
        }), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/run-dcf-analysis", methods=["POST"])
def run_dcf_analysis():
    try:
        # ───── STEP 1: PARSE INPUTS ─────
        ticker = request.form.get("ticker")
        cmp = float(request.form.get("cmp", 0))
        wacc = float(request.form.get("wacc", 8.0))
        mode = request.form.get("mode", "llm")
        confirm_guidance = request.form.get("confirm_guidance", "false") == "true"

        uploaded_files = request.files.getlist("files")
        excel_file = request.files.get("financials")

        if not ticker or not cmp or not uploaded_files or not excel_file:
            return jsonify({"error": "Missing ticker, CMP, financials, or files."}), 400

        # ───── STEP 2: LOAD FINANCIALS ─────
        df = pd.read_excel(excel_file)

        # ───── STEP 3: PARSE DOCUMENTS ─────
        filepaths = []
        for f in uploaded_files:
            filename = secure_filename(f.filename)
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            f.save(filepath)
            filepaths.append(filepath)

        file_objects = [open(fp, "rb") for fp in filepaths]
        documents_text = extract_text_from_files(file_objects)

        # ───── STEP 4: EXTRACT GUIDANCE FROM DOCUMENTS ─────
        guidance_summary = extract_kpi_drivers(documents_text)

        if not confirm_guidance:
            # Return guidance preview for confirmation step
            guidance_html = f"""
                <div class="guidance-confirm-box">
                    <h3>📌 Extracted 2025 Guidance Summary</h3>
                    <pre style="white-space: pre-wrap; font-family: monospace; background:#222; padding:1em; color:#eee;">{guidance_summary}</pre>
                    <p>✅ If the above looks accurate, re-run the DCF with <code>confirm_guidance=true</code>.</p>
                </div>
            """
            return jsonify({
                "html": guidance_html
            })

        # ───── STEP 5: COMBINE STRUCTURED CONTEXT FOR DCF ─────
        documents_text_annotated = f"""
🔹 Extracted 2025 Guidance from Uploaded Documents:

{guidance_summary}

📄 Full Supporting Extract (for optional reference):

{documents_text}
"""

        dcf_mode = "Quick mechanical DCF" if mode == "own" else "Detailed LLM-based DCF with strategy commentary"

        # ───── STEP 6: RUN GPT DCF ─────
        raw_output = generate_dcf_logic(df, documents_text_annotated, wacc, cmp, dcf_mode)
        html_output = clean_and_format_dcf_output(raw_output, cmp)

        # ───── STEP 7: (Optional) Mismatch Highlight ─────
        mismatches = []
        if "free cash flow" in guidance_summary.lower():
            import re
            fcf_range = re.findall(r"\$([0-9.]+)B\s*-\s*\$([0-9.]+)B", guidance_summary)
            if fcf_range:
                try:
                    low, high = map(float, fcf_range[0])
                    if f"${low:.1f}B" not in raw_output and f"${high:.1f}B" not in raw_output:
                        mismatches.append(f"⚠️ DCF output may understate 2025 FCF. Guidance was ${low:.1f}B–${high:.1f}B.")
                except:
                    pass

        if mismatches:
            mismatch_html = "<div class='mismatch-warning'><h4>🔍 Mismatch Alerts</h4><ul>"
            mismatch_html += "".join(f"<li>{m}</li>" for m in mismatches)
            mismatch_html += "</ul></div>"
            html_output = mismatch_html + html_output

        return jsonify({
            "html": html_output
        }), 200

    except Exception as e:
        print(f"❌ Error in /run-dcf-analysis: {e}")
        return jsonify({"error": str(e)}), 500





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

        print(f"✅ Using sanitized ticker: {ticker}")
        print(f"✅ Report successfully generated: {file_url}")

        return jsonify({
            "message": "Stock report generated successfully!",
            "file_url": file_url
        }), 200

    except Exception as e:
        print(f"🔥 Error in /generate-report: {str(e)}")
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500


@app.route('/generate-debt-covenants', methods=['POST'])
def generate_debt_covenants():
    data = request.get_json()
    query = data.get("query", "")
    print("📩 Raw query from frontend:", query)

    ticker = extract_ticker_from_query(query)
    print("🎯 Extracted ticker:", ticker)

    if not ticker:
        return jsonify({"response": "<p>Could not identify a valid ticker.</p>"})

    result = analyze_debt_covenants(ticker)
    return jsonify({"response": result})


@app.route('/analyze-redflags', methods=['POST'])
def analyze_red_flags():
    try:
        # 1) Check for file upload (PDF)
        if 'file' in request.files:
            file = request.files['file']
            reader = PdfReader(file)
            text = ''
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text

        # 2) Or check for raw text input
        else:
            json_data = request.get_json()
            text = json_data.get("text", "")

        # 3) Validate
        if not text.strip():
            return jsonify({"error": "No valid text provided"}), 400

        # 4) Run red flag extraction
        red_flags = get_red_flag_sentences(text)

        # 5) Return JSON result — inline only
        return jsonify({
            "count": len(red_flags),
            "red_flags": [{"sentence": s, "score": score} for s, score in red_flags]
        }), 200

    except Exception as e:
        print("❌ Error in /analyze-redflags:", e)
        return jsonify({"error": str(e)}), 500


@app.route('/optimize-factor-portfolio', methods=['POST'])
def optimize_factor_portfolio():
    
    try:
        file = request.files.get('file')
        if not file or not file.filename.endswith('.csv'):
            return jsonify({"status": "error", "message": "Please upload a valid CSV file."}), 400

        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        file.save(temp_path)

        raw = request.form.get('target_exposures')
        to_raw = request.form.get('turnover_limit')

        if not raw:
            return jsonify({"status": "error", "message": "Target exposures are required."}), 400

        try:
            target_exposures = json.loads(raw)
        except json.JSONDecodeError:
            return jsonify({"status": "error", "message": "Invalid format for target_exposures. Must be JSON."}), 400

        turnover_limit = float(to_raw) if to_raw else None

        result = run_factor_optimizer_csv(temp_path, target_exposures, turnover_limit)
        return jsonify(result)

    except Exception as e:
        print(f"🔥 Error in /optimize-factor-portfolio: {e}", file=sys.stderr)
        return jsonify({"status": "error", "message": str(e)}), 500



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


@app.route('/query-pdf', methods=['POST'])
def query_pdf():
    try:
        file = request.files.get('file')
        query = request.form.get('query', '')

        if not file or not query:
            return jsonify({'error': 'File and query are required'}), 400

        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        file.save(temp_path)

        engine = PDFQueryEngine(api_key=DEEPSEEK_API_KEY)
        answer, cited_pages = engine.answer_query(temp_path, query)

        return jsonify({
            "answer": answer,
            "pages": cited_pages
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500



@app.route('/analyze-guidance-change', methods=['POST'])
def analyze_guidance_change_route():
    try:
        data = request.get_json()
        query = data.get("query", "").strip()

        if not query:
            return jsonify({"error": "Query parameter is missing."}), 400

        print(f"📩 Guidance change query received: {query}")
        response = process_guidance_query(query)

        if not response:
            return jsonify({"error": "Failed to generate a response."}), 500

        return jsonify({"message": response}), 200

    except Exception as e:
        print(f"🔥 Error in /analyze-guidance-change: {str(e)}")
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