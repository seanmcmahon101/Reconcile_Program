import os
import logging
import time
import re
import json
import pandas as pd
import numpy as np
from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash, make_response, jsonify
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import google.generativeai as genai
import markdown
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import threading
import queue
from concurrent.futures import ThreadPoolExecutor
import uuid
import tempfile
from decimal import Decimal, ROUND_HALF_UP
import hashlib
from datetime import timedelta
try:
    from flask.json import JSONEncoder
except ImportError:
    # For Flask 2.3+
    from flask.json.provider import JSONEncoder

# Import for security
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from flask_wtf.csrf import CSRFProtect
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

# Custom JSON encoder for NumPy types
class CustomJSONEncoder(JSONEncoder):
    def default(self, obj):
        if isinstance(obj, np.integer):
            return int(obj)
        elif isinstance(obj, np.floating):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        elif isinstance(obj, pd.DataFrame):
            return obj.to_dict()
        elif isinstance(obj, pd.Series):
            return obj.to_dict()
        else:
            return super().default(obj)

# Create Flask app
app = Flask(__name__)
app.json_encoder = CustomJSONEncoder
app.secret_key = os.environ.get('FLASK_SECRET_KEY')
if not app.secret_key:
    # Generate a random key if not provided, but log a warning
    app.secret_key = os.urandom(24).hex()
    logging.warning("FLASK_SECRET_KEY not set. Using a random secret key - sessions will reset on server restart.")

# Make sessions persistent
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=1)
@app.before_request
def make_session_permanent():
    session.permanent = True

# Security enhancements
csrf = CSRFProtect(app)
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["200 per day", "50 per hour"],
    storage_uri="memory://"
)

# --- Configuration ---
API_KEY = os.getenv("GEMINI_API_KEY")
DEFAULT_MODEL_NAME = 'gemini-2.0-flash'
THINKING_MODEL_NAME = 'gemini-2.0-flash-thinking-exp'
PREMIUM_MODEL_NAME = 'gemini-2.0-pro'  # Added premium model option

# System prompts with more accounting-specific guidance
SYSTEM_PROMPT = """You are an expert in analyzing Excel spreadsheets specifically for accounting and financial purposes.
Focus on identifying account types, transaction patterns, financial formulas, and reconciliation structures.
Explain the purpose, structure, and formulas in a clear and detailed manner appropriate for accounting professionals.
For financial data, identify the accounting principles being applied (like GAAP, IFRS) where possible.
Format your explanations in Markdown, using headings, bullet points, and code blocks for readability.
Highlight important financial insights, potential errors in formulas, and accounting best practices."""

FORMULA_SYSTEM_PROMPT = """You are an expert in creating Excel formulas specifically for accounting and financial reconciliation.
Your task is to provide the most efficient Excel formula solution.
For accounting functions, consider:
1. Conditional logic for matching transactions across systems
2. VLOOKUP/XLOOKUP for finding corresponding entries
3. SUMIF/SUMIFS for conditional totaling
4. DATE functions for period matching
5. Text manipulation for standardizing reference numbers
Explain both the formula and its accounting purpose. Provide examples with realistic accounting data.
Include error handling and data validation best practices for reliable financial reporting."""

RECONCILIATION_SYSTEM_PROMPT = """You are an expert accountant specializing in financial reconciliation.
Analyze these two Excel sheets representing financial data sets that need to be reconciled.

In your detailed reconciliation report, identify and explain:
1. QUANTITATIVE ANALYSIS:
   - Total variance between datasets (sum, average, percentage)
   - Specific transactions with discrepancies (exact amounts and line references)
   - Missing transactions in either dataset
   - Duplicate entries or potential double-counting

2. QUALITATIVE ANALYSIS:
   - Root causes of discrepancies (timing differences, classification errors, etc.)
   - Patterns in the variances (consistent mismatches, systematic errors)
   - Data integrity issues (format inconsistencies, calculation errors)
   - Recommendation for reconciliation approach

3. ACCOUNTING IMPLICATIONS:
   - Financial reporting impact of discrepancies
   - Potential compliance or audit concerns
   - Suggested journal entries to resolve differences
   - Controls to prevent future reconciliation issues

Format your report professionally with clear headings, tables for numerical comparisons, and executive summary.
Include specific account codes, amounts, and variance calculations where relevant."""

# Don't forget these important prompt components
PROMPT_PREFIX = "The Excel sheet contains the following information in a structured way:\n"
PROMPT_SUFFIX = "\nPlease provide a detailed explanation in Markdown format. Explain what this sheet is about, what each column/section represents, and how the data is structured and what its purpose might be. If there are formulas, explain their logic in simple terms. Structure your answer with headings, bullet points, and code blocks where appropriate for formulas or data examples to enhance readability."

UPLOAD_FOLDER = tempfile.mkdtemp()  # Use temporary directory for security
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}
EXPORTS_FOLDER = 'exports'
os.makedirs(EXPORTS_FOLDER, exist_ok=True)

# File export configuration with better naming
DEFAULT_DOCX_FILENAME = "excel_explanation.docx"
FORMULA_DOCX_FILENAME = "excel_formula.docx"
CHAT_DOCX_FILENAME = "excel_chat.docx"
RECONCILIATION_DOCX_FILENAME = "reconciliation_report.docx"
RECONCILIATION_EXCEL_FILENAME = "reconciliation_detailed.xlsx"

# Set up memory cache for API responses
CACHE = {}
CACHE_TIMEOUT = 3600  # 1 hour

# Enhanced logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(module)s - %(message)s',
    handlers=[
        logging.FileHandler("app.log"),
        logging.StreamHandler()
    ]
)

# --- Flask-Login Configuration ---
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = "Please log in to access this reconciliation system."

# --- User Management with Enhanced Security ---
users = {}
for i in range(1, 10):  # Support for up to 9 users
    username = os.getenv(f"USER{i}_USERNAME")
    password = os.getenv(f"USER{i}_PASSWORD")
    role = os.getenv(f"USER{i}_ROLE", "user")  # Added role-based permissions
    if username and password:
        users[i] = {
            'username': username,
            'password_hash': generate_password_hash(password),
            'role': role
        }
    else:
        if i <= 4:  # Only log warnings for first 4 users that might be expected
            logging.warning(f"User {i} credentials not fully configured. User {i} will not be available.")

# If no users defined, create a default admin user with environment variables or strong defaults
if not users:
    default_admin = os.getenv("DEFAULT_ADMIN_USERNAME", "admin")
    default_password = os.getenv("DEFAULT_ADMIN_PASSWORD")
    if not default_password:
        # Generate and display a random password if none is set
        import secrets
        default_password = secrets.token_urlsafe(12)
        print(f"WARNING: No users configured. Created default admin user:")
        print(f"Username: {default_admin}")
        print(f"Password: {default_password}")
        logging.warning(f"No users configured. Created default admin user: {default_admin}")

    users[1] = {
        'username': default_admin,
        'password_hash': generate_password_hash(default_password),
        'role': 'admin'
    }

class User(UserMixin):
    def __init__(self, id, username, password_hash, role='user'):
        self.id = id
        self.username = username
        self.password_hash = password_hash
        self.role = role

    def verify_password(self, password):
        return check_password_hash(self.password_hash, password)

    def is_admin(self):
        return self.role == 'admin'

@login_manager.user_loader
def load_user(user_id):
    user_data = users.get(int(user_id))
    if user_data:
        return User(
            id=user_id,
            username=user_data['username'],
            password_hash=user_data['password_hash'],
            role=user_data.get('role', 'user')
        )
    return None

# --- Helper Functions ---
def make_session_safe(data):
    """Convert data to JSON-serializable format safe for session storage."""
    if isinstance(data, dict):
        return {str(key): make_session_safe(value) for key, value in data.items()}
    elif isinstance(data, list):
        return [make_session_safe(item) for item in data]
    elif isinstance(data, (np.integer, np.int64)):
        return int(data)
    elif isinstance(data, (np.floating, np.float64)):
        return float(data)
    elif isinstance(data, np.ndarray):
        return make_session_safe(data.tolist())
    elif isinstance(data, pd.DataFrame):
        return data.to_dict()
    elif isinstance(data, pd.Series):
        return make_session_safe(data.to_dict())
    else:
        return data

def configure_api():
    """Configures the Gemini API with the API key with error handling."""
    if not API_KEY:
        logging.error("API_KEY environment variable not set.")
        return False
    try:
        genai.configure(api_key=API_KEY)
        # Test connection with a simple request
        model = genai.GenerativeModel(DEFAULT_MODEL_NAME)
        response = model.generate_content("Hello")
        if response:
            logging.info("Successfully connected to Gemini API.")
            return True
        else:
            logging.error("Failed to get response from Gemini API.")
            return False
    except Exception as e:
        logging.error(f"Error configuring Gemini API: {e}")
        return False

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def sanitize_filename(filename):
    """Sanitize filename to avoid path traversal attacks."""
    return secure_filename(filename)

def generate_unique_filename(original_filename):
    """Generate a unique filename to prevent overwriting."""
    timestamp = int(time.time())
    random_id = uuid.uuid4().hex[:8]
    _, ext = os.path.splitext(original_filename)
    return f"{timestamp}_{random_id}{ext}"

def load_excel_data(file_path):
    """Enhanced Excel loading with pandas support and error handling."""
    try:
        # Try to load with pandas first for better data handling
        excel_data = {
            'pandas_df': None,
            'sheet': None,
            'headers': [],
            'file_type': os.path.splitext(file_path)[1].lower(),
            'sheet_names': []
        }

        # Load with pandas for data analysis
        if excel_data['file_type'] == '.csv':
            excel_data['pandas_df'] = pd.read_csv(file_path)
        else:
            # For Excel files, get all sheets
            excel_data['pandas_df'] = pd.read_excel(file_path)
            xls = pd.ExcelFile(file_path)
            excel_data['sheet_names'] = xls.sheet_names

        # Extract headers
        if excel_data['pandas_df'] is not None:
            excel_data['headers'] = list(excel_data['pandas_df'].columns)

        # Also load with openpyxl for formula access
        if excel_data['file_type'] != '.csv':
            wb = openpyxl.load_workbook(file_path, data_only=False)
            excel_data['sheet'] = wb.active
            excel_data['workbook'] = wb

        return excel_data
    except Exception as e:
        logging.error(f"Error loading file {file_path}: {e}")
        return None

def normalize_value(value):
    """Normalize different value types for comparison."""
    if pd.isna(value):
        return None

    # Convert to string but handle numeric values carefully
    if isinstance(value, (int, float, Decimal)):
        # Round decimal values to 2 places for financial data
        if isinstance(value, Decimal):
            return value.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        elif isinstance(value, float):
            return round(value, 2)
        return value

    # For dates, convert to ISO format string
    if pd.Timestamp(value) is not pd.NaT:
        try:
            return pd.Timestamp(value).date().isoformat()
        except:
            pass

    # For strings, normalize whitespace and case
    if isinstance(value, str):
        return re.sub(r'\s+', ' ', value).strip().lower()

    return str(value)

def detect_common_accounting_columns(df):
    """Detect common accounting columns in the dataframe for smart matching."""
    column_patterns = {
        'date': ['date', 'transaction_date', 'posting_date', 'invoice_date', 'payment_date'],
        'amount': ['amount', 'transaction_amount', 'debit', 'credit', 'value', 'total'],
        'description': ['description', 'narration', 'particulars', 'details', 'memo'],
        'reference': ['reference', 'ref', 'ref_no', 'check_no', 'invoice_no', 'transaction_id'],
        'account': ['account', 'gl_account', 'account_code', 'account_number'],
    }

    detected_columns = {}

    for col_type, patterns in column_patterns.items():
        for col in df.columns:
            col_lower = str(col).lower()
            if any(pattern in col_lower for pattern in patterns):
                detected_columns[col_type] = col
                break

    return detected_columns

def find_matching_columns(df1, df2):
    """Find automatically matching columns between two dataframes."""
    matching_columns = {}

    # Try exact column name matches first
    common_columns = set(df1.columns).intersection(set(df2.columns))
    for col in common_columns:
        matching_columns[col] = col

    # For columns without direct matches, try to infer based on content or naming patterns
    columns1 = detect_common_accounting_columns(df1)
    columns2 = detect_common_accounting_columns(df2)

    # Match columns of the same type
    for col_type in columns1:
        if col_type in columns2 and columns1[col_type] not in matching_columns:
            matching_columns[columns1[col_type]] = columns2[col_type]

    return matching_columns

def identify_discrepancies(df1, df2, matching_columns=None, tolerance=0.01):
    """
    Identify discrepancies between two dataframes with improved accounting logic.
    Returns a detailed analysis dataframe.
    """
    if matching_columns is None:
        matching_columns = find_matching_columns(df1, df2)

    # If no matching columns found and columns count is the same, try positional matching
    if not matching_columns and len(df1.columns) == len(df2.columns):
        matching_columns = {df1.columns[i]: df2.columns[i] for i in range(len(df1.columns))}

    analysis = {
        'total_rows_df1': len(df1),
        'total_rows_df2': len(df2),
        'matching_rows': 0,
        'missing_in_df2': [],
        'missing_in_df1': [],
        'value_differences': [],
        'duplicates_df1': df1.duplicated().sum(),
        'duplicates_df2': df2.duplicated().sum(),
    }

    # Calculate totals for numerical columns (common accounting practice)
    numeric_totals = {
        'df1': {},
        'df2': {}
    }

    for df_idx, df in [(1, df1), (2, df2)]:
        for col in df.select_dtypes(include=['number']).columns:
            numeric_totals[f'df{df_idx}'][col] = float(df[col].sum())  # Convert to Python float

    analysis['numeric_totals'] = numeric_totals

    # Create comparison dataframe for detailed line-by-line matching
    if matching_columns:
        # For each row in df1, find matching row(s) in df2
        for idx1, row1 in df1.iterrows():
            match_found = False

            # Build filter for df2
            match_condition = pd.Series(True, index=df2.index)
            for col1, col2 in matching_columns.items():
                val1 = normalize_value(row1[col1])

                # Skip None values
                if val1 is None:
                    continue

                # Handle numeric columns with tolerance
                if isinstance(val1, (int, float, Decimal)):
                    lower_bound = float(val1) - tolerance
                    upper_bound = float(val1) + tolerance
                    match_condition &= (df2[col2] >= lower_bound) & (df2[col2] <= upper_bound)
                else:
                    match_condition &= df2[col2].apply(lambda x: normalize_value(x) == val1)

            matching_rows_df2 = df2[match_condition]

            if len(matching_rows_df2) > 0:
                match_found = True
                analysis['matching_rows'] += 1

                # Check for value differences in non-matching columns
                for _, row2 in matching_rows_df2.iterrows():
                    for col1 in df1.columns:
                        if col1 not in matching_columns:
                            continue

                        col2 = matching_columns[col1]
                        val1 = normalize_value(row1[col1])
                        val2 = normalize_value(row2[col2])

                        # If values are numeric, check with tolerance
                        if isinstance(val1, (int, float, Decimal)) and isinstance(val2, (int, float, Decimal)):
                            if abs(float(val1) - float(val2)) > tolerance:
                                analysis['value_differences'].append({
                                    'row_df1': int(idx1) if isinstance(idx1, (np.integer, np.int64)) else idx1,
                                    'row_df2': int(row2.name) if isinstance(row2.name, (np.integer, np.int64)) else row2.name,
                                    'column_df1': str(col1),
                                    'column_df2': str(col2),
                                    'value_df1': float(val1) if isinstance(val1, (np.floating, np.float64)) else val1,
                                    'value_df2': float(val2) if isinstance(val2, (np.floating, np.float64)) else val2,
                                    'difference': float(float(val1) - float(val2))
                                })
                        # Otherwise check for exact match
                        elif val1 != val2:
                            analysis['value_differences'].append({
                                'row_df1': int(idx1) if isinstance(idx1, (np.integer, np.int64)) else idx1,
                                'row_df2': int(row2.name) if isinstance(row2.name, (np.integer, np.int64)) else row2.name,
                                'column_df1': str(col1),
                                'column_df2': str(col2),
                                'value_df1': val1,
                                'value_df2': val2,
                                'difference': 'Non-numeric difference'
                            })

            if not match_found:
                analysis['missing_in_df2'].append({
                    'row': int(idx1) if isinstance(idx1, (np.integer, np.int64)) else idx1,
                    'data': {k: (float(v) if isinstance(v, (np.floating, np.float64)) else 
                                (int(v) if isinstance(v, (np.integer, np.int64)) else v)) 
                             for k, v in row1.to_dict().items()}
                })

        # Find rows in df2 that don't match any in df1
        for idx2, row2 in df2.iterrows():
            match_found = False

            match_condition = pd.Series(True, index=df1.index)
            for col1, col2 in matching_columns.items():
                val2 = normalize_value(row2[col2])

                if val2 is None:
                    continue

                if isinstance(val2, (int, float, Decimal)):
                    lower_bound = float(val2) - tolerance
                    upper_bound = float(val2) + tolerance
                    match_condition &= (df1[col1] >= lower_bound) & (df1[col1] <= upper_bound)
                else:
                    match_condition &= df1[col1].apply(lambda x: normalize_value(x) == val2)

            matching_rows_df1 = df1[match_condition]

            if len(matching_rows_df1) > 0:
                match_found = True

            if not match_found:
                analysis['missing_in_df1'].append({
                    'row': int(idx2) if isinstance(idx2, (np.integer, np.int64)) else idx2,
                    'data': {k: (float(v) if isinstance(v, (np.floating, np.float64)) else 
                                (int(v) if isinstance(v, (np.integer, np.int64)) else v)) 
                             for k, v in row2.to_dict().items()}
                })

    # Convert numpy types to Python native types for safe serialization
    analysis['matching_rows'] = int(analysis['matching_rows'])
    analysis['duplicates_df1'] = int(analysis['duplicates_df1'])
    analysis['duplicates_df2'] = int(analysis['duplicates_df2'])
    
    return analysis

def create_reconciliation_excel(df1, df2, analysis, file_path):
    """Create a detailed reconciliation Excel file."""
    writer = pd.ExcelWriter(file_path, engine='openpyxl')

    # Write original datasets
    df1.to_excel(writer, sheet_name='Sheet1_Original', index=False)
    df2.to_excel(writer, sheet_name='Sheet2_Original', index=False)

    # Create summary sheet
    summary_data = {
        'Metric': [
            'Total Rows in Sheet 1',
            'Total Rows in Sheet 2',
            'Matching Rows',
            'Missing in Sheet 2',
            'Missing in Sheet 1',
            'Rows with Value Differences',
            'Duplicates in Sheet 1',
            'Duplicates in Sheet 2'
        ],
        'Value': [
            analysis['total_rows_df1'],
            analysis['total_rows_df2'],
            analysis['matching_rows'],
            len(analysis['missing_in_df2']),
            len(analysis['missing_in_df1']),
            len(analysis['value_differences']),
            analysis['duplicates_df1'],
            analysis['duplicates_df2']
        ]
    }
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(writer, sheet_name='Reconciliation Summary', index=False)

    # Create detailed differences sheets
    if analysis['value_differences']:
        diff_df = pd.DataFrame(analysis['value_differences'])
        diff_df.to_excel(writer, sheet_name='Value Differences', index=False)

    if analysis['missing_in_df2']:
        missing_df2 = pd.DataFrame([item['data'] for item in analysis['missing_in_df2']])
        missing_df2.to_excel(writer, sheet_name='Missing in Sheet 2', index=False)

    if analysis['missing_in_df1']:
        missing_df1 = pd.DataFrame([item['data'] for item in analysis['missing_in_df1']])
        missing_df1.to_excel(writer, sheet_name='Missing in Sheet 1', index=False)

    # Create numeric totals comparison
    if analysis['numeric_totals']['df1'] or analysis['numeric_totals']['df2']:
        totals_data = []
        all_cols = set(analysis['numeric_totals']['df1'].keys()).union(
                       set(analysis['numeric_totals']['df2'].keys()))

        for col in all_cols:
            total1 = analysis['numeric_totals']['df1'].get(col, 0)
            total2 = analysis['numeric_totals']['df2'].get(col, 0)
            difference = total1 - total2 if col in analysis['numeric_totals']['df1'] and col in analysis['numeric_totals']['df2'] else "N/A"

            totals_data.append({
                'Column': col,
                'Sheet 1 Total': total1 if col in analysis['numeric_totals']['df1'] else "N/A",
                'Sheet 2 Total': total2 if col in analysis['numeric_totals']['df2'] else "N/A",
                'Difference': difference
            })

        totals_df = pd.DataFrame(totals_data)
        totals_df.to_excel(writer, sheet_name='Numeric Totals', index=False)

    writer.close()
    return file_path

def build_prompt_reconciliation(sheet1_data, sheet2_data):
    """Enhanced prompt builder for reconciliation with pandas support."""
    prompt_content = "# Sheet 1 Data Summary:\n"

    # Add sheet metadata
    if sheet1_data.get('pandas_df') is not None:
        df1 = sheet1_data['pandas_df']
        prompt_content += f"Columns: {', '.join(df1.columns.tolist())}\n"
        prompt_content += f"Row count: {len(df1)}\n"

        # Include sample data (first 5 rows)
        prompt_content += "\nSample data (first 5 rows):\n"
        prompt_content += df1.head(5).to_string() + "\n\n"

        # Add summary statistics for numeric columns
        numeric_cols = df1.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            prompt_content += "Numeric column statistics:\n"
            stats = df1[numeric_cols].describe().to_string()
            prompt_content += stats + "\n\n"

        # Detect potential accounting columns
        accounting_cols = detect_common_accounting_columns(df1)
        if accounting_cols:
            prompt_content += "Detected accounting columns:\n"
            for col_type, col_name in accounting_cols.items():
                prompt_content += f"- {col_type}: {col_name}\n"
            prompt_content += "\n"

    prompt_content += "\n# Sheet 2 Data Summary:\n"

    # Add sheet 2 metadata
    if sheet2_data.get('pandas_df') is not None:
        df2 = sheet2_data['pandas_df']
        prompt_content += f"Columns: {', '.join(df2.columns.tolist())}\n"
        prompt_content += f"Row count: {len(df2)}\n"

        # Include sample data (first 5 rows)
        prompt_content += "\nSample data (first 5 rows):\n"
        prompt_content += df2.head(5).to_string() + "\n\n"

        # Add summary statistics for numeric columns
        numeric_cols = df2.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            prompt_content += "Numeric column statistics:\n"
            stats = df2[numeric_cols].describe().to_string()
            prompt_content += stats + "\n\n"

        # Detect potential accounting columns
        accounting_cols = detect_common_accounting_columns(df2)
        if accounting_cols:
            prompt_content += "Detected accounting columns:\n"
            for col_type, col_name in accounting_cols.items():
                prompt_content += f"- {col_type}: {col_name}\n"
            prompt_content += "\n"

    # If we have pandas dataframes for both sheets, add comparison analysis
    if sheet1_data.get('pandas_df') is not None and sheet2_data.get('pandas_df') is not None:
        df1 = sheet1_data['pandas_df']
        df2 = sheet2_data['pandas_df']

        # Try to find matching columns
        matching_columns = find_matching_columns(df1, df2)

        if matching_columns:
            prompt_content += "\n# Preliminary Analysis:\n"
            prompt_content += f"Automatically matched columns: {matching_columns}\n"

            # Perform discrepancy analysis
            analysis = identify_discrepancies(df1, df2, matching_columns)

            prompt_content += f"\nSummary statistics:\n"
            prompt_content += f"- Total rows in Sheet 1: {analysis['total_rows_df1']}\n"
            prompt_content += f"- Total rows in Sheet 2: {analysis['total_rows_df2']}\n"
            prompt_content += f"- Matching rows: {analysis['matching_rows']}\n"
            prompt_content += f"- Missing in Sheet 2: {len(analysis['missing_in_df2'])}\n"
            prompt_content += f"- Missing in Sheet 1: {len(analysis['missing_in_df1'])}\n"
            prompt_content += f"- Value differences: {len(analysis['value_differences'])}\n"

            # Show sample discrepancies
            if analysis['value_differences']:
                prompt_content += "\nSample value differences (first 3):\n"
                for i, diff in enumerate(analysis['value_differences'][:3]):
                    prompt_content += f"{i+1}. Row {diff['row_df1']} in Sheet 1, Row {diff['row_df2']} in Sheet 2\n"
                    prompt_content += f"   Column: {diff['column_df1']} vs {diff['column_df2']}\n"
                    prompt_content += f"   Values: {diff['value_df1']} vs {diff['value_df2']}\n"
                    prompt_content += f"   Difference: {diff['difference']}\n"

            # Show numeric totals
            if analysis['numeric_totals']['df1'] or analysis['numeric_totals']['df2']:
                prompt_content += "\nNumeric column totals comparison:\n"
                all_cols = set(analysis['numeric_totals']['df1'].keys()).union(
                               set(analysis['numeric_totals']['df2'].keys()))

                for col in all_cols:
                    total1 = analysis['numeric_totals']['df1'].get(col, "N/A")
                    total2 = analysis['numeric_totals']['df2'].get(col, "N/A")

                    if col in analysis['numeric_totals']['df1'] and col in analysis['numeric_totals']['df2']:
                        difference = float(total1) - float(total2)
                        prompt_content += f"- {col}: Sheet 1 Total={total1}, Sheet 2 Total={total2}, Difference={difference}\n"
                    else:
                        prompt_content += f"- {col}: Sheet 1 Total={total1}, Sheet 2 Total={total2}\n"

    full_prompt = RECONCILIATION_SYSTEM_PROMPT + "\n\n" + prompt_content

    # Cache the analysis for later use
    session['reconciliation_analysis_data'] = make_session_safe({
        'matching_columns': matching_columns if 'matching_columns' in locals() else None,
        'analysis': analysis if 'analysis' in locals() else None
    })

    logging.info("Enhanced reconciliation prompt built successfully.")
    return full_prompt

def build_prompt(sheet_data):
    """Enhanced prompt builder with pandas integration."""
    prompt_content = ""

    # Add metadata about the sheet
    if 'pandas_df' in sheet_data and sheet_data['pandas_df'] is not None:
        df = sheet_data['pandas_df']
        prompt_content += f"### File Overview\n"
        prompt_content += f"Total rows: {len(df)}\n"
        prompt_content += f"Total columns: {len(df.columns)}\n"
        prompt_content += f"Column names: {', '.join(df.columns.tolist())}\n\n"

        # Add data types
        prompt_content += f"### Data Types\n"
        for col, dtype in df.dtypes.items():
            prompt_content += f"- {col}: {dtype}\n"
        prompt_content += "\n"

        # Add sample data
        prompt_content += f"### Sample Data (First 5 rows)\n"
        prompt_content += df.head(5).to_string() + "\n\n"

        # Add statistics for numeric columns
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            prompt_content += f"### Numeric Column Statistics\n"
            stats = df[numeric_cols].describe().to_string()
            prompt_content += stats + "\n\n"

        # Check for potential accounting columns
        accounting_cols = detect_common_accounting_columns(df)
        if accounting_cols:
            prompt_content += f"### Detected Accounting Columns\n"
            for col_type, col_name in accounting_cols.items():
                prompt_content += f"- {col_type}: {col_name}\n"
            prompt_content += "\n"

    # Also include openpyxl data for formulas
    if 'sheet' in sheet_data and sheet_data['sheet'] is not None:
        sheet = sheet_data['sheet']
        prompt_content += f"### Cell Details\n"

        formulas_found = False
        for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 20), min_col=1, max_col=sheet.max_column):
            for cell in row:
                if cell.value is not None or cell.comment is not None or cell.data_type == 'f':
                    cell_info = ""

                    if cell.data_type == 'f':
                        formulas_found = True
                        cell_info = f"formula '{cell.value}'"
                    elif cell.value is not None:
                        cell_info = f"value '{cell.value}'"
                    else:
                        continue  # Skip cells with no value or formula

                    comment_text = ""
                    if cell.comment:
                        comment_text_raw = cell.comment.text.strip()
                        comment_text_processed = comment_text_raw.replace('\n', ' ')
                        comment_text = f" with comment '{comment_text_processed}'"

                    prompt_content += f"- Cell {cell.coordinate} has {cell_info}{comment_text}.\n"

        # If we found formulas, add extra details for AI analysis
        if formulas_found:
            prompt_content += "\n### Formula Analysis\nThe sheet contains formulas, please analyze their purpose and how they contribute to the accounting or reconciliation function of this spreadsheet.\n\n"

    full_prompt = SYSTEM_PROMPT + "\n\n" + PROMPT_PREFIX + prompt_content + PROMPT_SUFFIX
    logging.info("Enhanced sheet analysis prompt built successfully.")
    return full_prompt

def generate_cache_key(prompt, model_name):
    """Generate a cache key based on the prompt and model."""
    # Use hash to create a stable, fixed-length key
    hash_object = hashlib.md5(prompt.encode())
    return f"{model_name}_{hash_object.hexdigest()}"

def get_explanation_from_gemini(prompt, model_name):
    """Gets explanation from Gemini API with caching."""
    cache_key = generate_cache_key(prompt, model_name)

    # Check if we have a cached response
    if cache_key in CACHE:
        cache_time, cached_response = CACHE[cache_key]
        if time.time() - cache_time < CACHE_TIMEOUT:
            logging.info(f"Using cached response for model: {model_name}")
            return cached_response

    # If no cache hit, call the API
    model = genai.GenerativeModel(model_name)
    try:
        # Set different temperatures based on task type
        if "RECONCILIATION_SYSTEM_PROMPT" in prompt:
            temperature = 0.2  # Lower temperature for factual reconciliation
        elif "FORMULA_SYSTEM_PROMPT" in prompt:
            temperature = 0.3  # Slightly higher for formula creation
        else:
            temperature = 0.1  # Default for general analysis

        response = model.generate_content(
            prompt,
            generation_config=genai.types.GenerationConfig(
                temperature=temperature,
                top_p=0.95,
                top_k=40
            )
        )
        explanation = response.text

        # Cache the response
        CACHE[cache_key] = (time.time(), explanation)

        logging.info(f"New explanation received from Gemini API using model: {model_name}")
        return explanation
    except Exception as e:
        logging.error(f"Error communicating with Gemini API: {e}")
        return None

def get_formula_from_gemini(prompt):
    """Gets formula from Gemini API using formula system prompt."""
    model = genai.GenerativeModel(DEFAULT_MODEL_NAME)
    full_prompt = FORMULA_SYSTEM_PROMPT + "\n\n" + prompt

    # Check cache first
    cache_key = generate_cache_key(full_prompt, DEFAULT_MODEL_NAME)
    if cache_key in CACHE:
        cache_time, cached_response = CACHE[cache_key]
        if time.time() - cache_time < CACHE_TIMEOUT:
            logging.info("Using cached formula response")
            return cached_response

    try:
        response = model.generate_content(
            full_prompt,
            generation_config=genai.types.GenerationConfig(
                temperature=0.3,  # Slightly higher for creative formula solutions
                top_p=0.95,
                top_k=40
            )
        )
        formula_explanation = response.text

        # Cache the response
        CACHE[cache_key] = (time.time(), formula_explanation)

        logging.info("Formula explanation received from Gemini API.")
        return formula_explanation
    except Exception as e:
        logging.error(f"Error communicating with Gemini API for formula: {e}")
        return None

def export_to_docx(explanation, filename=DEFAULT_DOCX_FILENAME, include_title=None):
    """Enhanced DOCX export with better formatting."""
    doc = Document()

    # Add document title if provided
    if include_title:
        title = doc.add_heading(include_title, level=1)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph('')  # Add spacing

    # Process markdown content
    lines = explanation.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]

        # Handle headers
        if line.startswith('# '):
            doc.add_heading(line[2:], level=1)
        elif line.startswith('## '):
            doc.add_heading(line[3:], level=2)
        elif line.startswith('### '):
            doc.add_heading(line[4:], level=3)
        elif line.startswith('#### '):
            doc.add_heading(line[5:], level=4)
        # Handle bullet points
        elif line.startswith('- ') or line.startswith('* '):
            p = doc.add_paragraph()
            p.style = 'List Bullet'
            p.add_run(line[2:])
        # Handle numbered lists
        elif re.match(r'^\d+\.\s', line):
            p = doc.add_paragraph()
            p.style = 'List Number'
            p.add_run(re.sub(r'^\d+\.\s', '', line))
        # Handle code blocks
        elif line.startswith('```'):
            # Find the end of the code block
            start_idx = i + 1
            end_idx = start_idx
            while end_idx < len(lines) and not lines[end_idx].startswith('```'):
                end_idx += 1

            if end_idx < len(lines):
                code_content = '\n'.join(lines[start_idx:end_idx])
                p = doc.add_paragraph(code_content)
                p.style = 'Quote'
                i = end_idx  # Skip to end of code block
        # Regular paragraph
        else:
            # Skip empty lines between paragraphs
            if line.strip() != '' or (i > 0 and lines[i-1].strip() != ''):
                p = doc.add_paragraph(line)

        i += 1

    docx_stream = BytesIO()
    try:
        doc.save(docx_stream)
        docx_stream.seek(0)
        logging.info(f"Content exported to DOCX in memory as {filename}.")
        return docx_stream
    except Exception as e:
        logging.error(f"Error exporting to DOCX: {e}")
        return None

def process_excel_file_async(file, file_path, callback):
    """Process Excel file in a separate thread to avoid blocking the web server."""
    try:
        file.save(file_path)
        sheet_data = load_excel_data(file_path)
        if sheet_data:
            callback(sheet_data)
        else:
            callback(None, "Failed to load Excel data.")
    except Exception as e:
        logging.error(f"Error in async Excel processing: {e}")
        callback(None, f"An error occurred: {e}")
    finally:
        if os.path.exists(file_path):
            os.remove(file_path)

# --- Routes ---
@app.route('/login', methods=['GET', 'POST'])
@limiter.limit("5 per minute")  # Rate limiting for login attempts
def login():
    """Login page with enhanced security."""
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')

        if not username or not password:
            flash('Both username and password are required.', 'error')
            return render_template('login.html')

        user_data = None
        user_id_found = None

        # Find matching user
        for user_id, data in users.items():
            if data['username'] == username:
                user_data = data
                user_id_found = user_id
                break

        if user_data and check_password_hash(user_data['password_hash'], password):
            user = User(
                id=user_id_found,
                username=username,
                password_hash=user_data['password_hash'],
                role=user_data.get('role', 'user')
            )
            login_user(user)

            # Log successful login
            logging.info(f"User '{username}' logged in successfully")
            flash('Logged in successfully.')

            next_page = request.args.get('next')
            return redirect(next_page or url_for('index'))
        else:
            # Log failed login attempt
            logging.warning(f"Failed login attempt for user '{username}'")
            flash('Invalid username or password', 'error')

    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    """Logout route with logging."""
    username = current_user.username
    logout_user()
    logging.info(f"User '{username}' logged out")
    flash('Logged out successfully.')
    return redirect(url_for('login'))

@app.route('/', methods=['GET', 'POST'])
@login_required
def index():
    """Handles the main application logic for Excel sheet explanation."""
    explanation_html = None
    docx_stream = None
    error = None
    model_name = DEFAULT_MODEL_NAME
    processing = False

    if request.method == 'POST':
        if 'excel_file' not in request.files:
            error = 'No file part'
        elif request.files['excel_file'].filename == '':
            error = 'No selected file'
        elif 'excel_file' in request.files and allowed_file(request.files['excel_file'].filename):
            file = request.files['excel_file']
            try:
                filename = sanitize_filename(file.filename)
                unique_filename = generate_unique_filename(filename)
                file_path = os.path.join(UPLOAD_FOLDER, unique_filename)

                selected_model = request.form.get('model_select')
                if selected_model == 'thinking':
                    model_name = THINKING_MODEL_NAME
                elif selected_model == 'premium':
                    model_name = PREMIUM_MODEL_NAME
                else:
                    model_name = DEFAULT_MODEL_NAME

                # Use an in-memory approach instead of saving to disk
                file_content = file.read()
                with open(file_path, 'wb') as f:
                    f.write(file_content)

                # Process the file
                sheet_data = load_excel_data(file_path)
                if sheet_data:
                    prompt = build_prompt(sheet_data)
                    explanation_markdown = get_explanation_from_gemini(prompt, model_name)

                    if explanation_markdown:
                        explanation_html = markdown.markdown(explanation_markdown)
                        session['explanation_markdown'] = explanation_markdown
                        session['current_explanation_html'] = explanation_html
                        session['analyzed_file_name'] = filename
                        # Mark session as modified to ensure it's saved
                        session.modified = True
                    else:
                        error = "Failed to get explanation from Gemini API."
                else:
                    error = "Failed to load Excel data."
            except Exception as e:
                error = f"An error occurred: {e}"
                logging.error(f"Error processing file: {e}", exc_info=True)
            finally:
                if os.path.exists(file_path):
                    os.remove(file_path)
        else:
            error = f'Invalid file type. Allowed types are {", ".join(ALLOWED_EXTENSIONS)}'

    response = make_response(render_template(
        'index.html',
        explanation_html=explanation_html,
        error=error,
        model_name=model_name,
        current_user=current_user,
        processing=processing,
        filename=session.get('analyzed_file_name')
    ))

    # Add cache control headers to prevent browser caching
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route('/export_docx')
@login_required
def export_docx_route():
    """Exports the explanation to DOCX format and allows download."""
    explanation_markdown = session.get('explanation_markdown')
    filename = session.get('analyzed_file_name', 'unknown_file.xlsx')

    if not explanation_markdown:
        flash("No explanation available to export.", "error")
        return redirect(url_for('index'))

    docx_stream = export_to_docx(
        explanation_markdown,
        DEFAULT_DOCX_FILENAME,
        include_title=f"Analysis of {filename}"
    )

    if docx_stream:
        return send_file(
            docx_stream,
            as_attachment=True,
            download_name=f"analysis_{filename.replace('.xlsx', '.docx')}",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        flash("Error exporting to DOCX.", "error")
        return redirect(url_for('index'))

@app.route('/formula_creator', methods=['GET', 'POST'])
@login_required
def formula_creator():
    """Handles the formula creation page with enhanced accounting focus."""
    formula_explanation_html = None
    docx_stream = None
    error = None
    formula_request = ""

    if request.method == 'POST':
        formula_description = request.form.get('formula_description')
        formula_request = formula_description  # Save request for display

        if formula_description:
            # Add accounting context if not already present
            if not any(term in formula_description.lower() for term in ['account', 'reconcil', 'financ', 'ledger', 'balance']):
                formula_description = f"For accounting purposes: {formula_description}"

            formula_explanation_markdown = get_formula_from_gemini(formula_description)

            if formula_explanation_markdown:
                formula_explanation_html = markdown.markdown(formula_explanation_markdown)
                session['formula_explanation_markdown'] = formula_explanation_markdown
                session['formula_request'] = formula_request
                session.modified = True  # Ensure session is saved
            else:
                error = "Failed to get formula explanation from Gemini API."
        else:
            error = "Please enter a description for the formula you need."

    response = make_response(render_template(
        'formula_creator.html',
        formula_explanation_html=formula_explanation_html,
        formula_request=formula_request,
        error=error,
        current_user=current_user
    ))

    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route('/export_formula_docx')
@login_required
def export_formula_docx_route():
    """Exports the formula explanation to DOCX format."""
    formula_explanation_markdown = session.get('formula_explanation_markdown')
    formula_request = session.get('formula_request', 'Excel Formula')

    if not formula_explanation_markdown:
        flash("No formula explanation available to export.", "error")
        return redirect(url_for('formula_creator'))

    docx_stream = export_to_docx(
        formula_explanation_markdown,
        FORMULA_DOCX_FILENAME,
        include_title=f"Excel Formula: {formula_request[:50]}..."
    )

    if docx_stream:
        return send_file(
            docx_stream,
            as_attachment=True,
            download_name=FORMULA_DOCX_FILENAME,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        flash("Error exporting formula explanation to DOCX.", "error")
        return redirect(url_for('formula_creator'))

@app.route('/chat', methods=['GET', 'POST'])
@login_required
def chat():
    """Handles the chat functionality after sheet analysis."""
    explanation_html = session.get('current_explanation_html')
    explanation_markdown = session.get('explanation_markdown')
    chat_history = session.get('chat_history', [])
    user_message = None
    error = None
    
    # Log session data for debugging
    logging.info(f"Session keys: {list(session.keys())}")
    logging.info(f"Has explanation HTML: {'current_explanation_html' in session}")
    logging.info(f"Has explanation markdown: {'explanation_markdown' in session}")

    if not explanation_html or not explanation_markdown:
        flash("Please analyze a spreadsheet first.", "error")
        return redirect(url_for('index'))

    if request.method == 'POST':
        user_message = request.form.get('chat_message')
        if user_message:
            # Create context with recent chat history
            recent_chat = chat_history[-3:] if chat_history else []
            chat_context = ""
            for msg in recent_chat:
                chat_context += f"User: {msg['user']}\nAssistant: {msg['bot']}\n\n"

            prompt_context = (
                f"The analysis of the Excel sheet is:\n\n{explanation_markdown}\n\n"
                f"Recent chat history:\n{chat_context}\n\n"
                f"User's new question: {user_message}\n\n"
                f"You're a financial and accounting expert. Answer the question specifically about the Excel sheet that was analyzed."
            )

            llm_response_markdown = get_explanation_from_gemini(prompt_context, DEFAULT_MODEL_NAME)

            if llm_response_markdown:
                llm_response_html = markdown.markdown(llm_response_markdown)
                chat_history.append({'user': user_message, 'bot': llm_response_markdown})  # Store markdown version
                session['chat_history'] = chat_history
                session.modified = True  # Ensure session is saved
            else:
                error = "Failed to get chat response from Gemini API."
        else:
            error = "Please enter a chat message."

    # Convert markdown to HTML for display
    display_chat_history = []
    for msg in chat_history:
        display_chat_history.append({
            'user': msg['user'],
            'bot': markdown.markdown(msg['bot']) if isinstance(msg['bot'], str) else msg['bot']
        })

    response = make_response(render_template(
        'chat.html',
        explanation_html=explanation_html,
        chat_history=display_chat_history,
        error=error,
        current_user=current_user,
        filename=session.get('analyzed_file_name')
    ))

    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route('/export_chat_docx')
@login_required
def export_chat_docx_route():
    """Exports the chat history to DOCX format."""
    chat_history = session.get('chat_history')
    filename = session.get('analyzed_file_name', 'unknown_file')

    if not chat_history:
        flash("No chat history available to export.", "error")
        return redirect(url_for('chat'))

    # Build a nicely formatted markdown document
    chat_markdown = f"# Chat History for {filename}\n\n"
    for i, message in enumerate(chat_history):
        chat_markdown += f"## Question {i+1}\n\n"
        chat_markdown += f"**User:** {message['user']}\n\n"
        chat_markdown += f"**Assistant:** {message['bot']}\n\n"
        if i < len(chat_history) - 1:
            chat_markdown += "---\n\n"

    docx_stream = export_to_docx(
        chat_markdown,
        CHAT_DOCX_FILENAME,
        include_title=f"Chat History: {filename}"
    )

    if docx_stream:
        download_name = f"chat_{filename.replace('.xlsx', '.docx').replace('.xls', '.docx')}"
        return send_file(
            docx_stream,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        flash("Error exporting chat history to DOCX.", "error")
        return redirect(url_for('chat'))

@app.route('/reconcile', methods=['GET', 'POST'])
@login_required
def reconcile():
    """Handles the accounts reconciliation page and logic."""
    reconciliation_explanation_html = None
    error = None
    model_name = DEFAULT_MODEL_NAME
    processing = False
    file1_name = None
    file2_name = None

    if request.method == 'POST':
        if 'excel_file_1' not in request.files or 'excel_file_2' not in request.files:
            error = 'Need to upload both Sheet 1 and Sheet 2'
        elif request.files['excel_file_1'].filename == '' or request.files['excel_file_2'].filename == '':
            error = 'Both Sheet 1 and Sheet 2 files need to be selected'
        elif 'excel_file_1' in request.files and allowed_file(request.files['excel_file_1'].filename) and 'excel_file_2' in request.files and allowed_file(request.files['excel_file_2'].filename):
            file1 = request.files['excel_file_1']
            file2 = request.files['excel_file_2']
            file1_name = sanitize_filename(file1.filename)
            file2_name = sanitize_filename(file2.filename)

            unique_filename1 = generate_unique_filename(file1_name)
            unique_filename2 = generate_unique_filename(file2_name)

            file_path_1 = os.path.join(UPLOAD_FOLDER, unique_filename1)
            file_path_2 = os.path.join(UPLOAD_FOLDER, unique_filename2)

            try:
                # Save files to disk temporarily
                file1.save(file_path_1)
                file2.save(file_path_2)

                selected_model = request.form.get('model_select')
                if selected_model == 'thinking':
                    model_name = THINKING_MODEL_NAME
                elif selected_model == 'premium':
                    model_name = PREMIUM_MODEL_NAME
                else:
                    model_name = DEFAULT_MODEL_NAME

                # Load the data
                sheet1_data = load_excel_data(file_path_1)
                sheet2_data = load_excel_data(file_path_2)

                if sheet1_data and sheet2_data:
                    # Store original filenames for export
                    session['reconcile_file1_name'] = file1_name
                    session['reconcile_file2_name'] = file2_name

                    # Prepare reconciliation Excel report
                    if sheet1_data.get('pandas_df') is not None and sheet2_data.get('pandas_df') is not None:
                        df1 = sheet1_data['pandas_df']
                        df2 = sheet2_data['pandas_df']

                        # Find matching columns
                        matching_columns = find_matching_columns(df1, df2)

                        # Analyze discrepancies
                        analysis = identify_discrepancies(df1, df2, matching_columns)

                        # Create reconciliation Excel file
                        recon_excel_path = os.path.join(EXPORTS_FOLDER, f"recon_{int(time.time())}.xlsx")
                        create_reconciliation_excel(df1, df2, analysis, recon_excel_path)

                        # Store path for download
                        session['reconciliation_excel_path'] = recon_excel_path

                        # Store matching columns and analysis in session - convert all to JSON-safe types
                        reconciliation_analysis = {
                            'matching_columns': {str(k): str(v) for k, v in matching_columns.items()},
                            'total_rows_1': int(len(df1)),
                            'total_rows_2': int(len(df2)),
                            'matching_rows': int(analysis['matching_rows']),
                            'missing_in_2': int(len(analysis['missing_in_df2'])),
                            'missing_in_1': int(len(analysis['missing_in_df1'])),
                            'value_differences': int(len(analysis['value_differences']))
                        }
                        
                        # Make sure all values are JSON-serializable
                        session['reconciliation_analysis'] = make_session_safe(reconciliation_analysis)

                    # Get AI explanation
                    prompt = build_prompt_reconciliation(sheet1_data, sheet2_data)
                    reconciliation_markdown = get_explanation_from_gemini(prompt, model_name)

                    if reconciliation_markdown:
                        reconciliation_explanation_html = markdown.markdown(reconciliation_markdown)
                        session['reconciliation_explanation_markdown'] = reconciliation_markdown
                        session.modified = True  # Ensure session is saved
                    else:
                        error = "Failed to get reconciliation explanation from Gemini API."
                else:
                    error = "Failed to load data from one or both Excel files."
            except Exception as e:
                error = f"An error occurred during reconciliation: {e}"
                logging.error(f"Reconciliation error: {e}", exc_info=True)
            finally:
                if os.path.exists(file_path_1):
                    os.remove(file_path_1)
                if os.path.exists(file_path_2):
                    os.remove(file_path_2)
        else:
            error = f'Invalid file types. Allowed types are {", ".join(ALLOWED_EXTENSIONS)} for both sheets.'

    # Get analysis from session if available
    analysis = session.get('reconciliation_analysis', {})

    response = make_response(render_template(
        'reconcile.html',
        reconciliation_explanation_html=reconciliation_explanation_html,
        error=error,
        current_user=current_user,
        model_name=model_name,
        processing=processing,
        analysis=analysis,
        file1_name=file1_name or session.get('reconcile_file1_name'),
        file2_name=file2_name or session.get('reconcile_file2_name')
    ))

    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route('/export_reconciliation_docx')
@login_required
def export_reconciliation_docx_route():
    """Exports the reconciliation explanation to DOCX format."""
    reconciliation_explanation_markdown = session.get('reconciliation_explanation_markdown')
    file1_name = session.get('reconcile_file1_name', 'Sheet1')
    file2_name = session.get('reconcile_file2_name', 'Sheet2')

    if not reconciliation_explanation_markdown:
        flash("No reconciliation explanation available to export.", "error")
        return redirect(url_for('reconcile'))

    docx_stream = export_to_docx(
        reconciliation_explanation_markdown,
        RECONCILIATION_DOCX_FILENAME,
        include_title=f"Reconciliation Report: {file1_name} vs {file2_name}"
    )

    if docx_stream:
        return send_file(
            docx_stream,
            as_attachment=True,
            download_name=f"reconciliation_{int(time.time())}.docx",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        flash("Error exporting reconciliation explanation to DOCX.", "error")
        return redirect(url_for('reconcile'))

@app.route('/export_reconciliation_excel')
@login_required
def export_reconciliation_excel():
    """Exports the detailed reconciliation data to Excel format."""
    recon_excel_path = session.get('reconciliation_excel_path')

    if not recon_excel_path or not os.path.exists(recon_excel_path):
        flash("No reconciliation Excel file available to export.", "error")
        return redirect(url_for('reconcile'))

    try:
        return send_file(
            recon_excel_path,
            as_attachment=True,
            download_name=f"reconciliation_detailed_{int(time.time())}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        logging.error(f"Error exporting reconciliation Excel: {e}")
        flash("Error exporting reconciliation Excel file.", "error")
        return redirect(url_for('reconcile'))

@app.route('/api/status')
@login_required
def api_status():
    """API endpoint for checking the status of the API connection."""
    if configure_api():
        return jsonify({
            'status': 'ok',
            'message': 'API connection successful',
            'user': current_user.username
        })
    else:
        return jsonify({
            'status': 'error',
            'message': 'API connection failed'
        }), 500

# Debugging endpoint for session data
@app.route('/debug-session')
@login_required
def debug_session():
    """View session data for debugging."""
    if not current_user.is_admin():
        flash("You don't have permission to access this page.", "error")
        return redirect(url_for('index'))
        
    return jsonify({
        'session_keys': list(session.keys()),
        'has_explanation_html': 'current_explanation_html' in session,
        'has_explanation_markdown': 'explanation_markdown' in session,
        'user': current_user.username
    })

# Error handlers
@app.errorhandler(404)
def page_not_found(e):
    return render_template('error.html', error="Page not found"), 404

@app.errorhandler(500)
def server_error(e):
    logging.error(f"500 error: {str(e)}", exc_info=True)
    return render_template('error.html', error="Internal server error"), 500

# Admin dashboard
@app.route('/admin')
@login_required
def admin_dashboard():
    """Admin dashboard for monitoring the application."""
    if not current_user.is_admin():
        flash("You don't have permission to access the admin dashboard.", "error")
        return redirect(url_for('index'))

    # Collect system information
    system_info = {
        'api_status': configure_api(),
        'cache_size': len(CACHE),
        'user_count': len(users),
        'upload_dir': UPLOAD_FOLDER,
        'exports_dir': EXPORTS_FOLDER
    }

    return render_template('admin.html', system_info=system_info, current_user=current_user)

if __name__ == '__main__':
    if configure_api():
        # Don't run in debug mode in production
        debug_mode = os.environ.get('FLASK_ENV') == 'development'
        app.run(
            debug=debug_mode,
            host="0.0.0.0",
            port=int(os.environ.get("PORT", 8000))
        )
    else:
        print("Failed to configure API. Exiting.")
