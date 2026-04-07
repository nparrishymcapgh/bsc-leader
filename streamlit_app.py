import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import json
import uuid
import smtplib
from email.message import EmailMessage
from urllib.parse import urlencode
import time
from functools import lru_cache

# ============================================================================
# PAGE CONFIG
# ============================================================================
st.set_page_config(
    page_title="Leader Level Balanced Score Card",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Global debug variables
load_responses_debug = []

# Minimal CSS to avoid breaking Streamlit's UI rendering
custom_css = """
<style>
    /* Only style specific elements without using !important excessively */
    /* Let Streamlit handle its own theme */
    .stButton>button {
        background-color: #006B6B;
        color: white;
        border: none;
        border-radius: 4px;
        font-weight: 500;
    }
    
    .stButton>button:hover {
        background-color: #005050;
    }
    
    code {
        background-color: #f0f0f0;
        border-radius: 4px;
        padding: 2px 4px;
    }
    
    /* Expander styling only */
    .streamlit-expanderHeader {
        font-weight: bold;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ============================================================================
# GOOGLE SHEETS CONFIG
# ============================================================================
GOOGLE_SHEET_ID = "1DfYJwlKy01G0tcZ11FUen4fPjSnK9vWT33a2sKfRKko"
EMPLOYEES_TAB = "Employees"
QUESTIONS_TAB = "Questions"
RESPONSES_TAB = "Responses"

# Optional Streamlit secrets values:
# [gcp_service_account]
# ... your service account JSON fields ...
# [smtp]
# server = "smtp.example.com"
# port = 587
# username = "user@example.com"
# password = "secret"
# from_email = "no-reply@example.com"
# [app]
# url = "https://your-app-url.streamlit.app"

# ============================================================================
# GOOGLE SHEETS UTILITIES
# ============================================================================

@st.cache_resource
def get_spreadsheet():
    try:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        client = gspread.authorize(creds)
        return client.open_by_key(GOOGLE_SHEET_ID)
    except KeyError:
        st.error("Google service account credentials not found in secrets.")
        st.info("Please configure your `.streamlit/secrets.toml` file with the service account credentials.")
        st.stop()
    except PermissionError:
        st.error("Permission denied accessing Google Sheet.")
        st.info("Make sure:")
        st.info("1. The service account email has been shared with the Google Sheet")
        st.info("2. The GOOGLE_SHEET_ID is correct")
        st.info("3. The service account has 'Editor' access to the sheet")
        st.stop()
    except Exception as e:
        st.error(f"Error connecting to Google Sheets: {e}")
        st.info("Please check your service account configuration and Google Sheet permissions.")
        st.stop()

@st.cache_data(ttl=3600)
@lru_cache(maxsize=None)
def load_sheet(tab_name):
    spreadsheet = get_spreadsheet()
    try:
        worksheet = spreadsheet.worksheet(tab_name)
        records = worksheet.get_all_records()
        return pd.DataFrame(records)
    except gspread.exceptions.WorksheetNotFound:
        return pd.DataFrame()
    except Exception as exc:
        st.error(f"Unable to load `{tab_name}` sheet: {exc}")
        return pd.DataFrame()

@st.cache_data(ttl=60)
@lru_cache(maxsize=None)
def load_responses():
    global load_responses_debug
    debug_messages = []
    debug_messages.append("DEBUG: load_responses() called")
    try:
        spreadsheet = get_spreadsheet()
        debug_messages.append(f"DEBUG: Got spreadsheet: {spreadsheet.title if spreadsheet else 'None'}")
        worksheet = ensure_responses_sheet(spreadsheet)
        debug_messages.append(f"DEBUG: Got worksheet: {worksheet.title if worksheet else 'None'}")
        records = worksheet.get_all_records()
        debug_messages.append(f"DEBUG: worksheet.get_all_records() returned {len(records)} records")
        if records:
            debug_messages.append(f"DEBUG: First record keys: {list(records[0].keys()) if records else 'None'}")
        df = pd.DataFrame(records)
        debug_messages.append(f"DEBUG: DataFrame shape after creation: {df.shape}")
        # Ensure all expected columns exist, even if sheet is empty
        expected_columns = [
            "response_id", "created_at", "updated_at", "manager_email", "manager_name",
            "employee_id", "employee_name", "employee_email", "branch", "dept",
            "job_title", "executive_email", "questions_score", "number_of_nos",
            "responses", "comments", "employee_agree", "manager_agree", "executive_agree",
            "employee_agree_ts", "manager_agree_ts", "executive_agree_ts",
            "status", "employee_token", "manager_token", "executive_token"
        ]
        for col in expected_columns:
            if col not in df.columns:
                df[col] = ""
        debug_messages.append(f"DEBUG: Final DataFrame shape: {df.shape}")
        
        # Store debug messages in a global variable that can be accessed by the UI
        load_responses_debug = debug_messages
        
        return df
    except Exception as e:
        debug_messages.append(f"DEBUG: load_responses() error: {e}")
        import traceback
        debug_messages.append(f"DEBUG: Traceback: {traceback.format_exc()}")
        
        load_responses_debug = debug_messages
        
        return pd.DataFrame()

def ensure_responses_sheet(spreadsheet):
    headers = [
        "response_id",
        "created_at",
        "updated_at",
        "manager_email",
        "manager_name",
        "employee_id",
        "employee_name",
        "employee_email",
        "branch",
        "dept",
        "job_title",
        "executive_email",
        "questions_score",
        "number_of_nos",
        "responses",
        "comments",
        "employee_agree",
        "manager_agree",
        "executive_agree",
        "employee_agree_ts",
        "manager_agree_ts",
        "executive_agree_ts",
        "status",
        "employee_token",
        "manager_token",
        "executive_token"
    ]
    try:
        worksheet = spreadsheet.worksheet(RESPONSES_TAB)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(RESPONSES_TAB, rows=1000, cols=len(headers))
        worksheet.append_row(headers)
    return worksheet

# ============================================================================
# EMAIL UTILITIES
# ============================================================================

def get_app_url():
    return st.secrets.get("app", {}).get("url", "").rstrip("/")

def build_action_link(response_id, token, action):
    base = get_app_url()
    if not base:
        return ""
    params = {
        "response_id": response_id,
        "token": token,
        "action": action
    }
    return f"{base}/?{urlencode(params)}"


def send_email(subject, html_body, recipient):
    smtp_config = st.secrets.get("smtp", {})
    if not smtp_config:
        return False

    try:
        message = EmailMessage()
        message["Subject"] = subject
        message["From"] = smtp_config.get("from_email", smtp_config.get("username"))
        message["To"] = recipient
        message.set_content("Please view this message in an HTML-capable email client.")
        message.add_alternative(html_body, subtype="html")

        server = smtplib.SMTP(smtp_config["server"], int(smtp_config.get("port", 587)))
        server.starttls()
        server.login(smtp_config["username"], smtp_config["password"])
        server.send_message(message)
        server.quit()
        return True
    except Exception as exc:
        st.error(f"Email send failed: {exc}")
        return False


def format_scorecard_summary(employee, question_rows, answers):
    lines = []
    lines.append(f"<p><strong>Employee:</strong> {employee['name']} ({employee['ID']})</p>")
    lines.append(f"<p><strong>Branch:</strong> {employee.get('branch', '')} &nbsp;&nbsp; <strong>Dept:</strong> {employee.get('dept', '')} &nbsp;&nbsp; <strong>Job Title:</strong> {employee.get('job_title', '')}</p>")
    lines.append("<table border='0' cellpadding='6' cellspacing='0' style='border-collapse:collapse;'>")
    lines.append("<tr><th align='left'>Section</th><th align='left'>Question</th><th align='left'>Answer</th></tr>")
    for _, row in question_rows.iterrows():
        answer = answers.get(str(row['ID']), "")
        lines.append(f"<tr><td>{row['question_section']}</td><td>{row['question']}</td><td>{answer}</td></tr>")
    lines.append("</table>")
    return "".join(lines)


def format_email_body(subject, employee, question_rows, answers, stage, approve_link, reject_link, comment=""):
    score_answers = [int(v) for qid, v in answers.items() if v in ["1", "2", "3"]]
    questions_score = int(round(sum(score_answers) / len(score_answers) * 100)) if score_answers else 0
    number_of_nos = sum(1 for qid, v in answers.items() if v == "No")

    html = [f"<h2>{subject}</h2>"]
    html.append(format_scorecard_summary(employee, question_rows, answers))
    html.append(f"<p><strong>Score:</strong> {questions_score}</p>")
    html.append(f"<p><strong>Number of No answers:</strong> {number_of_nos}</p>")

    if comment.strip():
        html.append(f"<p><strong>Manager Comments:</strong></p><blockquote style='border-left: 4px solid #006B6B; padding-left: 10px; margin: 10px 0; font-style: italic;'>{comment.strip()}</blockquote>")

    html.append("<p>Please approve or reject using the links below.</p>")
    if approve_link:
        html.append(f"<p><a href=\"{approve_link}\" style=\"background:#006B6B;color:white;padding:10px 14px;text-decoration:none;border-radius:4px;\">Approve</a></p>")
    if reject_link:
        html.append(f"<p><a href=\"{reject_link}\" style=\"background:#A80000;color:white;padding:10px 14px;text-decoration:none;border-radius:4px;\">Reject</a></p>")
    html.append("<p>If you have any questions, please contact your manager.</p>")
    return "".join(html)

# ============================================================================
# RESPONSE OPERATIONS
# ============================================================================

def find_response_by_id(response_id):
    # Clear cache to ensure fresh data
    load_responses.clear()
    spreadsheet = get_spreadsheet()
    worksheet = ensure_responses_sheet(spreadsheet)
    records = worksheet.get_all_records()
    df = pd.DataFrame(records)
    # Ensure all expected columns exist, even if sheet is empty
    expected_columns = [
        "response_id", "created_at", "updated_at", "manager_email", "manager_name",
        "employee_id", "employee_name", "employee_email", "branch", "dept",
        "job_title", "executive_email", "questions_score", "number_of_nos",
        "responses", "comments", "employee_agree", "manager_agree", "executive_agree",
        "employee_agree_ts", "manager_agree_ts", "executive_agree_ts",
        "status", "employee_token", "manager_token", "executive_token"
    ]
    for col in expected_columns:
        if col not in df.columns:
            df[col] = ""
    if df.empty:
        return None, None, None
    matches = df[df['response_id'] == response_id]
    if matches.empty:
        return None, None, None
    row = matches.iloc[0]
    row_index = matches.index[0] + 2  # account for header row
    return row.to_dict(), row_index, df


def update_response(response_id, updates):
    try:
        spreadsheet = get_spreadsheet()
        worksheet = ensure_responses_sheet(spreadsheet)
        records = worksheet.get_all_records()
        df = pd.DataFrame(records)
        # Ensure all expected columns exist
        expected_columns = [
            "response_id", "created_at", "updated_at", "manager_email", "manager_name",
            "employee_id", "employee_name", "employee_email", "branch", "dept",
            "job_title", "executive_email", "questions_score", "number_of_nos",
            "responses", "comments", "employee_agree", "manager_agree", "executive_agree",
            "employee_agree_ts", "manager_agree_ts", "executive_agree_ts",
            "status", "employee_token", "manager_token", "executive_token"
        ]
        for col in expected_columns:
            if col not in df.columns:
                df[col] = ""
        
        if df.empty:
            print(f"DEBUG: update_response - No records found")
            return False

        match = df[df['response_id'] == response_id]
        if match.empty:
            print(f"DEBUG: update_response - No match found for response_id: {response_id}")
            return False

        row_index = match.index[0] + 2
        header = worksheet.row_values(1)
        row_values = worksheet.row_values(row_index)
        row_data = {header[i]: row_values[i] if i < len(row_values) else "" for i in range(len(header))}

        print(f"DEBUG: update_response - row_index: {row_index}, header length: {len(header)}, row_values length: {len(row_values)}")
        print(f"DEBUG: update_response - Before update: status = {row_data.get('status', 'N/A')}")
        
        for key, value in updates.items():
            if key not in header:
                print(f"DEBUG: update_response - Adding new column: {key}")
                header.append(key)
                # Update the worksheet header row with the new column
                worksheet.update(f"{chr(64 + len(header))}{1}", [[key]])
                row_data[key] = value
            else:
                print(f"DEBUG: update_response - Updating existing column: {key} = {value}")
            row_data[key] = value

        ordered_row = [row_data.get(col, "") for col in header]
        print(f"DEBUG: update_response - Final ordered_row length: {len(ordered_row)}, header length: {len(header)}")
        update_range = f"A{row_index}:{chr(64 + len(header))}{row_index}"
        print(f"DEBUG: update_response - Updating range: {update_range}")
        worksheet.update(update_range, [ordered_row])
        load_responses.clear()
        load_sheet.clear()  # Clear sheet cache in case columns were added
        
        # Verify the update
        records_after = worksheet.get_all_records()
        df_after = pd.DataFrame(records_after)
        match_after = df_after[df_after['response_id'] == response_id]
        if not match_after.empty:
            new_status = match_after.iloc[0].get('status', 'N/A')
            print(f"DEBUG: update_response - After update: status = {new_status}")
        
        return True
    except Exception as e:
        print(f"DEBUG: update_response - Error: {e}")
        import traceback
        traceback.print_exc()
        return False


def append_response(row):
    spreadsheet = get_spreadsheet()
    worksheet = ensure_responses_sheet(spreadsheet)
    header = worksheet.row_values(1)
    ordered_row = [row.get(col, "") for col in header]
    worksheet.append_row(ordered_row)
    load_responses.clear()
    load_sheet.clear()  # Clear sheet cache when new data is added


def create_response_entry(manager, employee, answers, comment=""):
    score_answers = [int(v) for qid, v in answers.items() if v in ["1", "2", "3"]]
    questions_score = int(round(sum(score_answers) / len(score_answers) * 100)) if score_answers else 0
    number_of_nos = sum(1 for qid, v in answers.items() if v == "No")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = {
        "response_id": str(uuid.uuid4()),
        "created_at": now,
        "updated_at": now,
        "manager_email": manager['manager_email'],
        "manager_name": manager.get('manager_name', ""),
        "employee_id": str(employee['ID']),
        "employee_name": employee['name'],
        "employee_email": employee['email'],
        "branch": employee.get('branch', ""),
        "dept": employee.get('dept', ""),
        "job_title": employee.get('job_title', ""),
        "executive_email": employee.get('executive_email', ""),
        "questions_score": questions_score,
        "number_of_nos": number_of_nos,
        "responses": json.dumps(answers),
        "comments": comment.strip(),
        "employee_agree": "",
        "manager_agree": "",
        "executive_agree": "",
        "employee_agree_ts": "",
        "manager_agree_ts": "",
        "executive_agree_ts": "",
        "status": "Pending Employee",
        "employee_token": str(uuid.uuid4()),
        "manager_token": str(uuid.uuid4()),
        "executive_token": str(uuid.uuid4())
    }
    return entry


def load_employee_questions():
    df = load_sheet(QUESTIONS_TAB)
    if df.empty:
        st.warning("Questions sheet is empty or missing. Please check the Google Sheet.")
    return df

# ============================================================================
# APPROVAL WORKFLOW
# ============================================================================

def get_stage_links(response):
    approve_link = None
    reject_link = None
    if response['status'] == 'Pending Employee':
        approve_link = build_action_link(response['response_id'], response['employee_token'], 'employee_approve')
        reject_link = build_action_link(response['response_id'], response['employee_token'], 'employee_reject')
    elif response['status'] == 'Pending Manager':
        approve_link = build_action_link(response['response_id'], response['manager_token'], 'manager_approve')
        reject_link = build_action_link(response['response_id'], response['manager_token'], 'manager_reject')
    elif response['status'] == 'Pending Executive':
        approve_link = build_action_link(response['response_id'], response['executive_token'], 'executive_approve')
        reject_link = build_action_link(response['response_id'], response['executive_token'], 'executive_reject')
    return approve_link, reject_link


def send_stage_email(response, stage):
    question_rows = load_employee_questions()
    questions_for_email = question_rows.copy()
    try:
        answers = json.loads(response['responses']) if isinstance(response['responses'], str) else response['responses']
    except Exception:
        answers = {}

    approve_link, reject_link = get_stage_links(response)
    stage_label = {
        'employee': 'Employee Verification',
        'manager': 'Manager Approval',
        'executive': 'Executive Approval',
        'rejected': 'Review Required'
    }.get(stage, stage)

    subject = f"Leader Level Balanced Score Card - {stage_label}"
    
    # Create employee object from response data for email formatting
    employee = {
        'name': response.get('employee_name', ''),
        'ID': response.get('employee_id', ''),
        'branch': response.get('branch', ''),
        'dept': response.get('dept', ''),
        'job_title': response.get('job_title', '')
    }
    
    body = format_email_body(subject, employee, questions_for_email, answers, stage, approve_link, reject_link, response.get('comments', ''))
    recipient = ''

    if stage == 'employee':
        recipient = response.get('employee_email', '')
    elif stage == 'manager':
        recipient = response.get('manager_email', '')
    elif stage == 'executive':
        recipient = response.get('executive_email', '')
    else:
        recipient = response.get('manager_email', '')

    # Only send email if recipient is valid
    if recipient and '@' in recipient:
        success = send_email(subject, body, recipient)
        return success, recipient, body
    else:
        # No valid recipient, return success=False but don't show error for missing executive emails
        return False, recipient, body


def process_action(action, response_id, token):
    try:
        response, _, _ = find_response_by_id(response_id)
        if not response:
            st.error("Approval record not found. Please check that the link is correct and the record exists.")
            st.info(f"Debug info: Looking for response_id: {response_id}")
            return

        st.info(f"Found record for {response.get('employee_name', 'Unknown')} - Status: {response['status']}")
        if response.get('comments', '').strip():
            st.info(f"Comments present: {len(response['comments'])} characters")
        else:
            st.info("No comments found in record")

        valid = False
        stage_email = None
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        updates = {"updated_at": now}

        if action == 'employee_approve' and response['status'] == 'Pending Employee' and token == response.get('employee_token'):
            valid = True
            updates['employee_agree'] = 'Yes'
            updates['employee_agree_ts'] = now
            updates['status'] = 'Pending Manager'
            stage_email = 'manager'
        elif action == 'employee_reject' and response['status'] == 'Pending Employee' and token == response.get('employee_token'):
            valid = True
            updates['employee_agree'] = 'No'
            updates['employee_agree_ts'] = now
            updates['status'] = 'Rejected by Employee'
            stage_email = 'rejected'
        elif action == 'manager_approve' and (response['status'] == 'Pending Employee' or response['status'] == 'Pending Manager') and token == response.get('manager_token'):
            valid = True
            updates['manager_agree'] = 'Yes'
            updates['manager_agree_ts'] = now
            if response['status'] == 'Pending Employee':
                # If employee hasn't responded yet, mark as manager-approved and pending executive
                updates['status'] = 'Pending Executive'
            else:
                # If employee already approved, just move to executive
                updates['status'] = 'Pending Executive'
            stage_email = 'executive'
        elif action == 'manager_reject' and (response['status'] == 'Pending Employee' or response['status'] == 'Pending Manager') and token == response.get('manager_token'):
            valid = True
            updates['manager_agree'] = 'No'
            updates['manager_agree_ts'] = now
            updates['status'] = 'Rejected by Manager'
            stage_email = 'rejected'
        elif action == 'executive_approve' and response['status'] == 'Pending Executive' and token == response.get('executive_token'):
            valid = True
            updates['executive_agree'] = 'Yes'
            updates['executive_agree_ts'] = now
            updates['status'] = 'Approved'
            # No stage_email for final approval
        elif action == 'executive_reject' and response['status'] == 'Pending Executive' and token == response.get('executive_token'):
            valid = True
            updates['executive_agree'] = 'No'
            updates['executive_agree_ts'] = now
            updates['status'] = 'Rejected by Executive'
            stage_email = 'rejected'

        if not valid:
            st.error("This approval link is invalid or the action is not allowed.")
            return

        if update_response(response_id, updates):
            st.success("Thank you. The scorecard status has been updated.")
            if stage_email:
                sent, recipient, body = send_stage_email({**response, **updates}, stage_email)
                if sent:
                    st.info(f"Notification email sent to {recipient}.")
                else:
                    if recipient and '@' in recipient:
                        st.info("Email could not be sent. Please use the app URL or email preview instead.")
                    else:
                        st.info("No email recipient configured for this stage.")
        else:
            st.error("Unable to update the response record. Please try again or contact support.")
    except Exception as e:
        st.error(f"An error occurred while processing the approval: {e}")
        st.info("Please try again or contact support if the issue persists.")

# ============================================================================
# MAIN UI
# ============================================================================

st.title("Leader Level Balanced Score Card")

# Check if secrets are configured
if "gcp_service_account" not in st.secrets:
    st.error("Google service account credentials not configured.")
    st.info("**Setup Required:**")
    st.info("1. Create a Google Cloud service account and download the JSON key")
    st.info("2. Share your Google Sheet with the service account email")
    st.info("3. Add the credentials to `.streamlit/secrets.toml`")
    st.info("See README.md for detailed setup instructions.")
    st.stop()

query_params = st.query_params
action = query_params.get('action')
response_id = query_params.get('response_id')
token = query_params.get('token')
debug_mode = query_params.get('debug') == 'connection'

if action and response_id and token:
    st.info("Processing approval link...")
    process_action(action, response_id, token)
    st.stop()

# Debug mode bypasses login and shows all data
if debug_mode:
    st.session_state.logged_in = True
    st.session_state.manager_email = 'debug@mode.com'
    st.session_state.manager_name = 'Debug User'
    st.session_state.data_loaded = getattr(st.session_state, 'data_loaded', False)
    if not st.session_state.data_loaded:
        with st.spinner("Loading data..."):
            employees_df = load_sheet(EMPLOYEES_TAB)
            questions_df = load_sheet(QUESTIONS_TAB)
            responses_df = load_responses()
            st.session_state.employees_df = employees_df
            st.session_state.questions_df = questions_df
            st.session_state.responses_df = responses_df
            st.session_state.data_loaded = True
    responses_df = load_responses()
    responses_df['employee_id'] = responses_df['employee_id'].astype(str)
    st.sidebar.warning("🔧 DEBUG MODE - Showing all data")
    st.sidebar.markdown("**Debug Info:**")
    st.sidebar.write(f"Responses: {len(responses_df)} records")
    st.sidebar.write(f"Employees: {len(st.session_state.employees_df)} records")
    st.sidebar.write(f"Questions: {len(st.session_state.questions_df)} records")

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.manager_email = ''
    st.session_state.manager_name = ''
    st.session_state.data_loaded = False

if not st.session_state.data_loaded:
    with st.spinner("Loading data..."):
        employees_df = load_sheet(EMPLOYEES_TAB)
        questions_df = load_sheet(QUESTIONS_TAB)
        responses_df = load_responses()
        st.session_state.employees_df = employees_df
        st.session_state.questions_df = questions_df
        st.session_state.responses_df = responses_df
        st.session_state.data_loaded = True

if not st.session_state.logged_in:
    st.subheader("🔐 Manager Login")
    login_email = st.text_input("Enter your manager email:", placeholder="manager@example.com").strip().lower()
    if st.button("Login", type='primary'):
        if login_email and login_email in st.session_state.employees_df['manager_email'].astype(str).str.lower().values:
            manager_row = st.session_state.employees_df[st.session_state.employees_df['manager_email'].astype(str).str.lower() == login_email].iloc[0]
            st.session_state.logged_in = True
            st.session_state.manager_email = login_email
            st.session_state.manager_name = manager_row.get('manager_name', login_email)
            st.rerun()
        else:
            st.error("Manager email not found. Please enter an email listed in the Employees sheet.")
    st.stop()

st.sidebar.markdown(f"**Signed in as:** {st.session_state.manager_name} ({st.session_state.manager_email})")
if st.sidebar.button("Logout"):
    st.session_state.logged_in = False
    st.session_state.manager_email = ''
    st.session_state.manager_name = ''
    st.rerun()

if debug_mode:
    manager_employees = st.session_state.employees_df  # Show all employees in debug mode
else:
    manager_employees = st.session_state.employees_df[st.session_state.employees_df['manager_email'].astype(str).str.lower() == st.session_state.manager_email]
    if manager_employees.empty:
        st.warning("No employees found for this manager email.")
        st.stop()

responses_df = load_responses()
responses_df['employee_id'] = responses_df['employee_id'].astype(str)

st.subheader("📋 Manager Dashboard")

tab_new, tab_status = st.tabs(["Submit Scorecard", "Scorecard Status"])

with tab_new:
    st.markdown("### Submit a new balanced score card")
    
    # Filter out employees that have already been reviewed
    reviewed_employee_ids = set()
    if not responses_df.empty:
        if debug_mode:
            # In debug mode, don't filter by manager - show all submissions
            reviewed_employee_ids = set(responses_df['employee_id'].astype(str).unique())
        else:
            manager_submissions = responses_df[responses_df['manager_email'].astype(str).str.lower() == st.session_state.manager_email]
            reviewed_employee_ids = set(manager_submissions['employee_id'].astype(str).unique())
    
    if debug_mode:
        # In debug mode, show all employees
        available_employees = st.session_state.employees_df[~st.session_state.employees_df['ID'].astype(str).isin(reviewed_employee_ids)]
        st.info("🔧 DEBUG MODE: Showing all employees (not filtered by manager)")
    else:
        available_employees = manager_employees[~manager_employees['ID'].astype(str).isin(reviewed_employee_ids)]
    
    if available_employees.empty:
        if debug_mode:
            st.success("🎉 **All employees in the system have been reviewed!**")
        else:
            st.success("🎉 **All employees under your supervision have been reviewed!**")
        st.info("You have successfully completed performance reviews for all your direct reports.")
        st.stop()
    
    selected_employee_id = st.selectbox(
        "Select employee to rate",
        available_employees['ID'].astype(str).tolist(),
        format_func=lambda eid: f"{available_employees[available_employees['ID'].astype(str) == eid].iloc[0]['name']} ({eid})"
    )
    selected_employee = available_employees[available_employees['ID'].astype(str) == selected_employee_id].iloc[0]
    st.write(f"Employee: {selected_employee['name']} | Branch: {selected_employee.get('branch', '')} | Dept: {selected_employee.get('dept', '')}")
    st.write(f"Title: {selected_employee.get('job_title', '')} | Executive: {selected_employee.get('executive_email', '')}")
    st.divider()

    questions_df = st.session_state.questions_df
    if questions_df.empty:
        st.warning("The Questions sheet is empty. Please add questions to the Google Sheet.")
    else:
        questions_df = questions_df.fillna("")
        sections = questions_df['question_section'].astype(str).fillna('General').unique().tolist()
        answers = {}

        # Live score calculation
        score_questions = questions_df[questions_df['type'].astype(str).str.strip().str.lower() == 'score']
        total_score_questions = len(score_questions)

        for section in sections:
            section_rows = questions_df[questions_df['question_section'].astype(str) == section]
            header_text = section_rows['header'].astype(str).fillna('').iloc[0]
            if header_text:
                st.markdown(f"### {header_text}")
            for _, question in section_rows.iterrows():
                key = f"q_{selected_employee_id}_{question['ID']}"
                if str(question['type']).strip().lower() == 'score':
                    answers[str(question['ID'])] = st.radio(
                        question['question'],
                        options=['1', '2', '3'],
                        key=key,
                        index=0
                    )
                else:
                    answers[str(question['ID'])] = st.radio(
                        question['question'],
                        options=['Yes', 'No'],
                        key=key,
                        index=0
                    )
            st.divider()

        # Calculate and display current score
        answered_score_questions = [qid for qid, val in answers.items() if val in ['1', '2', '3']]
        current_score = 0
        if answered_score_questions:
            score_values = [int(answers[qid]) for qid in answered_score_questions]
            current_score = int(round(sum(score_values) / len(score_values) * 100)) if score_values else 0

        # Count "No" answers
        no_answers = sum(1 for qid, val in answers.items() if val == 'No')

        # Display current progress
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Current Score", f"{current_score}/100")
        with col2:
            answered = len([v for v in answers.values() if v not in [None, '']])
            total_questions = len(questions_df)
            st.metric("Questions Answered", f"{answered}/{total_questions}")
        with col3:
            st.metric("No Answers", no_answers)

        st.divider()

        # Comment field
        st.markdown("### 📝 Additional Comments (Optional)")
        manager_comment = st.text_area(
            "Add any additional comments or notes about this employee:",
            height=100,
            placeholder="Enter your comments here...",
            help="These comments will be included in the email sent to the employee but won't affect the score."
        )

        st.divider()

        if st.button("Submit Scorecard", type='primary'):
            missing = [qid for qid, value in answers.items() if value in [None, '']]
            if missing:
                st.error("Please answer every question before submitting.")
            else:
                # Create success message
                manager_row = manager_employees.iloc[0]
                response_entry = create_response_entry(manager_row, selected_employee, answers, manager_comment)

                # Submit the response
                append_response(response_entry)

                # Send email
                stage_email = 'employee'
                sent, recipient, preview = send_stage_email(response_entry, stage_email)

                # Success feedback
                st.success("**Scorecard Submitted Successfully!**")
                st.info(f"Verification email sent to employee: **{recipient}**")

                # Check if there are more employees to review
                reviewed_employee_ids = set()
                if not responses_df.empty:
                    manager_submissions = responses_df[responses_df['manager_email'].astype(str).str.lower() == st.session_state.manager_email]
                    reviewed_employee_ids = set(manager_submissions['employee_id'].astype(str).unique())

                remaining_employees = []
                for _, emp in manager_employees.iterrows():
                    if str(emp['ID']) not in reviewed_employee_ids:
                        remaining_employees.append(f"{emp['name']} ({emp['ID']})")

                if remaining_employees:
                    st.warning(f"**{len(remaining_employees)} employees still need review:**")
                    for emp in remaining_employees[:3]:  # Show first 3
                        st.write(f"• {emp}")
                    if len(remaining_employees) > 3:
                        st.write(f"• ... and {len(remaining_employees) - 3} more")
                else:
                    st.success("🎉 **All employees under your supervision have been reviewed!**")

                if not sent:
                    st.warning("**Email not configured** - Copy approval links below and send manually:")
                    approve_link, reject_link = get_stage_links(response_entry)
                    st.code(f"Approve: {approve_link}")
                    st.code(f"Reject: {reject_link}")

                # Auto-refresh after 3 seconds to show updated status
                time.sleep(3)
                st.rerun()

with tab_status:
    try:
        st.markdown("### Your scorecard status dashboard")
        
        # Show debug information
        if 'load_responses_debug' in globals() and load_responses_debug:
            with st.expander("🔧 Debug Information", expanded=True):
                st.markdown("**Data Loading Debug:**")
                for msg in load_responses_debug:
                    st.code(msg)
                st.markdown("---")
        
        # Always load fresh data for the status page
        current_responses_df = load_responses()
        st.write(f"Debug: responses_df shape = {current_responses_df.shape}")
        st.write(f"Debug: manager_email = {st.session_state.manager_email}")
        
        if current_responses_df.empty:
            st.info("No scorecards submitted yet.")
        else:
            # In debug mode, show all responses; otherwise filter by manager
            if debug_mode:
                manager_responses = current_responses_df
                st.info("🔧 DEBUG MODE: Showing all responses from all managers")
            else:
                manager_responses = current_responses_df[current_responses_df['manager_email'].astype(str).str.lower() == st.session_state.manager_email]
            
            st.write(f"Debug: manager_responses shape = {manager_responses.shape}")
            
            if manager_responses.empty:
                if debug_mode:
                    st.error("No responses found in the system.")
                else:
                    st.info("No scorecards found for your manager email.")
            else:
                manager_responses = manager_responses.sort_values(['created_at'], ascending=False)
                for _, row in manager_responses.iterrows():
                    with st.expander(f"{row['employee_name']} — {row['status']}"):
                        # Create a cleaner layout with columns
                        col1, col2 = st.columns(2)

                        with col1:
                            st.markdown("Scorecard Details:")
                            st.write(f"Score: {row['questions_score']}/100")
                            st.write(f"No answers: {row['number_of_nos']}")
                            st.write(f"Created: {row['created_at']}")
                            st.write(f"Last updated: {row['updated_at']}")

                        with col2:
                            st.markdown("Approval Status:")
                            emp_status = "Yes" if row['employee_agree'] == 'Yes' else "No" if row['employee_agree'] == 'No' else "Pending"
                            st.write(f"Employee: {emp_status}")
                            if row['employee_agree_ts']:
                                st.write(f"  {row['employee_agree_ts']}")

                            mgr_status = "Yes" if row['manager_agree'] == 'Yes' else "No" if row['manager_agree'] == 'No' else "Pending"
                            st.write(f"Manager: {mgr_status}")
                            if row['manager_agree_ts']:
                                st.write(f"  {row['manager_agree_ts']}")

                            exec_status = "Yes" if row['executive_agree'] == 'Yes' else "No" if row['executive_agree'] == 'No' else "Pending"
                            st.write(f"Executive: {exec_status}")
                            if row['executive_agree_ts']:
                                st.write(f"  {row['executive_agree_ts']}")

                        # Status message with better formatting
                        st.markdown("---")
                        if row['status'] == 'Pending Employee':
                            st.info("Next Step: Waiting for employee verification via email.")
                        elif row['status'] == 'Pending Manager':
                            st.warning("Action Required: Your approval is needed.")
                        elif row['status'] == 'Pending Executive':
                            st.info("Next Step: Waiting for executive approval.")
                        elif row['status'] == 'Approved':
                            st.success("Complete: This scorecard is fully approved!")
                        else:
                            st.error("Rejected: This scorecard was rejected and requires review.")

                        # Resend email button
                        if st.button(f"Resend approval email", key=f"resend_{row['response_id']}", help="Send the current approval email again"):
                            stage = 'employee' if row['status'] == 'Pending Employee' else 'manager' if row['status'] == 'Pending Manager' else 'executive' if row['status'] == 'Pending Executive' else 'rejected'
                            sent, recipient, preview = send_stage_email(row, stage)
                            if sent:
                                st.success(f"Email resent to {recipient}.")
                            else:
                                st.warning("SMTP not configured. Use the preview links below.")
                                approve_link, reject_link = get_stage_links(row)
                                st.markdown(f"Approve: {approve_link}")
                                st.markdown(f"Reject: {reject_link}")
    except Exception as e:
        st.error(f"Error loading scorecard status: {e}")
        import traceback
        st.code(traceback.format_exc())
