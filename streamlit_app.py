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

# Minimal CSS to avoid breaking Streamlit's UI rendering
custom_css = """
<style>
    .stApp {
        background:
            radial-gradient(1200px 400px at 15% -10%, color-mix(in srgb, var(--primary-color) 10%, transparent), transparent 60%),
            radial-gradient(1000px 350px at 90% 0%, color-mix(in srgb, var(--primary-color) 8%, transparent), transparent 55%),
            var(--background-color);
    }

    [data-baseweb="tab-list"] {
        gap: 0.5rem;
    }

    [data-baseweb="tab"] {
        border-radius: 10px;
        padding: 0.5rem 0.9rem;
        background: color-mix(in srgb, var(--secondary-background-color) 88%, var(--background-color));
        border: 1px solid color-mix(in srgb, var(--text-color) 12%, transparent);
    }

    [data-baseweb="tab"][aria-selected="true"] {
        border-color: color-mix(in srgb, var(--primary-color) 45%, transparent);
    }

    [data-testid="stMetricValue"] {
        color: var(--primary-color);
    }

    .stButton>button {
        background-color: var(--primary-color);
        color: var(--background-color);
        border: none;
        border-radius: 8px;
        font-weight: 500;
    }
    
    .stButton>button:hover {
        filter: brightness(0.92);
    }
    
    code {
        background-color: color-mix(in srgb, var(--secondary-background-color) 92%, var(--background-color));
        color: var(--text-color);
        border-radius: 4px;
        padding: 2px 4px;
    }
    
    /* Expander styling only */
    .streamlit-expanderHeader {
        font-weight: bold;
        color: var(--text-color);
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
EMPLOYEE_QUESTIONS_TAB = "Employee_Questions"
EMPLOYEE_RESPONSES_TAB = "Employee_Responses"

MANAGER_RESPONSE_COLUMNS = [
    "response_id", "created_at", "updated_at", "manager_email", "manager_name",
    "employee_id", "employee_name", "employee_email", "branch", "dept",
    "job_title", "executive_email", "questions_score", "number_of_nos",
    "responses", "comments", "employee_agree", "manager_agree", "executive_agree",
    "employee_agree_ts", "manager_agree_ts", "executive_agree_ts",
    "status", "employee_token", "manager_token", "executive_token"
]

EMPLOYEE_RESPONSE_COLUMNS = [
    "response_id", "created_at", "updated_at", "employee_id", "employee_name",
    "employee_email", "branch", "dept", "job_title", "responses", "status"
]

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


def ensure_dataframe_columns(df, expected_columns):
    if df.empty:
        return pd.DataFrame(columns=expected_columns)

    for col in expected_columns:
        if col not in df.columns:
            df[col] = ""

    return df


def column_letter(index):
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def ensure_sheet_headers(worksheet, headers):
    existing_headers = worksheet.row_values(1)
    if not existing_headers:
        worksheet.append_row(headers)
        return worksheet

    if existing_headers != headers:
        worksheet.update(f"A1:{column_letter(len(headers))}1", [headers])

    return worksheet

def load_responses():
    """Load responses from the Responses sheet."""
    try:
        spreadsheet = get_spreadsheet()
        worksheet = ensure_responses_sheet(spreadsheet)
        records = worksheet.get_all_records()
        df = pd.DataFrame(records)
        return ensure_dataframe_columns(df, MANAGER_RESPONSE_COLUMNS)
        
    except Exception as e:
        st.error("Could not load responses from the sheet.")
        st.error(str(e))
        return pd.DataFrame()

def ensure_responses_sheet(spreadsheet):
    try:
        worksheet = spreadsheet.worksheet(RESPONSES_TAB)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(RESPONSES_TAB, rows=1000, cols=len(MANAGER_RESPONSE_COLUMNS))
        worksheet.append_row(MANAGER_RESPONSE_COLUMNS)
        return worksheet

    return ensure_sheet_headers(worksheet, MANAGER_RESPONSE_COLUMNS)


def ensure_employee_responses_sheet(spreadsheet):
    try:
        worksheet = spreadsheet.worksheet(EMPLOYEE_RESPONSES_TAB)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(EMPLOYEE_RESPONSES_TAB, rows=1000, cols=len(EMPLOYEE_RESPONSE_COLUMNS))
        worksheet.append_row(EMPLOYEE_RESPONSE_COLUMNS)
        return worksheet

    return ensure_sheet_headers(worksheet, EMPLOYEE_RESPONSE_COLUMNS)


def load_employee_responses():
    try:
        spreadsheet = get_spreadsheet()
        worksheet = ensure_employee_responses_sheet(spreadsheet)
        records = worksheet.get_all_records()
        df = pd.DataFrame(records)
        return ensure_dataframe_columns(df, EMPLOYEE_RESPONSE_COLUMNS)
    except Exception as e:
        st.error("Could not load employee responses from the sheet.")
        st.error(str(e))
        return pd.DataFrame()

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

    if approve_link or reject_link:
        html.append("<p>Please approve or reject using the links below.</p>")
        if approve_link:
            html.append(f"<p><a href=\"{approve_link}\" style=\"background:#006B6B;color:white;padding:10px 14px;text-decoration:none;border-radius:4px;\">Approve</a></p>")
        if reject_link:
            html.append(f"<p><a href=\"{reject_link}\" style=\"background:#A80000;color:white;padding:10px 14px;text-decoration:none;border-radius:4px;\">Reject</a></p>")
    html.append("<p>If you have any questions, please contact your manager.</p>")

    # Keep comments at the end for all stages (employee, manager, executive).
    if comment.strip():
        html.append(f"<p><strong>Manager Comments:</strong></p><blockquote style='border-left: 4px solid #006B6B; padding-left: 10px; margin: 10px 0; font-style: italic;'>{comment.strip()}</blockquote>")

    return "".join(html)

# ============================================================================
# RESPONSE OPERATIONS
# ============================================================================

def find_response_by_id(response_id):
    spreadsheet = get_spreadsheet()
    worksheet = ensure_responses_sheet(spreadsheet)
    records = worksheet.get_all_records()
    df = ensure_dataframe_columns(pd.DataFrame(records), MANAGER_RESPONSE_COLUMNS)
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
        df = ensure_dataframe_columns(pd.DataFrame(records), MANAGER_RESPONSE_COLUMNS)
        
        if df.empty:
            return False

        match = df[df['response_id'] == response_id]
        if match.empty:
            return False

        row_index = match.index[0] + 2
        header = worksheet.row_values(1)
        row_values = worksheet.row_values(row_index)
        row_data = {header[i]: row_values[i] if i < len(row_values) else "" for i in range(len(header))}

        for key, value in updates.items():
            if key not in header:
                header.append(key)
                # Update the worksheet header row with the new column
                worksheet.update(f"{column_letter(len(header))}{1}", [[key]])
                row_data[key] = value
            row_data[key] = value

        ordered_row = [row_data.get(col, "") for col in header]
        update_range = f"A{row_index}:{column_letter(len(header))}{row_index}"
        worksheet.update(update_range, [ordered_row])

        return True
    except Exception as e:
        st.error(f"Unable to update response: {e}")
        return False


def append_response(row):
    spreadsheet = get_spreadsheet()
    worksheet = ensure_responses_sheet(spreadsheet)
    header = worksheet.row_values(1)
    ordered_row = [row.get(col, "") for col in header]
    worksheet.append_row(ordered_row)


def find_employee_response_by_email(employee_email):
    spreadsheet = get_spreadsheet()
    worksheet = ensure_employee_responses_sheet(spreadsheet)
    records = worksheet.get_all_records()
    df = ensure_dataframe_columns(pd.DataFrame(records), EMPLOYEE_RESPONSE_COLUMNS)
    if df.empty:
        return None, None, None

    matches = df[df['employee_email'].astype(str).str.strip().str.lower() == employee_email.strip().lower()]
    if matches.empty:
        return None, None, df

    row = matches.sort_values('created_at', ascending=False).iloc[0]
    row_index = matches.sort_values('created_at', ascending=False).index[0] + 2
    return row.to_dict(), row_index, df


def get_latest_employee_response_for_email(employee_responses_df, employee_email):
    if employee_responses_df.empty:
        return None

    matches = employee_responses_df[
        employee_responses_df['employee_email'].astype(str).str.strip().str.lower() == employee_email.strip().lower()
    ].copy()

    if matches.empty:
        return None

    if 'updated_at' in matches.columns:
        matches['sort_timestamp'] = matches['updated_at'].astype(str)
    elif 'created_at' in matches.columns:
        matches['sort_timestamp'] = matches['created_at'].astype(str)
    else:
        matches['sort_timestamp'] = ""

    latest_row = matches.sort_values('sort_timestamp', ascending=False).iloc[0]
    return latest_row.to_dict()


def append_employee_response(row):
    spreadsheet = get_spreadsheet()
    worksheet = ensure_employee_responses_sheet(spreadsheet)
    header = worksheet.row_values(1)
    ordered_row = [row.get(col, "") for col in header]
    worksheet.append_row(ordered_row)


def delete_employee_response(response_id):
    try:
        spreadsheet = get_spreadsheet()
        worksheet = ensure_employee_responses_sheet(spreadsheet)
        records = worksheet.get_all_records()
        df = ensure_dataframe_columns(pd.DataFrame(records), EMPLOYEE_RESPONSE_COLUMNS)
        if df.empty or 'response_id' not in df.columns:
            return False

        matches = df[df['response_id'] == response_id]
        if matches.empty:
            return False

        row_index = matches.index[0] + 2
        worksheet.delete_rows(row_index)
        return True
    except Exception as e:
        st.error(f"Unable to delete employee response: {e}")
        return False


def delete_response(response_id):
    """Delete a response row by response_id so manager can resubmit a new review."""
    try:
        spreadsheet = get_spreadsheet()
        worksheet = ensure_responses_sheet(spreadsheet)
        records = worksheet.get_all_records()
        df = pd.DataFrame(records)
        if df.empty or 'response_id' not in df.columns:
            return False

        matches = df[df['response_id'] == response_id]
        if matches.empty:
            return False

        row_index = matches.index[0] + 2  # header row offset
        worksheet.delete_rows(row_index)
        return True
    except Exception as e:
        st.error(f"Unable to delete response: {e}")
        return False


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


def create_employee_response_entry(employee, answers):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return {
        "response_id": str(uuid.uuid4()),
        "created_at": now,
        "updated_at": now,
        "employee_id": str(employee['ID']),
        "employee_name": employee['name'],
        "employee_email": employee['email'],
        "branch": employee.get('branch', ""),
        "dept": employee.get('dept', ""),
        "job_title": employee.get('job_title', ""),
        "responses": json.dumps(answers),
        "status": "Submitted"
    }


def load_manager_questions():
    df = load_sheet(QUESTIONS_TAB)
    if df.empty:
        st.warning("Questions sheet is empty or missing. Please check the Google Sheet.")
    return df


def load_employee_questions():
    df = load_sheet(EMPLOYEE_QUESTIONS_TAB)
    if df.empty:
        st.warning("Employee_Questions sheet is empty or missing. Please check the Google Sheet.")
    return df


def parse_response_blob(response_blob):
    try:
        return json.loads(response_blob) if isinstance(response_blob, str) else response_blob
    except Exception:
        return {}


def normalize_employee_question_type(question_type):
    normalized = str(question_type).strip().lower().replace("-", " ")
    return "three_line" if normalized == "three line" else "multi_line"


def employee_answer_complete(answer):
    if isinstance(answer, list):
        return any(str(line).strip() for line in answer)
    return bool(str(answer).strip())


def ensure_employee_answer_shape(question_type, answer):
    if normalize_employee_question_type(question_type) == "three_line":
        if isinstance(answer, list):
            values = [str(item) for item in answer[:3]]
        else:
            values = []
        values.extend([""] * (3 - len(values)))
        return values[:3]
    return str(answer or "")


def prepare_employee_questions(question_rows):
    prepared_questions = question_rows.fillna("").copy().reset_index(drop=True)

    if 'question_section' not in prepared_questions.columns:
        prepared_questions['question_section'] = ""

    if 'ID' not in prepared_questions.columns:
        prepared_questions['ID'] = ""

    id_counts = prepared_questions['ID'].astype(str).str.strip().value_counts()
    id_occurrences = {}
    response_keys = []

    for row_index, raw_id in enumerate(prepared_questions['ID'].astype(str).str.strip()):
        base_id = raw_id or f"row_{row_index + 1}"
        id_occurrences[base_id] = id_occurrences.get(base_id, 0) + 1

        if raw_id and id_counts.get(raw_id, 0) == 1:
            response_key = raw_id
        else:
            response_key = f"{base_id}__{id_occurrences[base_id]}"

        response_keys.append(response_key)

    prepared_questions['_response_key'] = response_keys
    return prepared_questions


def render_employee_question_inputs(question_rows, key_prefix):
    answers = {}
    prepared_questions = prepare_employee_questions(question_rows)

    grouped_sections = prepared_questions.groupby('question_section', dropna=False, sort=False)

    for section_name, section_rows in grouped_sections:
        if section_name:
            st.markdown(f"#### {section_name}")

        for _, question in section_rows.iterrows():
            question_id = str(question['ID'])
            response_key = str(question['_response_key'])
            question_type = normalize_employee_question_type(question.get('type', ''))

            st.markdown(f"**{question['question']}**")
            if question_type == "three_line":
                line_values = []
                for line_number in range(1, 4):
                    line_values.append(
                        st.text_input(
                            f"Line {line_number}",
                            key=f"{key_prefix}_{response_key}_line_{line_number}"
                        )
                    )
                answers[response_key] = line_values
            else:
                answers[response_key] = st.text_area(
                    "Response",
                    key=f"{key_prefix}_{response_key}",
                    height=120,
                    label_visibility='collapsed'
                )
            st.divider()

    return answers


def display_employee_response(question_rows, answers, key_prefix, read_only=True):
    prepared_questions = prepare_employee_questions(question_rows)

    grouped_sections = prepared_questions.groupby('question_section', dropna=False, sort=False)

    for section_name, section_rows in grouped_sections:
        if section_name:
            st.markdown(f"#### {section_name}")

        for _, question in section_rows.iterrows():
            question_id = str(question['ID'])
            response_key = str(question['_response_key'])
            question_type = normalize_employee_question_type(question.get('type', ''))
            stored_value = answers.get(response_key)
            if stored_value is None:
                stored_value = answers.get(question_id, [] if question_type == "three_line" else "")

            value = ensure_employee_answer_shape(question.get('type', ''), stored_value)

            st.markdown(f"**{question['question']}**")
            if question_type == "three_line":
                for line_number, line_value in enumerate(value, start=1):
                    st.text_input(
                        f"Line {line_number}",
                        value=line_value,
                        disabled=read_only,
                        key=f"{key_prefix}_{response_key}_line_{line_number}"
                    )
            else:
                st.text_area(
                    "Response",
                    value=value,
                    height=120,
                    disabled=read_only,
                    label_visibility='collapsed',
                    key=f"{key_prefix}_{response_key}"
                )
            st.divider()


def reset_login_state():
    st.session_state.logged_in = False
    st.session_state.user_role = ''
    st.session_state.manager_email = ''
    st.session_state.manager_name = ''
    st.session_state.employee_email = ''
    st.session_state.employee_name = ''

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
    question_rows = load_manager_questions()
    questions_for_email = question_rows.copy()
    try:
        answers = parse_response_blob(response['responses'])
    except Exception:
        answers = {}

    approve_link, reject_link = get_stage_links(response)
    stage_label = {
        'employee': 'Employee Verification',
        'manager': 'Manager Approval',
        'executive': 'Executive Approval'
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


def send_rejection_notice_to_manager(response, rejected_by_role, rejected_by_name, rejection_comment):
    """Notify manager that the review was rejected and removed for resubmission."""
    recipient = response.get('manager_email', '')
    if not recipient or '@' not in recipient:
        return False, recipient

    app_url = get_app_url()
    subject = "Leader Level Balanced Score Card - Rejected and Reset"
    body = [f"<h2>{subject}</h2>"]
    body.append("<p>The balanced score card below was rejected and has been removed.</p>")
    body.append(f"<p><strong>Rejected By:</strong> {rejected_by_name} ({rejected_by_role})</p>")
    body.append(f"<p><strong>Employee:</strong> {response.get('employee_name', '')} ({response.get('employee_id', '')})</p>")
    body.append(f"<p><strong>Manager:</strong> {response.get('manager_name', '')} ({response.get('manager_email', '')})</p>")

    if rejection_comment.strip():
        body.append("<p><strong>Rejection Comments:</strong></p>")
        body.append(
            f"<blockquote style='border-left:4px solid #A80000;padding-left:10px;margin:10px 0;font-style:italic;'>{rejection_comment.strip()}</blockquote>"
        )

    if app_url:
        body.append(f"<p>Please return to the app and submit a new review: <a href=\"{app_url}\">{app_url}</a></p>")
    else:
        body.append("<p>Please return to the app and submit a new review.</p>")

    success = send_email(subject, "".join(body), recipient)
    return success, recipient


def send_self_evaluation_reminder_email(employee, manager_name):
    recipient = str(employee.get('email', '')).strip()
    if not recipient or '@' not in recipient:
        return False, recipient, ""

    app_url = get_app_url()
    subject = "Action Required: Complete your self-evaluation"
    body = [
        "<h2>Self-evaluation required</h2>",
        f"<p>Hello {employee.get('name', 'Employee')},</p>",
        f"<p>{manager_name} is ready to complete your balanced score card review, but your self-evaluation has not been submitted yet.</p>",
        "<p>Please log in to the Leader Level Balanced Score Card app using your employee email and submit your self-evaluation.</p>"
    ]

    if app_url:
        body.append(f"<p><a href=\"{app_url}\" style=\"background:#006B6B;color:white;padding:10px 14px;text-decoration:none;border-radius:4px;\">Open the App</a></p>")
        body.append(f"<p>Direct link: <a href=\"{app_url}\">{app_url}</a></p>")
    else:
        body.append("<p>Please contact your manager for the app link.</p>")

    body.append("<p>Thank you.</p>")

    success = send_email(subject, "".join(body), recipient)
    return success, recipient, app_url


def process_action(action, response_id, token):
    try:
        response, _, _ = find_response_by_id(response_id)
        if not response:
            st.error("Approval record not found. Please check that the link is correct and the record exists.")
            return

        rejector_map = {
            'employee_reject': ('Employee', response.get('employee_name', 'Employee')),
            'manager_reject': ('Manager', response.get('manager_name', response.get('manager_email', 'Manager'))),
            'executive_reject': ('Executive', response.get('executive_email', 'Executive'))
        }

        rejection_comment = ""
        if action in rejector_map:
            role_label, rejector_name = rejector_map[action]
            st.warning(f"You are rejecting this score card as {role_label}.")
            rejection_comment = st.text_area(
                "Please provide a rejection comment (required):",
                key=f"reject_comment_{response_id}_{action}",
                height=120
            )
            if not st.button("Submit Rejection", type='primary', key=f"submit_reject_{response_id}_{action}"):
                return
            if not rejection_comment.strip():
                st.error("A rejection comment is required.")
                return

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

        if not valid:
            st.error("This approval link is invalid or the action is not allowed.")
            return

        if action in rejector_map:
            role_label, rejector_name = rejector_map[action]
            deleted = delete_response(response_id)
            if not deleted:
                st.error("Unable to reset this review after rejection. Please contact support.")
                return

            sent, recipient = send_rejection_notice_to_manager(response, role_label, rejector_name, rejection_comment)
            st.success("Thank you. The scorecard was rejected and removed so a new review can be submitted.")
            if sent:
                st.info(f"Manager notification sent to {recipient}.")
            else:
                st.warning("Could not send manager notification email. Please notify the manager manually.")
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
    st.session_state.user_role = 'manager'
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

if 'logged_in' not in st.session_state:
    reset_login_state()
    st.session_state.data_loaded = False

if 'user_role' not in st.session_state:
    st.session_state.user_role = ''

if 'employee_email' not in st.session_state:
    st.session_state.employee_email = ''

if 'employee_name' not in st.session_state:
    st.session_state.employee_name = ''

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
    manager_col, employee_col = st.columns(2)

    with manager_col:
        st.subheader("Manager Login")
        manager_login_email = st.text_input(
            "Enter your manager email:",
            placeholder="manager@example.com",
            key='manager_login_email'
        ).strip().lower()
        if st.button("Login as Manager", type='primary'):
            manager_emails = st.session_state.employees_df['manager_email'].astype(str).str.lower().values
            if manager_login_email and manager_login_email in manager_emails:
                manager_row = st.session_state.employees_df[
                    st.session_state.employees_df['manager_email'].astype(str).str.lower() == manager_login_email
                ].iloc[0]
                st.session_state.logged_in = True
                st.session_state.user_role = 'manager'
                st.session_state.manager_email = manager_login_email
                st.session_state.manager_name = manager_row.get('manager_name', manager_login_email)
                st.session_state.employee_email = ''
                st.session_state.employee_name = ''
                st.rerun()
            else:
                st.error("Manager email not found. Please enter an email listed in the Employees sheet.")

    with employee_col:
        st.subheader("Employee Login")
        employee_login_email = st.text_input(
            "Enter your employee email:",
            placeholder="employee@example.com",
            key='employee_login_email'
        ).strip().lower()
        if st.button("Login as Employee", type='primary'):
            employee_emails = st.session_state.employees_df['email'].astype(str).str.lower().values
            if employee_login_email and employee_login_email in employee_emails:
                employee_row = st.session_state.employees_df[
                    st.session_state.employees_df['email'].astype(str).str.lower() == employee_login_email
                ].iloc[0]
                st.session_state.logged_in = True
                st.session_state.user_role = 'employee'
                st.session_state.employee_email = employee_login_email
                st.session_state.employee_name = employee_row.get('name', employee_login_email)
                st.session_state.manager_email = ''
                st.session_state.manager_name = ''
                st.rerun()
            else:
                st.error("Employee email not found. Please enter an email listed in the Employees sheet.")

    st.stop()

if st.session_state.user_role == 'manager':
    sidebar_identity = f"{st.session_state.manager_name} ({st.session_state.manager_email})"
else:
    sidebar_identity = f"{st.session_state.employee_name} ({st.session_state.employee_email})"

st.sidebar.markdown(f"**Signed in as:** {sidebar_identity}")
st.sidebar.caption(f"Role: {st.session_state.user_role.title()}")
if st.sidebar.button("Logout"):
    reset_login_state()
    st.rerun()

if st.session_state.user_role == 'manager':
    if debug_mode:
        manager_employees = st.session_state.employees_df
    else:
        manager_employees = st.session_state.employees_df[
            st.session_state.employees_df['manager_email'].astype(str).str.lower() == st.session_state.manager_email
        ]

    responses_df = load_responses()
    responses_df['employee_id'] = responses_df['employee_id'].astype(str)

    manager_has_responses = not responses_df[
        responses_df['manager_email'].astype(str).str.lower() == st.session_state.manager_email
    ].empty if not debug_mode else True

    if not debug_mode and manager_employees.empty and not manager_has_responses:
        st.warning("Manager email not found in the Employees sheet.")
        st.info("You can still view scorecard statuses for any scorecards you've submitted.")

    employee_self_responses_df = load_employee_responses()

    st.subheader("Manager Dashboard")

    tab_new, tab_status, tab_self_eval = st.tabs(["Submit Scorecard", "Scorecard Status", "Employee Self-Evaluations"])

    with tab_new:
        st.markdown("### Submit a new balanced score card")
        can_submit = True

        if debug_mode and manager_employees.empty:
            st.info("No employees found matching current filters.")
            can_submit = False
        elif not debug_mode and manager_employees.empty:
            st.info("You don't have any employees assigned as a manager in the system.")
            st.info("You can still view the status of any scorecards you've submitted using the Scorecard Status tab.")
            can_submit = False

        if can_submit:
            reviewed_employee_ids = set()
            if not responses_df.empty:
                if debug_mode:
                    reviewed_employee_ids = set(responses_df['employee_id'].astype(str).unique())
                else:
                    manager_submissions = responses_df[
                        responses_df['manager_email'].astype(str).str.lower() == st.session_state.manager_email
                    ]
                    reviewed_employee_ids = set(manager_submissions['employee_id'].astype(str).unique())

            if debug_mode:
                available_employees = st.session_state.employees_df[
                    ~st.session_state.employees_df['ID'].astype(str).isin(reviewed_employee_ids)
                ]
            else:
                available_employees = manager_employees[
                    ~manager_employees['ID'].astype(str).isin(reviewed_employee_ids)
                ]

            if available_employees.empty:
                if debug_mode:
                    st.success("All employees in the system have been reviewed.")
                else:
                    st.success("All employees under your supervision have been reviewed.")
                can_submit = False

        if can_submit:
            selected_employee_id = st.selectbox(
                "Select employee to rate",
                available_employees['ID'].astype(str).tolist(),
                format_func=lambda eid: f"{available_employees[available_employees['ID'].astype(str) == eid].iloc[0]['name']} ({eid})"
            )
            selected_employee = available_employees[available_employees['ID'].astype(str) == selected_employee_id].iloc[0]
            selected_employee_email = str(selected_employee.get('email', '')).strip().lower()
            selected_employee_self_eval = get_latest_employee_response_for_email(
                employee_self_responses_df,
                selected_employee_email
            )

            st.write(f"Employee: {selected_employee['name']} | Branch: {selected_employee.get('branch', '')} | Dept: {selected_employee.get('dept', '')}")
            st.write(f"Title: {selected_employee.get('job_title', '')} | Executive: {selected_employee.get('executive_email', '')}")
            st.divider()

            st.markdown("### Employee Self-Evaluation")
            employee_questions_df = load_employee_questions()

            if selected_employee_self_eval:
                st.success("Self-evaluation submitted. You can proceed with manager review.")
                if employee_questions_df.empty:
                    st.info("Employee questions are not configured, so the self-evaluation detail cannot be rendered.")
                else:
                    display_employee_response(
                        employee_questions_df,
                        parse_response_blob(selected_employee_self_eval.get('responses', {})),
                        key_prefix=f"manager_self_eval_view_{selected_employee_id}_{selected_employee_self_eval.get('response_id', 'latest')}"
                    )
            else:
                st.warning("This employee has not submitted a self-evaluation yet. Manager review is disabled until they submit one.")
                reminder_button_key = f"send_self_eval_reminder_{selected_employee_id}"
                if st.button("Send Self-Evaluation Reminder Email", key=reminder_button_key):
                    sent, recipient, app_url = send_self_evaluation_reminder_email(selected_employee, st.session_state.manager_name)
                    if sent:
                        st.success(f"Reminder email sent to {recipient}.")
                    else:
                        st.warning("Could not send reminder email. SMTP may not be configured.")
                        if app_url:
                            st.info("Share this app link with the employee:")
                            st.code(app_url)
                        else:
                            st.info("Set app.url in Streamlit secrets so a direct login link can be included.")

            if selected_employee_self_eval:
                questions_df = st.session_state.questions_df
                if questions_df.empty:
                    st.warning("The Questions sheet is empty. Please add questions to the Google Sheet.")
                else:
                    questions_df = questions_df.fillna("")
                    answers = {}
                    questions_df['question_section'] = questions_df['question_section'].astype(str).fillna('').str.strip()
                    grouped_sections = questions_df.groupby('question_section', dropna=False, sort=False)

                    for section_name, section_rows in grouped_sections:
                        if section_name:
                            st.markdown(f"#### {section_name}")

                        for _, question in section_rows.iterrows():
                            header_text = str(question.get('header', '')).strip()
                            if header_text:
                                st.caption(header_text)

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

                    answered_score_questions = [qid for qid, val in answers.items() if val in ['1', '2', '3']]
                    current_score = 0
                    if answered_score_questions:
                        score_values = [int(answers[qid]) for qid in answered_score_questions]
                        current_score = int(round(sum(score_values) / len(score_values) * 100)) if score_values else 0

                    no_answers = sum(1 for qid, val in answers.items() if val == 'No')

                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Current Score", f"{current_score}")
                    with col2:
                        answered = len([v for v in answers.values() if v not in [None, '']])
                        total_questions = len(questions_df)
                        st.metric("Questions Answered", f"{answered}/{total_questions}")
                    with col3:
                        st.metric("No Answers", no_answers)

                    st.divider()

                    st.markdown("### Additional Comments (Optional)")
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
                            manager_row = manager_employees.iloc[0]
                            response_entry = create_response_entry(manager_row, selected_employee, answers, manager_comment)
                            append_response(response_entry)

                            stage_email = 'employee'
                            sent, recipient, preview = send_stage_email(response_entry, stage_email)

                            st.success("Scorecard submitted successfully.")
                            if recipient:
                                st.info(f"Verification email sent to employee: {recipient}")

                            reviewed_employee_ids = set()
                            if not responses_df.empty:
                                manager_submissions = responses_df[
                                    responses_df['manager_email'].astype(str).str.lower() == st.session_state.manager_email
                                ]
                                reviewed_employee_ids = set(manager_submissions['employee_id'].astype(str).unique())

                            remaining_employees = []
                            for _, emp in manager_employees.iterrows():
                                if str(emp['ID']) not in reviewed_employee_ids:
                                    remaining_employees.append(f"{emp['name']} ({emp['ID']})")

                            if remaining_employees:
                                st.warning(f"{len(remaining_employees)} employees still need review:")
                                for emp in remaining_employees[:3]:
                                    st.write(f"• {emp}")
                                if len(remaining_employees) > 3:
                                    st.write(f"• ... and {len(remaining_employees) - 3} more")
                            else:
                                st.success("All employees under your supervision have been reviewed.")

                            if not sent:
                                st.warning("Email not configured. Copy approval links below and send manually:")
                                approve_link, reject_link = get_stage_links(response_entry)
                                st.code(f"Approve: {approve_link}")
                                st.code(f"Reject: {reject_link}")

                            time.sleep(3)
                            st.rerun()

    with tab_status:
        st.markdown("### Your scorecard status dashboard")
        st.caption("Track approvals and progress for scorecards you've submitted.")

        try:
            current_responses_df = load_responses()
            manager_email = st.session_state.get('manager_email', '').strip().lower()

            if current_responses_df.empty:
                st.info("No scorecards submitted yet.")
            else:
                if debug_mode:
                    manager_responses = current_responses_df.copy()
                else:
                    manager_responses = current_responses_df[
                        current_responses_df['manager_email'].astype(str).str.strip().str.lower() == manager_email
                    ].copy()

                if manager_responses.empty:
                    st.info("No scorecards found for your manager email.")
                else:
                    manager_responses = manager_responses.sort_values(['created_at'], ascending=False)

                    total_count = len(manager_responses)
                    approved_count = int((manager_responses['status'].astype(str) == 'Approved').sum())
                    pending_count = int(manager_responses['status'].astype(str).str.startswith('Pending').sum())
                    rejected_count = int(manager_responses['status'].astype(str).str.contains('Rejected').sum())

                    k1, k2, k3, k4 = st.columns(4)
                    k1.metric("Total Submitted", total_count)
                    k2.metric("Approved", approved_count)
                    k3.metric("Pending", pending_count)
                    k4.metric("Rejected", rejected_count)

                    st.markdown("#### Submission Summary")
                    summary_columns = [
                        'employee_name', 'employee_email', 'status',
                        'questions_score', 'number_of_nos', 'created_at', 'updated_at'
                    ]
                    available_summary_columns = [c for c in summary_columns if c in manager_responses.columns]
                    summary_df = manager_responses[available_summary_columns].copy()
                    summary_df = summary_df.rename(columns={
                        'employee_name': 'Employee',
                        'employee_email': 'Email',
                        'status': 'Status',
                        'questions_score': 'Score',
                        'number_of_nos': 'No Answers',
                        'created_at': 'Created',
                        'updated_at': 'Updated'
                    })
                    st.dataframe(summary_df, use_container_width=True, hide_index=True)

                    st.markdown("#### Scorecard Details")
                    for _, row in manager_responses.iterrows():
                        title = f"{row.get('employee_name', 'Unknown Employee')} • {row.get('status', '')}"
                        with st.expander(title):
                            c1, c2 = st.columns(2)
                            with c1:
                                st.write(f"Score: {row.get('questions_score', '')}")
                                st.write(f"No answers: {row.get('number_of_nos', '')}")
                                st.write(f"Created: {row.get('created_at', '')}")
                            with c2:
                                st.write(f"Updated: {row.get('updated_at', '')}")
                                st.write(f"Employee: {row.get('employee_email', '')}")
                                st.write(f"Manager: {row.get('manager_email', '')}")
        except Exception as e:
            st.error(f"Error loading scorecard status: {e}")
            import traceback
            st.code(traceback.format_exc())

    with tab_self_eval:
        st.markdown("### Employee self-evaluations")
        st.caption("View submitted self-evaluations for employees assigned to you.")

        if manager_employees.empty:
            st.info("No employees found for this manager.")
        elif employee_self_responses_df.empty:
            st.info("No employee self-evaluations have been submitted yet.")
        else:
            employee_questions_df = load_employee_questions()
            if employee_questions_df.empty:
                st.warning("Employee_Questions sheet is empty. Add questions to render self-evaluation details.")
            else:
                for _, manager_employee in manager_employees.iterrows():
                    employee_name = manager_employee.get('name', 'Unknown Employee')
                    employee_id = str(manager_employee.get('ID', ''))
                    employee_email = str(manager_employee.get('email', '')).strip().lower()

                    self_eval = get_latest_employee_response_for_email(employee_self_responses_df, employee_email)
                    title_suffix = "Submitted" if self_eval else "Not Submitted"

                    with st.expander(f"{employee_name} ({employee_id}) • {title_suffix}"):
                        st.write(f"Email: {manager_employee.get('email', '')}")
                        if not self_eval:
                            st.info("No self-evaluation submitted.")
                        else:
                            st.write(f"Last updated: {self_eval.get('updated_at', self_eval.get('created_at', ''))}")
                            display_employee_response(
                                employee_questions_df,
                                parse_response_blob(self_eval.get('responses', {})),
                                key_prefix=f"manager_tab_self_eval_{employee_id}_{self_eval.get('response_id', 'latest')}"
                            )
else:
    employee_row = st.session_state.employees_df[
        st.session_state.employees_df['email'].astype(str).str.lower() == st.session_state.employee_email
    ]

    st.subheader("Employee Dashboard")

    if employee_row.empty:
        st.error("Employee email not found in the Employees sheet.")
        st.stop()

    employee_record = employee_row.iloc[0]
    st.write(f"Employee: {employee_record['name']} | Branch: {employee_record.get('branch', '')} | Dept: {employee_record.get('dept', '')}")
    st.write(f"Title: {employee_record.get('job_title', '')}")
    st.divider()

    employee_questions_df = load_employee_questions()
    employee_responses_df = load_employee_responses()
    existing_response, _, _ = find_employee_response_by_email(st.session_state.employee_email)

    if employee_questions_df.empty:
        st.warning("The Employee_Questions sheet is empty. Please add questions to the Google Sheet.")
        st.stop()

    if existing_response:
        st.markdown("### Your submitted response")
        st.caption("You have already submitted your employee response. Delete it if you need to start over.")
        display_employee_response(
            employee_questions_df,
            parse_response_blob(existing_response.get('responses', {})),
            key_prefix=f"employee_existing_{existing_response['response_id']}"
        )

        if st.button("Delete Existing Response and Start Over", type='primary'):
            if delete_employee_response(existing_response['response_id']):
                st.success("Your existing response was deleted. You can now submit a new one.")
                st.rerun()
            else:
                st.error("Unable to delete your response. Please try again.")
    else:
        st.markdown("### Submit your employee response")
        st.caption("Each employee can submit one response for themselves. Your answers do not affect scorecard scoring.")
        employee_answers = render_employee_question_inputs(
            employee_questions_df,
            key_prefix=f"employee_form_{employee_record['ID']}"
        )

        if st.button("Submit Employee Response", type='primary'):
            missing_answers = [
                question_id for question_id, answer in employee_answers.items()
                if not employee_answer_complete(answer)
            ]

            if missing_answers:
                st.error("Please answer every employee question before submitting.")
            else:
                response_entry = create_employee_response_entry(employee_record, employee_answers)
                append_employee_response(response_entry)
                st.success("Your employee response has been submitted.")
                st.rerun()
