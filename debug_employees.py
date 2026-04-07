#!/usr/bin/env python3
import os
import sys
import pandas as pd
from google.oauth2 import service_account
import gspread

# Add current directory to path
sys.path.insert(0, os.path.dirname(__file__))

# Load secrets
try:
    import tomllib
    with open('.streamlit/secrets.toml', 'rb') as f:
        secrets = tomllib.load(f)
except ImportError:
    import tomli
    with open('.streamlit/secrets.toml', 'rb') as f:
        secrets = tomli.load(f)

GOOGLE_SHEET_ID = "1DfYJwlKy01G0tcZ11FUen4fPjSnK9vWT33a2sKfRKko"
EMPLOYEES_TAB = "Employees"

def get_spreadsheet():
    creds = service_account.Credentials.from_service_account_info(
        secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    client = gspread.authorize(creds)
    return client.open_by_key(GOOGLE_SHEET_ID)

def load_employees():
    spreadsheet = get_spreadsheet()
    worksheet = spreadsheet.worksheet(EMPLOYEES_TAB)
    records = worksheet.get_all_records()
    df = pd.DataFrame(records)
    return df

if __name__ == '__main__':
    try:
        employees_df = load_employees()
        print('Employees DataFrame shape:', employees_df.shape)
        print('Columns:', list(employees_df.columns))

        if not employees_df.empty:
            print('\nManager emails in Employees sheet:')
            manager_emails = employees_df['manager_email'].unique()
            for email in manager_emails:
                print(f'  {email}')

            print('\nAll employees:')
            for idx, row in employees_df.iterrows():
                print(f'  {row.get("employee_name", "N/A")} - Manager: {row.get("manager_email", "N/A")}')
        else:
            print('No employees found')

    except Exception as e:
        print(f'Error: {e}')
        import traceback
        traceback.print_exc()