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
RESPONSES_TAB = 'Responses'

def get_spreadsheet():
    creds = service_account.Credentials.from_service_account_info(
        secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    client = gspread.authorize(creds)
    return client.open_by_key(GOOGLE_SHEET_ID)

def load_responses():
    spreadsheet = get_spreadsheet()
    worksheet = spreadsheet.worksheet(RESPONSES_TAB)
    records = worksheet.get_all_records()
    df = pd.DataFrame(records)
    return df

if __name__ == '__main__':
    try:
        responses_df = load_responses()
        print('Responses DataFrame shape:', responses_df.shape)
        print('Columns:', list(responses_df.columns))

        if not responses_df.empty:
            print('\nAll response_ids in sheet:')
            for idx, row in responses_df.iterrows():
                print(f'  {row["response_id"]} - Status: {row["status"]} - Manager: {row.get("manager_email", "N/A")}')

            # Show full details of the first response
            print('\nFull details of first response:')
            first_row = responses_df.iloc[0]
            for col in responses_df.columns:
                print(f'  {col}: {first_row[col]}')

            # Check for the specific response_id
            target_id = '0885fe7b-0dd0-4fab-8768-62a1bdbdbfd4'
            matches = responses_df[responses_df['response_id'] == target_id]
            if not matches.empty:
                print(f'\nFound target response_id: {target_id}')
                row = matches.iloc[0]
                print(f'Status: {row["status"]}')
                print(f'Employee Token: {row.get("employee_token", "N/A")}')
                print(f'Employee Email: {row.get("employee_email", "N/A")}')
                print(f'All tokens:')
                print(f'  Employee: {row.get("employee_token", "N/A")}')
                print(f'  Manager: {row.get("manager_token", "N/A")}')
                print(f'  Executive: {row.get("executive_token", "N/A")}')
            else:
                print(f'\nTarget response_id {target_id} NOT found')
        else:
            print('No responses found')

    except Exception as e:
        print(f'Error: {e}')
        import traceback
        traceback.print_exc()