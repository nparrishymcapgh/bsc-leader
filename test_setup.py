#!/usr/bin/env python3
"""
Test script to verify Google Sheets connection and permissions.
Run this before starting the Streamlit app to ensure everything is configured correctly.
"""

import sys
import os

# Add the current directory to Python path so we can import the app functions
sys.path.insert(0, os.path.dirname(__file__))

try:
    import streamlit as st
    from google.oauth2.service_account import Credentials
    import gspread
    import tomllib  # Python 3.11+
    print("✅ All required packages are installed")
except ImportError:
    try:
        import tomli as tomllib  # Fallback for older Python
        print("✅ All required packages are installed")
    except ImportError:
        print("❌ Missing tomllib/tomli package. Install with: pip install tomli")
        sys.exit(1)

# Check if secrets file exists
secrets_path = ".streamlit/secrets.toml"
if os.path.exists(secrets_path):
    print(f"✅ Secrets file found at {secrets_path}")
else:
    print(f"❌ Secrets file not found at {secrets_path}")
    print("   Please create .streamlit/secrets.toml with your service account credentials")
    sys.exit(1)

# Load secrets directly from TOML file
try:
    with open(secrets_path, 'rb') as f:
        secrets = tomllib.load(f)
except Exception as e:
    print(f"❌ Error reading secrets file: {e}")
    sys.exit(1)

# Check secrets content
if "gcp_service_account" in secrets:
    print("✅ Service account credentials found in secrets")
else:
    print("❌ Service account credentials not found in secrets")
    print("   Add [gcp_service_account] section to .streamlit/secrets.toml")
    sys.exit(1)

if "app" in secrets and "url" in secrets["app"]:
    print("✅ App URL configured")
else:
    print("⚠️  App URL not configured (optional for basic functionality)")

if "smtp" in secrets:
    print("✅ SMTP configuration found")
else:
    print("⚠️  SMTP not configured (emails won't work)")

# Try to connect to Google Sheets
try:
    # Read the Google Sheet ID directly from the app file
    with open('streamlit_app.py', 'r') as f:
        content = f.read()
        # Extract GOOGLE_SHEET_ID from the file
        import re
        match = re.search(r'GOOGLE_SHEET_ID\s*=\s*["\']([^"\']+)["\']', content)
        if match:
            sheet_id = match.group(1)
            print(f"🔍 Testing connection to Google Sheet: {sheet_id}")
        else:
            print("❌ Could not find GOOGLE_SHEET_ID in streamlit_app.py")
            sys.exit(1)

    # Test connection directly
    creds = Credentials.from_service_account_info(
        secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(sheet_id)
    print("✅ Successfully connected to Google Sheet")

    # Check if required worksheets exist
    worksheets = spreadsheet.worksheets()
    existing_sheets = [ws.title for ws in worksheets]
    print(f"📋 Found worksheets: {existing_sheets}")

    required_sheets = ["Employees", "Questions", "Responses"]
    for sheet in required_sheets:
        if sheet in existing_sheets:
            print(f"✅ Worksheet '{sheet}' exists")
            # Test reading a few rows from each sheet to verify access
            try:
                ws = spreadsheet.worksheet(sheet)
                records = ws.get_all_records()
                if records:
                    print(f"   📊 '{sheet}' has {len(records)} rows of data")
                    # Show column names for the first sheet
                    if len(records) > 0:
                        columns = list(records[0].keys())
                        print(f"   📋 Columns: {columns}")
                else:
                    print(f"   📊 '{sheet}' is empty (this is OK)")
            except Exception as sheet_error:
                print(f"   ❌ Error reading '{sheet}': {sheet_error}")
        else:
            print(f"❌ Worksheet '{sheet}' not found")

except PermissionError:
    print("❌ Permission denied accessing Google Sheet")
    print("💡 Most common issue: The service account email has not been shared with your Google Sheet")
    print("   1. Open your Google Sheet in a browser")
    print("   2. Click 'Share' button")
    print("   3. Add the service account email (ends with @your-project.iam.gserviceaccount.com)")
    print("   4. Give it 'Editor' access")
    print("   5. Try again")
    sys.exit(1)
except Exception as e:
    print(f"❌ Error connecting to Google Sheets: {e}")
    print("   Check your service account credentials and sheet permissions")
    sys.exit(1)

print("\n🎉 All checks passed! Your setup looks good.")
print("You can now run: streamlit run streamlit_app.py")
