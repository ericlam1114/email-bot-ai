#!/usr/bin/env python3
"""
Email Bot AI - Setup Guide
This script helps verify that all the required files and credentials are in place.
"""

import os
import sys
import re
from dotenv import load_dotenv

def check_environment_variables():
    """Check if all required environment variables are set"""
    # Try to load environment variables from .env.local
    load_dotenv('.env.local')
    
    required_vars = [
        # Google Sheets API
        'GOOGLE_SHEET_ID',
        'GOOGLE_SHEET_RANGE',
        
        # Microsoft Graph API
        'MS_CLIENT_ID',
        'MS_TENANT_ID',
        'MS_CLIENT_SECRET',
        'MS_USER_EMAIL'
    ]
    
    optional_vars = [
        'GOOGLE_SHEETS_CREDENTIALS_FILE',
        'GOOGLE_SHEETS_TOKEN_FILE',
        'DAILY_EMAIL_LIMIT',
        'EMAIL_INTERVAL_MINUTES',
        'EMAIL_SUBJECT',
        'STATUS_COLUMN'
    ]
    
    missing_vars = []
    for var in required_vars:
        if not os.environ.get(var):
            missing_vars.append(var)
    
    # Print results
    if missing_vars:
        print("❌ Missing required environment variables:")
        for var in missing_vars:
            print(f"  - {var}")
        print("\nPlease add these to your .env.local file. Example format:")
        print("""
# Google Sheets API
GOOGLE_SHEET_ID=your_sheet_id_here
GOOGLE_SHEET_RANGE=Sheet1!A1:Z1000

# Microsoft Graph API (Outlook)
MS_CLIENT_ID=your_azure_client_id
MS_TENANT_ID=your_azure_tenant_id
MS_CLIENT_SECRET=your_azure_client_secret
MS_USER_EMAIL=your_outlook_email@example.com
        """)
    else:
        print("✅ All required environment variables are set.")
    
    # Also check optional variables
    print("\nOptional environment variables:")
    for var in optional_vars:
        value = os.environ.get(var)
        if value:
            print(f"  ✅ {var} = {value}")
        else:
            default_value = get_default_value(var)
            print(f"  ⚠️ {var} not set (will use default: {default_value})")

def get_default_value(var_name):
    """Returns the default value for an environment variable"""
    defaults = {
        'GOOGLE_SHEETS_CREDENTIALS_FILE': 'credentials.json',
        'GOOGLE_SHEETS_TOKEN_FILE': 'token.json',
        'DAILY_EMAIL_LIMIT': '10',
        'EMAIL_INTERVAL_MINUTES': '2',
        'EMAIL_SUBJECT': 'Reaching out regarding AI solutions',
        'STATUS_COLUMN': 'Status'
    }
    return defaults.get(var_name, 'None')

def check_files():
    """Check if required files exist"""
    # Check for the HTML template
    if os.path.exists('email_template.html'):
        with open('email_template.html', 'r') as f:
            content = f.read()
        
        # Extract placeholders from the template
        placeholders = re.findall(r'\{([^}]+)\}', content)
        print(f"\n✅ Email template found with placeholders: {', '.join(placeholders)}")
    else:
        print("\n❌ email_template.html not found")
    
    # Check for credentials.json
    credentials_file = os.environ.get('GOOGLE_SHEETS_CREDENTIALS_FILE', 'credentials.json')
    if os.path.exists(credentials_file):
        print(f"✅ Google credentials file ({credentials_file}) found")
    else:
        print(f"❌ Google credentials file ({credentials_file}) not found")
        print("   You need to create a Google Cloud project, enable the Sheets API, and download credentials.")
        print("   See: https://developers.google.com/sheets/api/quickstart/python")

def check_dependencies():
    """Check if the required Python packages are installed"""
    try:
        import google
        import msal
        import apscheduler
        print("\n✅ Required Python packages are installed")
    except ImportError as e:
        print(f"\n❌ Missing Python package: {str(e)}")
        print("   Please run: pip install -r requirements.txt")

def main():
    """Main function to check everything"""
    print("=" * 60)
    print("Email Bot AI - Setup Verification")
    print("=" * 60)
    
    check_environment_variables()
    check_files()
    check_dependencies()
    
    print("\nIf everything is set up correctly, you can run the bot with:")
    print("python email_bot.py")
    print("=" * 60)

if __name__ == "__main__":
    main() 