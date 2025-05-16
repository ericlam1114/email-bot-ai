import os
import json
import datetime
from typing import List, Dict, Any
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from dotenv import load_dotenv

# Load environment variables from .env.local
load_dotenv('.env.local')

# Define the scopes
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

class GoogleSheetsHandler:
    def __init__(self):
        self.sheet_id = os.environ.get('GOOGLE_SHEET_ID')
        self.sheet_range = os.environ.get('GOOGLE_SHEET_RANGE', 'Sheet1!A1:Z1000')
        self.creds = None
        self.service = None
        self._authenticate()

    def _authenticate(self):
        """Authenticate with Google Sheets API"""
        creds = None
        token_file = os.environ.get('GOOGLE_SHEETS_TOKEN_FILE', 'token.json')
        credentials_file = os.environ.get('GOOGLE_SHEETS_CREDENTIALS_FILE', 'credentials.json')

        # Check if token.json exists with valid credentials
        if os.path.exists(token_file):
            creds = Credentials.from_authorized_user_info(
                json.loads(open(token_file).read()), SCOPES)

        # If there are no valid credentials, let the user log in
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    credentials_file, SCOPES)
                creds = flow.run_local_server(port=0)
            
            # Save the credentials for the next run
            with open(token_file, 'w') as token:
                token.write(creds.to_json())

        self.creds = creds
        self.service = build('sheets', 'v4', credentials=creds)

    def get_recipients(self, status_column_index=None, status_filter="Not Sent") -> List[Dict[str, str]]:
        """
        Get recipients from Google Sheets
        
        Args:
            status_column_index: Index of the column that contains the email status
            status_filter: Filter to apply to the status column (e.g., "Not Sent")
            
        Returns:
            List of dictionaries with recipient data
        """
        sheet = self.service.spreadsheets()
        result = sheet.values().get(spreadsheetId=self.sheet_id, range=self.sheet_range).execute()
        values = result.get('values', [])

        if not values:
            print('No data found in the sheet.')
            return []

        headers = values[0]  # First row contains headers
        recipients = []

        for i, row in enumerate(values[1:], 1):  # Skip header row
            # Pad row with empty strings if it's shorter than headers
            padded_row = row + [''] * (len(headers) - len(row))
            
            # Create dictionary with header as key and cell value as value
            recipient = {headers[j]: padded_row[j] for j in range(len(headers))}
            
            # Add row index for later updates
            recipient['_row_index'] = i + 1  # +1 because we're skipping header
            
            # If a status column is specified, filter by it
            if status_column_index is not None:
                try:
                    status = padded_row[status_column_index]
                    if status != status_filter:
                        continue
                except IndexError:
                    pass  # If status column doesn't exist, include all recipients
            
            recipients.append(recipient)

        return recipients

    def update_status(self, row_index: int, status_column: str, status: str) -> None:
        """
        Update the status of an email in the sheet and add today's date
        
        Args:
            row_index: Index of the row to update (1-based)
            status_column: Name of the status column
            status: New status value
        """
        # First get the headers to find the status column index
        sheet = self.service.spreadsheets()
        result = sheet.values().get(spreadsheetId=self.sheet_id, 
                                    range=f"{self.sheet_range.split('!')[0]}!1:1").execute()
        headers = result.get('values', [[]])[0]
        
        if status_column not in headers:
            raise ValueError(f"Status column '{status_column}' not found in sheet headers.")
        
        # Find the status column index
        status_column_index = headers.index(status_column)
        status_column_letter = chr(ord('A') + status_column_index)
        
        # Find the Date column index if it exists
        date_column_index = None
        date_column_letter = None
        
        if 'Date' in headers:
            date_column_index = headers.index('Date')
            date_column_letter = chr(ord('A') + date_column_index)
        
        # Update the status cell
        range_to_update = f"{self.sheet_range.split('!')[0]}!{status_column_letter}{row_index}"
        body = {
            'values': [[status]]
        }
        
        sheet.values().update(
            spreadsheetId=self.sheet_id,
            range=range_to_update,
            valueInputOption='RAW',
            body=body
        ).execute()
        
        print(f"Updated row {row_index}, column {status_column} to '{status}'")
        
        # If we found a Date column and the status is "Sent", update the date too
        if date_column_letter and status == "Sent":
            today = datetime.datetime.now().strftime("%-m/%-d")  # Format as M/D
            date_range = f"{self.sheet_range.split('!')[0]}!{date_column_letter}{row_index}"
            date_body = {
                'values': [[today]]
            }
            
            sheet.values().update(
                spreadsheetId=self.sheet_id,
                range=date_range,
                valueInputOption='RAW',
                body=date_body
            ).execute()
            
            print(f"Updated row {row_index}, column Date to '{today}'")
        
    def find_status_column_index(self, status_column: str = "Status") -> int:
        """Find the index of the status column in the sheet"""
        sheet = self.service.spreadsheets()
        result = sheet.values().get(spreadsheetId=self.sheet_id, 
                                    range=f"{self.sheet_range.split('!')[0]}!1:1").execute()
        headers = result.get('values', [[]])[0]
        
        if status_column not in headers:
            # If status column doesn't exist, create it
            headers.append(status_column)
            sheet.values().update(
                spreadsheetId=self.sheet_id,
                range=f"{self.sheet_range.split('!')[0]}!1:1",
                valueInputOption='RAW',
                body={'values': [headers]}
            ).execute()
            return len(headers) - 1
        
        return headers.index(status_column) 