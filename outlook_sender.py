import os
import json
import base64
import msal
import requests
import time
from dotenv import load_dotenv

# Load environment variables from .env.local
load_dotenv('.env.local')

class OutlookSender:
    def __init__(self):
        # Get credentials from environment variables
        self.client_id = os.environ.get('MS_CLIENT_ID')
        self.tenant_id = os.environ.get('MS_TENANT_ID')
        self.client_secret = os.environ.get('MS_CLIENT_SECRET')
        self.user_email = os.environ.get('MS_USER_EMAIL')
        
        # Microsoft Graph API endpoints
        self.authority = f'https://login.microsoftonline.com/{self.tenant_id}'
        self.graph_endpoint = 'https://graph.microsoft.com/v1.0'
        self.send_mail_endpoint = f'{self.graph_endpoint}/users/{self.user_email}/sendMail'
        
        # Get access token
        self.access_token = None
        self._get_access_token()
    
    def _get_access_token(self):
        """Get access token for Microsoft Graph API"""
        app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=self.authority
        )
        
        # The scope is what permissions we're requesting
        scopes = ['https://graph.microsoft.com/.default']
        
        # Get token using client credentials flow
        result = app.acquire_token_for_client(scopes=scopes)
        
        if 'access_token' in result:
            self.access_token = result['access_token']
        else:
            error_description = result.get('error_description', 'No error description provided')
            raise Exception(f"Could not get access token: {error_description}")
    
    def send_email(self, to_email, subject, content_html, retries=3):
        """
        Send an email using Microsoft Graph API (MODIFIED FOR TESTING - PRINTS INSTEAD OF SENDING)
        
        Args:
            to_email: Recipient's email address
            subject: Email subject
            content_html: HTML content of the email
            retries: Number of retries if sending fails (not used in test mode)
        
        Returns:
            bool: True (simulates successful send for testing)
        """
        print("\n--- TESTING: EMAIL WOULD BE SENT WITH THE FOLLOWING DETAILS ---")
        print(f"To: {to_email}")
        print(f"Subject: {subject}")
        print("Body (HTML):")
        print(content_html)
        print("--- END OF TEST EMAIL ---\n")
        
        # Simulate successful send for the rest of the bot's logic to proceed
        return True

        # Original sending logic commented out below:
        # if not self.access_token:
        #     self._get_access_token()
        # 
        # # Create email message
        # email_msg = {
        #     'message': {
        #         'subject': subject,
        #         'body': {
        #             'contentType': 'HTML',
        #             'content': content_html
        #         },
        #         'toRecipients': [
        #             {
        #                 'emailAddress': {
        #                     'address': to_email
        #                 }
        #             }
        #         ]
        #     },
        #     'saveToSentItems': 'true'
        # }
        # 
        # headers = {
        #     'Authorization': f'Bearer {self.access_token}',
        #     'Content-Type': 'application/json'
        # }
        # 
        # attempt = 0
        # while attempt < retries:
        #     try:
        #         response = requests.post(
        #             self.send_mail_endpoint,
        #             headers=headers,
        #             json=email_msg
        #         )
        #         
        #         # Check if successful (202 Accepted or 204 No Content is a success response)
        #         if response.status_code == 202 or response.status_code == 204:
        #             print(f"Email send request accepted for {to_email} (Status: {response.status_code})")
        #             return True
        #         else:
        #             print(f"Failed to send email: {response.status_code} - {response.text}")
        #             
        #             # If token expired, get a new one
        #             if response.status_code == 401:
        #                 print("Access token expired, refreshing...")
        #                 self._get_access_token()
        #                 headers['Authorization'] = f'Bearer {self.access_token}'
        #             
        #             attempt += 1
        #             if attempt < retries:
        #                 print(f"Retrying in 5 seconds... (Attempt {attempt+1}/{retries})")
        #                 time.sleep(5)
        #     except Exception as e:
        #         print(f"Exception while sending email: {str(e)}")
        #         attempt += 1
        #         if attempt < retries:
        #             print(f"Retrying in 5 seconds... (Attempt {attempt+1}/{retries})")
        #             time.sleep(5)
        # 
        # return False 