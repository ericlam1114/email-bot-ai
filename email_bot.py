import os
import sys
import time
import datetime
import random
import re
import openai
from typing import List, Dict, Any
from dotenv import load_dotenv
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

from google_sheets import GoogleSheetsHandler
from outlook_sender import OutlookSender
from template_handler import TemplateHandler

# Load environment variables from .env.local
load_dotenv('.env.local')

class EmailBot:
    def __init__(self):
        # Initialize components
        self.sheets_handler = GoogleSheetsHandler()
        self.outlook_sender = OutlookSender()
        self.template_handler = TemplateHandler()
        
        # Email settings
        self.daily_limit = int(os.environ.get('DAILY_EMAIL_LIMIT', 10))
        self.interval_minutes = int(os.environ.get('EMAIL_INTERVAL_MINUTES', 2))
        
        # OpenAI API Key
        self.openai_api_key = os.environ.get('OPENAI_API_KEY')
        if not self.openai_api_key:
            print("Warning: OPENAI_API_KEY not found in .env.local. ChatGPT integration will not work.")
        else:
            openai.api_key = self.openai_api_key

        # Status column in Google Sheet
        self.status_column = os.environ.get('STATUS_COLUMN', 'Status')
        self.status_column_index = None
        
        # Scheduler
        self.scheduler = BackgroundScheduler()
        self.emails_sent_today = 0
        self.last_sent_time = None
    
    def initialize(self):
        """Initialize the bot and set up the scheduler"""
        print("Initializing Email Bot...")
        
        # Find or create status column
        self.status_column_index = self.sheets_handler.find_status_column_index(self.status_column)
        
        # Log the required fields from the template
        required_fields = self.template_handler.get_required_fields()
        print(f"Required fields in email template: {', '.join(required_fields)}")
        
        # Schedule the daily reset of email counter
        self.scheduler.add_job(
            self._reset_daily_counter,
            CronTrigger(hour=0, minute=0),  # Reset at midnight
            id='reset_counter'
        )
        
        # Start the scheduler
        self.scheduler.start()
        print("Scheduler started.")
    
    def _reset_daily_counter(self):
        """Reset the daily email counter"""
        self.emails_sent_today = 0
        print(f"Daily email counter reset to 0 at {datetime.datetime.now()}")
    
    def _get_chatgpt_suggestion(self, sector: str) -> str:
        if not self.openai_api_key or not sector:
            print("Skipping ChatGPT suggestion: OpenAI API key not configured or sector is missing.")
            return "" 

        prompt = f"You're writing a very short, casual follow-up sentence for an email. The recipient is in the {sector} industry. " \
                 f"After the main offer ('...we help identify and implement tailor-made solutions.'), add one brief, practical idea for a small AI automation they might find helpful. " \
                 f"Use a friendly, approachable tone. " \
                 f"Focus on a common, slightly tedious administrative or back-office task specific to the {sector} field that isn't a core strategic function and could be easily automated with a small AI tool to save a bit of time or reduce minor daily frustrations. " \
                 f"The idea should not imply replacing staff or major operational overhauls. Keep it super short and simple. " \
                 f"Do not use quotation marks. " \
                 f"For example, for a 'Law Firm' sector, a good suggestion might be: 'Just thinking, AI could probably help with organizing discovery documents or even drafting routine client updates.' " \
                 f"Now, generate a similar type of casual, practical, single sentence for the {sector} industry."
        
        try:
            print(f"Requesting ChatGPT suggestion for sector: {sector}...")
            response = openai.chat.completions.create(
                model="gpt-4.1", 
                messages=[
                    {"role": "system", "content": "You are an assistant that generates concise and relevant AI automation ideas for email outreach."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=70, 
                temperature=0.75 
            )
            suggestion = response.choices[0].message.content.strip()
            # Ensure it's a single sentence and doesn't have unwanted prefixes/suffixes if any.
            suggestion = suggestion.split('\n')[0] # Take first line if multiple
            if suggestion.startswith('"') and suggestion.endswith('"'):
                suggestion = suggestion[1:-1]

            print(f"ChatGPT suggestion for {sector}: {suggestion}")
            return suggestion
        except Exception as e:
            print(f"Error getting ChatGPT suggestion for {sector}: {str(e)}")
            return ""
    
    def _can_send_email(self):
        """Check if we can send an email (limits and timing)"""
        # Check daily limit
        if self.emails_sent_today >= self.daily_limit:
            return False
        
        # Check interval
        if self.last_sent_time:
            elapsed = (datetime.datetime.now() - self.last_sent_time).total_seconds()
            if elapsed < (self.interval_minutes * 60):
                return False
        
        return True
    
    def send_emails(self):
        """Main function to send emails"""
        while self.emails_sent_today < self.daily_limit:
            if not self._can_send_email():
                # If we can't send now, wait until next interval
                wait_time = self.interval_minutes * 60
                if self.last_sent_time:
                    elapsed = (datetime.datetime.now() - self.last_sent_time).total_seconds()
                    wait_time = max(0, (self.interval_minutes * 60) - elapsed)
                
                print(f"Waiting {wait_time:.1f} seconds until next email...")
                time.sleep(wait_time)
                continue
            
            # Get recipients who haven't been emailed yet
            recipients = self.sheets_handler.get_recipients(
                status_column_index=self.status_column_index,
                status_filter="Not Sent"
            )
            
            if not recipients:
                print("No more recipients to email. Exiting.")
                break
            
            # Take the first available recipient
            recipient = recipients[0]
            
            # Add some randomness to avoid exact same timing
            jitter = random.randint(1, 30)  # Random 1-30 second jitter
            time.sleep(jitter)
            
            # Send the email
            success = self._send_email_to_recipient(recipient)
            
            if success:
                self.emails_sent_today += 1
                self.last_sent_time = datetime.datetime.now()
                print(f"Emails sent today: {self.emails_sent_today}/{self.daily_limit}")
                
                # Update status in spreadsheet
                self.sheets_handler.update_status(
                    row_index=recipient['_row_index'],
                    status_column=self.status_column,
                    status="Sent"
                )
            else:
                # Mark as failed in spreadsheet
                self.sheets_handler.update_status(
                    row_index=recipient['_row_index'],
                    status_column=self.status_column,
                    status="Failed"
                )
            
            # If we've reached the daily limit, stop
            if self.emails_sent_today >= self.daily_limit:
                print(f"Daily limit of {self.daily_limit} emails reached.")
                break
    
    def _send_email_to_recipient(self, recipient: Dict[str, Any]) -> bool:
        """Send an email to a specific recipient"""
        try:
            # Check if recipient has email
            if 'email' not in recipient or not recipient['email']:
                print(f"No email address for recipient in row {recipient['_row_index']}")
                return False
            
            # Ensure 'sector' key exists, even if empty, for template filling
            recipient_sector = recipient.get('Sector', '')
            if isinstance(recipient_sector, list): # Handle potential list type from sheets
                recipient_sector = recipient_sector[0] if recipient_sector else ''

            recipient['sector'] = recipient_sector if recipient_sector else "your industry" # Provide fallback for subject

            # --- DEBUG PRINT --- 
            print(f"DEBUG: For recipient {recipient.get('email')}, Sector: '{recipient_sector}', OpenAI Key Loaded: {bool(self.openai_api_key)}")
            # --- END DEBUG PRINT ---

            chatgpt_suggestion = ""
            if recipient_sector and self.openai_api_key:
                chatgpt_suggestion = self._get_chatgpt_suggestion(recipient_sector)
            
            recipient['sector_specific_ai_idea'] = chatgpt_suggestion
            
            full_filled_html = self.template_handler.fill_template(recipient)
            
            subject_match = re.search(r'<title>(.*?)</title>', full_filled_html, re.IGNORECASE | re.DOTALL)
            # Fallback subject if title tag is missing or empty, or if sector was empty
            default_subject = f"AI Automation Idea for {recipient_sector if recipient_sector else 'Your Business'}"
            current_subject = default_subject
            if subject_match:
                extracted_subject = subject_match.group(1).strip()
                if extracted_subject and "{sector}" not in extracted_subject: # Check if placeholder was filled
                    current_subject = extracted_subject
                elif not recipient_sector: # if sector is empty, title might be "automation in {} idea"
                     current_subject = "AI Automation Idea"


            email_content_match = re.search(r'<body>(.*?)</body>', full_filled_html, re.IGNORECASE | re.DOTALL)
            current_email_body = full_filled_html 
            if email_content_match:
                current_email_body = email_content_match.group(1).strip()
            else: # If no body tag, maybe it's a fragment, use as is but log.
                print(f"Warning: <body> tag not found in template for recipient {recipient.get('email')}. Sending full template content.")


            success = self.outlook_sender.send_email(
                to_email=recipient['email'],
                subject=current_subject,
                content_html=current_email_body
            )
            
            return success
        except Exception as e:
            print(f"Error sending email to {recipient.get('email', 'unknown')}: {str(e)}")
            return False

    def run(self):
        """Run the email bot"""
        try:
            # Initialize the bot
            self.initialize()
            
            # Main email sending loop
            print("Starting to send emails...")
            self.send_emails()
            
            print("Email sending process completed.")
        except KeyboardInterrupt:
            print("Bot interrupted by user.")
        except Exception as e:
            print(f"Error running email bot: {str(e)}")
        finally:
            # Shutdown the scheduler
            if self.scheduler.running:
                self.scheduler.shutdown()
                print("Scheduler shut down.")


if __name__ == "__main__":
    # Create and run the email bot
    bot = EmailBot()
    bot.run() 