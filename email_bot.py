import os
import sys
import time
import datetime
import random
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
        self.email_subject = os.environ.get('EMAIL_SUBJECT', 'Reaching out regarding AI solutions')
        
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
            
            # Fill the email template with recipient data
            email_content = self.template_handler.fill_template(recipient)
            
            # Send the email using Outlook
            success = self.outlook_sender.send_email(
                to_email=recipient['email'],
                subject=self.email_subject,
                content_html=email_content
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