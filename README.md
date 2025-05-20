# Email Bot AI

A Python-based bot that sends personalized emails to contacts in a Google Sheet through Outlook, with scheduling capabilities and tracking.

## Features

- Reads recipient data from Google Sheets
- Personalizes email content using an HTML template with placeholders
- Sends emails through Outlook (Microsoft Graph API)
- Limits sending to 10 emails per day with customizable intervals
- Tracks the sending status in Google Sheets
- Automatically handles authentication and token refresh

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Google Sheets API Setup

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Enable the Google Sheets API
4. Create OAuth 2.0 credentials (Desktop application)
5. Download the credentials JSON file
6. Rename it to `credentials.json` and place it in the project root

### 3. Microsoft Graph API Setup

1. Go to the [Azure Portal](https://portal.azure.com/)
2. Register a new application in Azure Active Directory
3. Add API permissions:
   - Microsoft Graph > Application permissions > Mail.Send
4. Grant admin consent for the permissions
5. Create a client secret and note it down. **This will be your `MS_CLIENT_SECRET`**.
6. Note your Application (client) ID (**this is `MS_CLIENT_ID`**) and Directory (tenant) ID (**this is `MS_TENANT_ID`**).

### 4. Environment Configuration

This project uses environment variables to securely store sensitive information like API keys and credentials.

1.  **Create your local environment file:**
    Locate the file named `.envexample` in the project root. This file serves as a template.
    Rename this file to `.env.local`. This `.env.local` file is listed in `.gitignore` and should NOT be committed to your repository.

2.  **Populate `.env.local`:**
    Open the `.env.local` file and fill in the values for each variable.

    ```
    # Google Sheets API
    # GOOGLE_SHEETS_CREDENTIALS_FILE: This should be 'credentials.json'.
    # Ensure the 'credentials.json' file you downloaded in 'Google Sheets API Setup' (Step 2.6)
    # is placed directly in the root of your project directory.
    GOOGLE_SHEETS_CREDENTIALS_FILE=credentials.json
    # GOOGLE_SHEETS_TOKEN_FILE: This is typically 'token.json' and will be generated automatically
    # after you run the application for the first time and authenticate with Google.
    GOOGLE_SHEETS_TOKEN_FILE=token.json
    # GOOGLE_SHEET_ID: The ID of your Google Sheet. You can find this in the URL of your sheet:
    # docs.google.com/spreadsheets/d/YOUR_SHEET_ID_IS_HERE/edit
    GOOGLE_SHEET_ID=your_sheet_id_here
    # GOOGLE_SHEET_RANGE: The specific sheet and range to read data from.
    # Example: 'Sheet1!A1:Z1000' means data from cell A1 to Z1000 in 'Sheet1'.
    GOOGLE_SHEET_RANGE=Sheet1!A1:Z1000

    # Microsoft Graph API (Outlook)
    # MS_CLIENT_ID: Your Azure Application (client) ID. See 'Microsoft Graph API Setup' (Step 3.6).
    MS_CLIENT_ID=your_azure_client_id
    # MS_TENANT_ID: Your Azure Directory (tenant) ID. See 'Microsoft Graph API Setup' (Step 3.6).
    MS_TENANT_ID=your_azure_tenant_id
    # MS_CLIENT_SECRET: Your Azure application client secret. See 'Microsoft Graph API Setup' (Step 3.5).
    MS_CLIENT_SECRET=your_azure_client_secret
    # MS_USER_EMAIL: The Outlook email address you will be sending emails from.
    MS_USER_EMAIL=your_outlook_email@example.com

    # App Settings
    # DAILY_EMAIL_LIMIT: Maximum number of emails to send per day.
    DAILY_EMAIL_LIMIT=10
    # EMAIL_INTERVAL_MINUTES: Time in minutes to wait between sending each email.
    EMAIL_INTERVAL_MINUTES=2
    # EMAIL_SUBJECT: The subject line for the emails.
    EMAIL_SUBJECT=Your Email Subject Here
    # STATUS_COLUMN: The name of the column in your Google Sheet where the sending status will be written.
    # If this column doesn't exist, the script will attempt to create it.
    STATUS_COLUMN=Status
    ```

3.  **`credentials.json` (for Google Sheets API):**
    As mentioned in the "Google Sheets API Setup" (Step 2.6), after creating your OAuth 2.0 credentials in the Google Cloud Console, you will download a JSON file.
    -   Rename this file to `credentials.json`.
    -   Place this `credentials.json` file in the **root directory** of this project.
    -   This file is crucial for the application to authenticate with the Google Sheets API.
    -   Ensure this file is also listed in your `.gitignore` to prevent accidental commits. *(Self-correction: I should check if `credentials.json` is already in `.gitignore`)*

### 5. Google Sheet Preparation

1. Create a Google Sheet with the following headers:
   - `name` - The recipient's name
   - `email` - The recipient's email address
   - Other fields that match the placeholders in your email template
   - `Status` - This column will be created automatically if it doesn't exist

2. Share the Google Sheet with the email address from your Google Cloud service account

### 6. Email Template Customization

Edit the `email_template.html` file with your desired email content. Use placeholders like `{name}`, `{company_name}`, etc., that match columns in your Google Sheet.

## Usage

Run the email bot:

```bash
python email_bot.py
```

The bot will:
1. Read recipients from your Google Sheet
2. Send up to 10 emails per day with 2-minute intervals (or as configured)
3. Update the status in your Google Sheet

## Scheduling

To run the bot automatically:

### Using cron (Linux/macOS)

```bash
# Edit crontab
crontab -e

# Add this line to run daily at 9 AM
0 9 * * * cd /path/to/email-bot-ai && python email_bot.py
```

### Using Task Scheduler (Windows)

1. Open Task Scheduler
2. Create a new task
3. Set a trigger (e.g., daily at 9 AM)
4. Add an action: Start a program
   - Program/script: `python`
   - Arguments: `email_bot.py`
   - Start in: `C:\path\to\email-bot-ai`
