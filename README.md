# AppsToSheet

Python script to connect Gmail to Sheets and parse job application confirmations.

## Steps to run and use
1. Setup through Google Cloud Platform.
2. Connect Gmail API and Google Sheets API through GCP project.
3. Configure OAuth and grab JSON file after configuration, place in project folder on desktop.
4. Create a Google Sheet.
5. Grab Google Sheet key from your Google Sheets URL:  
   `https://docs.google.com/spreadsheets/d/[KEY]/edit?gid=0#gid=0`
6. Place key in a `.env` file under variable `JOB_APPS_SHEET_ID`.
7. Run Python file.
8. Sign in with the account with Gmail and Sheets to parse.
9. Give time to allow update on your sheet.
10. Voila
