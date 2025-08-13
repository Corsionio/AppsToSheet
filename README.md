# AppsToSheet

Python script to connect gmail to sheets and parse job application confirmations and track jobs applied to.

Steps to run and use:
  /nSetup through Google Cloud Platform.
  Connect gmail API and google sheets API through GCP project
  Configure OAuth and grab Json file after configuration, place in project folder on desktop
  Create a google sheet
  Grab google sheet key through your google sheets url: https://docs.google.com/spreadsheets/d/[KEY WILL BE HERE]/edit?gid=0#gid=0
  Place key in a .env file under variable JOB_APPS_SHEET_ID
  Run python file
  Sign in under account with gmail and sheets to parse
  Give time to allow update on your sheet
  voila
