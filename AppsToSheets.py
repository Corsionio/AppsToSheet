import os
import re
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from dateutil import tz

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle

# ====== CONFIG ======
SHEET_ID = os.getenv("JOB_APPS_SHEET_ID")  # or paste the string directly
SHEET_RANGE = "Sheet1!A:G"  # adjust if your tab/range differs

# Gmail search query (tweak to your tastes/locale)
# Tip: add a Gmail label like label:jobapps to make this ultra-reliable.
GMAIL_QUERY = r'''
newer_than:90d
(subject:("application received" OR "application confirmation" OR "thank you for applying" OR "thanks for applying")
 OR "We received your application")
'''

# Scopes: read-only Gmail + read/write Sheets
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

# ====== AUTH HELPERS ======
def get_service(api_name, api_version, scopes):
    creds = None
    token_file = f"token_{api_name}.pickle"

    if os.path.exists(token_file):
        with open(token_file, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
            try:
                # Try local server on a fixed, commonly-open port
                creds = flow.run_local_server(
                    host="localhost",
                    port=8080,
                    open_browser=True,
                    authorization_prompt_message="Opening browser for Google sign-in…",
                    success_message="You may close this tab and return to the app."
                )
            except Exception as e:
                print(f"[Auth warning] Local redirect failed ({e}). Falling back to copy-paste flow.")
                # Rock-solid fallback: paste code from browser into the terminal
                creds = flow.run_console()

        with open(token_file, "wb") as f:
            pickle.dump(creds, f)

    return build(api_name, api_version, credentials=creds)


def gmail_service():
    return get_service("gmail", "v1", [SCOPES[0]])

def sheets_service():
    return get_service("sheets", "v4", [SCOPES[1]])

# ====== SHEETS IO ======
def get_existing_message_ids(svc):
    resp = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=SHEET_RANGE
    ).execute()
    rows = resp.get("values", []) or []
    if not rows:
        return set()
    # assumes header row present; "Message-ID" is column G (index 6)
    ids = set()
    for r in rows[1:]:
        if len(r) >= 7 and r[6]:
            ids.add(r[6])
    return ids

def append_rows(svc, rows):
    if not rows:
        return
    body = {"values": rows}
    svc.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range=SHEET_RANGE,
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body=body
    ).execute()

# ====== GMAIL PARSING ======
HDR = "headers"
def _get_header(headers, name):
    for h in headers:
        if h["name"].lower() == name.lower():
            return h["value"]
    return ""

def _extract_company_and_role(subject, sender):
    """
    Very rough heuristics—customize for your inbox patterns.
    """
    # Try role between quotes or before a dash
    role = ""
    m = re.search(r'["“](.+?)["”]', subject)
    if m: role = m.group(1)
    if not role:
        m = re.search(r"for (.+?)(?: at| with|$)", subject, flags=re.I)
        if m: role = m.group(1)

    # Company from From: header or subject tail
    company = ""
    # From might be "Company Careers <no-reply@company.com>"
    sender_name = re.sub(r"<.*?>", "", sender).strip()
    if sender_name and "@" not in sender_name and not sender_name.lower().startswith("no-reply"):
        company = sender_name

    if not company:
        m = re.search(r" at ([^-–|]+)", subject, flags=re.I)
        if m: company = m.group(1).strip(" .")

    return company, role

def _gmail_thread_link(thread_id):
    return f"https://mail.google.com/mail/u/0/#all/{thread_id}"

def _parse_message_dt(headers):
    # RFC2822 date → local time
    dt_hdr = _get_header(headers, "Date")
    try:
        dt = parsedate_to_datetime(dt_hdr)
        if not dt.tzinfo:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(tz.tzlocal())
    except Exception:
        return datetime.now()

def fetch_confirmation_rows(gsvc, existing_ids):
    rows = []
    try:
        search = gsvc.users().messages().list(userId="me", q=" ".join(GMAIL_QUERY.split())).execute()
        message_ids = []
        while search and search.get("messages"):
            message_ids.extend(m["id"] for m in search["messages"])
            if "nextPageToken" in search:
                search = gsvc.users().messages().list(
                    userId="me", q=" ".join(GMAIL_QUERY.split()),
                    pageToken=search["nextPageToken"]
                ).execute()
            else:
                break

        for mid in message_ids:
            msg = gsvc.users().messages().get(userId="me", id=mid, format="full").execute()
            thread_id = msg.get("threadId")
            payload = msg.get("payload", {})
            headers = payload.get(HDR, [])

            subject = _get_header(headers, "Subject")
            sender = _get_header(headers, "From")
            message_id = _get_header(headers, "Message-ID") or mid
            if message_id in existing_ids:
                continue

            dt_local = _parse_message_dt(headers)
            date_str = dt_local.strftime("%Y-%m-%d %H:%M")

            company, role = _extract_company_and_role(subject, sender)

            row = [
                date_str,
                company,
                role,
                sender,
                subject,
                _gmail_thread_link(thread_id),
                message_id
            ]
            rows.append(row)

    except HttpError as e:
        print(f"Gmail API error: {e}")
    return rows

def main():
    if not SHEET_ID:
        raise SystemExit("Set JOB_APPS_SHEET_ID environment variable to your Google Sheet ID.")

    gsvc = gmail_service()
    ssvc = sheets_service()

    existing_ids = get_existing_message_ids(ssvc)
    new_rows = fetch_confirmation_rows(gsvc, existing_ids)
    append_rows(ssvc, new_rows)
    print(f"Added {len(new_rows)} new confirmations.")

if __name__ == "__main__":
    main()
