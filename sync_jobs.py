"""
Job Application Tracker — Gmail → Excel Sync
Reads acknowledgment/rejection emails from Gmail and appends new entries to your Excel tracker.
"""

import os
import re
import base64
import pickle
from datetime import datetime

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Config ────────────────────────────────────────────────────────────────────

EXCEL_FILE        = "job_applications.xlsx"
SHEET_NAME        = "Applications"
TOKEN_FILE        = "token.pickle"
CREDENTIALS_FILE  = "credentials.json"
SCOPES            = ["https://www.googleapis.com/auth/gmail.readonly"]

GMAIL_QUERY = (
    'subject:("application received" OR "thank you for applying" OR '
    '"thanks for applying" OR "we received your application" OR '
    '"application confirmation" OR "your application" OR '
    '"application for" OR "thank you for your interest" OR '
    '"unfortunately" OR "not been successful" OR "not move forward" OR '
    '"position has been filled" OR "we will not" OR "regret to inform")'
)

# ── Rejection signals ─────────────────────────────────────────────────────────

REJECTION_PATTERNS = [
    r'\bregret\b',
    r'\bunfortunately\b',
    r'\bnot (?:been )?successful\b',
    r'\bnot (?:be )?moving forward\b',
    r'\bnot move forward\b',
    r'\bwill not be moving\b',
    r'\bdecided not to proceed\b',
    r'\bposition has been filled\b',
    r'\bno longer (?:being )?considered\b',
    r'\bwe have (?:decided|chosen) to\b.*\bother\b',
    r'\bdid not (?:make|pass)\b',
    r'\bnot selected\b',
    r'\bnot shortlisted\b',
    r'\bwe(?:\'ve| have) decided to move forward with (?:other|another)\b',
    r'\bthank you for your (?:time|interest).{0,80}(?:regret|unfortunately|not)\b',
]

def detect_status(subject: str, body: str) -> str:
    text = (subject + " " + body[:3000]).lower()
    for pattern in REJECTION_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return "Rejected"
    return "Pending"

# ── Gmail auth ────────────────────────────────────────────────────────────────

def get_gmail_service():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)
    return build("gmail", "v1", credentials=creds)

# ── Email fetching ────────────────────────────────────────────────────────────

def fetch_all_messages(service):
    """Fetch the latest 200 matching messages."""
    response = service.users().messages().list(
        userId="me", q=GMAIL_QUERY, maxResults=200
    ).execute()
    return response.get("messages", [])

# ── Email parsing ─────────────────────────────────────────────────────────────

def get_email_body(msg_payload):
    if msg_payload.get("mimeType") == "text/plain":
        data = msg_payload.get("body", {}).get("data", "")
        return base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore") if data else ""
    for part in msg_payload.get("parts", []):
        result = get_email_body(part)
        if result:
            return result
    return ""


def extract_company_from_sender(sender: str) -> str:
    name_match = re.match(r'^"?([^"<]+)"?\s*<', sender)
    if name_match:
        name = name_match.group(1).strip()
        for prefix in ["careers at ", "jobs at ", "recruiting at ", "talent at ",
                        "no-reply at ", "noreply at ", "hr at "]:
            if name.lower().startswith(prefix):
                name = name[len(prefix):]
        return name.strip()
    email_match = re.search(r'@([\w.-]+)', sender)
    if email_match:
        domain = email_match.group(1)
        company = domain.split(".")[0]
        return company.replace("-", " ").title()
    return sender.strip()


# Phrases that are definitely NOT job titles
BLOCKLIST = re.compile(
    r'^(?:an update on your|we have received your|thank you for your|' 
    r'i am looking for|what happens next|applying for the|'
    r'has been received|forwarded to the|an update|'
    r'we received your|your application|dear |hello |hi |'
    r'thanks for your|we have an update|don\'t forget)',
    re.IGNORECASE
)

# Strip trailing reference numbers e.g. "951668" or "- Ref 123" or "vacancy 259"
REF_CLEANUP = re.compile(
    r'(?:\s*[-–]?\s*(?:ref\.?|vacancy|job|req|id)[\s#]*\d+|\s+\d{4,})\s*$',
    re.IGNORECASE
)

# Valid job title: 2–6 words, starts capital, no digits
TITLE_PATTERN = re.compile(r'^[A-Z][a-zA-Z]+(?:\s+(?:and\s+)?[A-Za-z]+){1,5}$')


def is_valid_role(text: str) -> bool:
    text = REF_CLEANUP.sub("", text.strip(" .,()-–"))
    if not text or len(text) < 4 or len(text) > 70:
        return False
    if BLOCKLIST.match(text):
        return False
    if re.search(r'\d', text):   # reject anything with numbers
        return False
    if not TITLE_PATTERN.match(text):
        return False
    words = text.split()
    capitalised = sum(1 for w in words if w[0].isupper())
    if capitalised / len(words) < 0.5:
        return False
    return True


def extract_job_role(subject: str, body: str) -> str:
    subject_patterns = [
        r'application (?:for|to)(?: the)?\s+(.+?)(?:\s*[-|@–]|$)',
        r'(?:for the\s+)?(.+?)\s+(?:role|position|opportunity|vacancy)\b',
        r'(?:re:\s*)?your application[:\s]+(.+?)(?:\s*[-|@–]|$)',
        r'(?:applying for(?:\s+the)?\s+)(.+?)(?:\s*[-|@–]|$)',
        r'(?:confirmation[:\s]+)(.+?)(?:\s*[-|@–]|$)',
    ]
    for pattern in subject_patterns:
        m = re.search(pattern, subject, re.IGNORECASE)
        if m:
            role = REF_CLEANUP.sub("", m.group(1).strip(" .,()-–"))
            if is_valid_role(role):
                return role

    body_patterns = [
        r'(?:applied for|application for|applying for)(?:\s+the)?\s+["\']?([A-Z][^\n"\']{{3,60}}?)(?:["\']|\s*(?:role|position|job|post)\b)',
        r'(?:position|role|job title)[:\s]+["\']?([A-Z][^\n"\']{{3,60}}?)(?:["\']|[\n,.])',
        r'(?:the\s+)([A-Z][a-zA-Z]+(?:\s+[A-Z]?[a-zA-Z]+){{1,4}})\s+(?:position|role|vacancy|post)\b',
    ]
    for pattern in body_patterns:
        m = re.search(pattern, body, re.IGNORECASE)
        if m:
            role = REF_CLEANUP.sub("", m.group(1).strip(" .,()-–"))
            if is_valid_role(role):
                return role

    return "N/A"


def parse_email(service, msg_id: str) -> dict | None:
    msg = service.users().messages().get(
        userId="me", id=msg_id, format="full"
    ).execute()

    headers  = {h["name"].lower(): h["value"] for h in msg["payload"].get("headers", [])}
    subject  = headers.get("subject", "")
    sender   = headers.get("from", "")
    date_str = headers.get("date", "")

    try:
        clean_date = re.sub(r'\s*\([^)]*\)', '', date_str).strip()
        from email.utils import parsedate_to_datetime
        date_obj = parsedate_to_datetime(clean_date).date()
    except Exception:
        date_obj = datetime.today().date()

    company  = extract_company_from_sender(sender)
    body     = get_email_body(msg["payload"])
    job_role = extract_job_role(subject, body)
    status   = detect_status(subject, body)

    return {"date": date_obj, "company": company, "job_role": job_role, "status": status}

# ── Excel helpers ─────────────────────────────────────────────────────────────

HEADERS    = ["Date", "Company", "Job Role", "Status"]
COL_WIDTHS = [18, 25, 30, 15]

STATUS_STYLES = {
    "Pending":  {"color": "C47A1E"},  # orange
    "Rejected": {"color": "C0392B"},  # red
}


def _header_style(cell):
    cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    cell.fill      = PatternFill("solid", start_color="2E4057")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="FFFFFF")
    cell.border    = Border(left=thin, right=thin, top=thin, bottom=thin)


def _apply_headers(ws):
    ws.append(HEADERS)
    for col, cell in enumerate(ws[1], 1):
        _header_style(cell)
        ws.column_dimensions[cell.column_letter].width = COL_WIDTHS[col - 1]


def ensure_workbook() -> Workbook:
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        if SHEET_NAME not in wb.sheetnames:
            ws = wb.create_sheet(SHEET_NAME)
            _apply_headers(ws)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        _apply_headers(ws)
        ws.freeze_panes = "A2"
    return wb


def load_existing_rows(wb: Workbook) -> dict:
    """Return {company_lower: row_number} for duplicate + status-update checks."""
    ws = wb[SHEET_NAME]
    rows = {}
    for r in range(2, ws.max_row + 1):
        val = ws.cell(row=r, column=2).value
        if val:
            rows[str(val).strip().lower()] = r
    return rows


def update_status_in_row(ws, row_num: int, new_status: str):
    """Overwrite the Status cell of an existing row."""
    style = STATUS_STYLES.get(new_status, {"color": "2E4057"})
    cell = ws.cell(row=row_num, column=4)
    cell.value     = new_status
    cell.font      = Font(name="Arial", size=10, bold=True, color=style["color"])
    cell.alignment = Alignment(horizontal="center", vertical="center")


def append_row(ws, entry: dict):
    ws.append([entry["date"], entry["company"], entry["job_role"], entry["status"]])
    r = ws.max_row
    fill_color = "F0F4F8" if r % 2 == 0 else "FFFFFF"
    fill = PatternFill("solid", start_color=fill_color)
    for col in range(1, 5):
        cell = ws.cell(row=r, column=col)
        cell.fill = fill
        cell.font = Font(name="Arial", size=10)
        cell.alignment = Alignment(vertical="center")
    ws.cell(row=r, column=1).number_format = "DD/MM/YYYY"
    style = STATUS_STYLES.get(entry["status"], {"color": "2E4057"})
    status_cell = ws.cell(row=r, column=4)
    status_cell.font      = Font(name="Arial", size=10, bold=True, color=style["color"])
    status_cell.alignment = Alignment(horizontal="center", vertical="center")

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("🔐 Authenticating with Gmail...")
    service = get_gmail_service()
    print("✅ Authenticated.\n")

    print("📬 Fetching matching emails...")
    messages = fetch_all_messages(service)
    print(f"   Found {len(messages)} matching emails.\n")

    wb = ensure_workbook()
    ws = wb[SHEET_NAME]
    existing_rows = load_existing_rows(wb)

    added   = 0
    updated = 0
    skipped = 0

    for msg_meta in messages:
        try:
            entry = parse_email(service, msg_meta["id"])
            if entry is None:
                continue
            company_key = entry["company"].strip().lower()

            if company_key in existing_rows:
                if entry["status"] == "Rejected":
                    row_num = existing_rows[company_key]
                    current = ws.cell(row=row_num, column=4).value
                    if current != "Rejected":
                        update_status_in_row(ws, row_num, "Rejected")
                        updated += 1
                        print(f"   🔄 Updated to Rejected: {entry['company']}")
                    else:
                        skipped += 1
                        print(f"   ⏭  Skipped (already Rejected): {entry['company']}")
                else:
                    skipped += 1
                    print(f"   ⏭  Skipped (duplicate): {entry['company']}")
            else:
                append_row(ws, entry)
                existing_rows[company_key] = ws.max_row
                added += 1
                icon = "❌" if entry["status"] == "Rejected" else "✅"
                print(f"   {icon} Added: {entry['company']} — {entry['job_role']} ({entry['date']}) [{entry['status']}]")
        except Exception as e:
            print(f"   ⚠️  Error processing email {msg_meta['id']}: {e}")

    wb.save(EXCEL_FILE)
    print(f"\n📊 Done! {added} new rows added, {updated} statuses updated to Rejected, {skipped} skipped.")
    print(f"   Saved to: {os.path.abspath(EXCEL_FILE)}")


if __name__ == "__main__":
    main()