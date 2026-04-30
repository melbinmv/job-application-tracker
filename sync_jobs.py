"""
Job Application Tracker — Gmail → Excel Sync
Reads acknowledgment emails from Gmail and appends new entries to your Excel tracker.
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

EXCEL_FILE = "job_applications.xlsx"
SHEET_NAME = "Applications"
TOKEN_FILE = "token.pickle"
CREDENTIALS_FILE = "credentials.json"
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

GMAIL_QUERY = (
    'subject:("application received" OR "thank you for applying" OR '
    '"thanks for applying" OR "we received your application" OR '
    '"application confirmation" OR "your application" OR '
    '"application for" OR "thank you for your interest")'
)

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


def extract_job_role(subject: str, body: str) -> str:
    subject_patterns = [
        r'application (?:for|to)(?: the)? (.+?)(?:\s*[-|@]|$)',
        r'(?:your |re: )?(.+?) application',
        r'(?:for|re:)\s+(.+?)\s+(?:role|position|job|opportunity)',
    ]
    for pattern in subject_patterns:
        m = re.search(pattern, subject, re.IGNORECASE)
        if m:
            role = m.group(1).strip(" .,")
            if 3 < len(role) < 80:
                return role

    body_snippet = body[:1500]
    body_patterns = [
        r'(?:applied for|application for|position[:\s]+|role[:\s]+|job title[:\s]+)\s*["\']?([A-Z][^\n"\']{3,60})',
        r'(?:the\s+)?([A-Z][a-z]+(?:\s+[A-Z]?[a-z]+){1,5})\s+(?:position|role|opportunity)',
    ]
    for pattern in body_patterns:
        m = re.search(pattern, body_snippet)
        if m:
            role = m.group(1).strip(" .,")
            if 3 < len(role) < 80:
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

    return {"date": date_obj, "company": company, "job_role": job_role}

# ── Excel helpers ─────────────────────────────────────────────────────────────

HEADERS    = ["Date", "Company", "Job Role", "Status"]
COL_WIDTHS = [18, 25, 30, 15]


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


def load_existing_companies(wb: Workbook) -> set:
    ws = wb[SHEET_NAME]
    return {
        str(ws.cell(row=r, column=2).value or "").strip().lower()
        for r in range(2, ws.max_row + 1)
        if ws.cell(row=r, column=2).value
    }


def append_row(ws, entry: dict):
    ws.append([
        entry["date"],
        entry["company"],
        entry["job_role"],
        "Pending",
    ])
    r = ws.max_row
    fill_color = "F0F4F8" if r % 2 == 0 else "FFFFFF"
    fill = PatternFill("solid", start_color=fill_color)
    for col in range(1, 5):
        cell = ws.cell(row=r, column=col)
        cell.fill = fill
        cell.font = Font(name="Arial", size=10)
        cell.alignment = Alignment(vertical="center")
    ws.cell(row=r, column=1).number_format = "DD/MM/YYYY"
    # Status — bold orange to flag as actionable
    status_cell = ws.cell(row=r, column=4)
    status_cell.font      = Font(name="Arial", size=10, bold=True, color="C47A1E")
    status_cell.alignment = Alignment(horizontal="center", vertical="center")

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("🔐 Authenticating with Gmail...")
    service = get_gmail_service()
    print("✅ Authenticated.\n")

    print("📬 Fetching acknowledgment emails...")
    results = service.users().messages().list(
        userId="me", q=GMAIL_QUERY, maxResults=200
    ).execute()
    messages = results.get("messages", [])
    print(f"   Found {len(messages)} matching emails.\n")

    wb = ensure_workbook()
    ws = wb[SHEET_NAME]
    existing = load_existing_companies(wb)

    added   = 0
    skipped = 0

    for msg_meta in messages:
        try:
            entry = parse_email(service, msg_meta["id"])
            if entry is None:
                continue
            company_key = entry["company"].strip().lower()
            if company_key in existing:
                skipped += 1
                print(f"   ⏭  Skipped (duplicate): {entry['company']}")
            else:
                append_row(ws, entry)
                existing.add(company_key)
                added += 1
                print(f"   ✅ Added: {entry['company']} — {entry['job_role']} ({entry['date']})")
        except Exception as e:
            print(f"   ⚠️  Error processing email {msg_meta['id']}: {e}")

    wb.save(EXCEL_FILE)
    print(f"\n📊 Done! {added} new rows added, {skipped} duplicates skipped.")
    print(f"   Saved to: {os.path.abspath(EXCEL_FILE)}")


if __name__ == "__main__":
    main()