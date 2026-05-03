"""
Job Application Tracker — Gmail → Excel Sync (v2 — Auto-Fixed)
Reads acknowledgment/rejection emails from Gmail and appends new entries to your Excel tracker.
"""

import os
import re
import base64
import pickle
from datetime import datetime
from html import unescape as html_unescape

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
    '"application feedback" OR "application status" OR '
    '"application update" OR "regarding your application" OR '
    '"following your application" OR "update on your application" OR '
    '"unfortunately" OR "not been successful" OR "not move forward" OR '
    '"position has been filled" OR "we will not" OR "regret to inform" OR '
    '"not be progressing" OR "will not be progressing" OR '
    '"decided not to proceed" OR "not shortlisted")'
)


REJECTION_PATTERNS = [
    # Original patterns
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
    r'\bpursue other candidates\b',
    # ── NEW patterns ──
    r'\bwon.?t\s+be\s+progress',                           # "won't be progressing"
    r'\bnot\s+(?:be\s+)?progress(?:ing|ed)\b',             # "not progressing" / "not be progressing"
    r'\bwill\s+not\s+be\s+(?:progressing|proceeding|continuing)\b',
    r'\bunable\s+to\s+(?:offer|proceed|progress)\b',       # "unable to offer/proceed"
    r'\bon\s+this\s+occasion\b',                           # "on this occasion"
    r'\bnot\s+(?:be\s+)?(?:taking|moving)\s+.{0,30}(?:further|forward)\b',  # "not be taking...further"
    r'\bmoved?\s+forward\s+with\s+(?:other|another)\b',    # "moved forward with other candidates"
    r'\bother\s+candidates?\s+(?:whose|who|that|more)\b',  # "other candidates whose..."
    r'\bnot\s+(?:be\s+)?proceed(?:ing)?\b',                # "not proceed" / "not be proceeding"
    r'\bafter\s+careful\s+(?:consideration|review).{0,120}(?:not|won|regret|unfortunately|unable)\b',
    r'\bwe\s+(?:are|have)\s+(?:decided|chosen)\s+(?:not\s+)?to\s+(?:not\s+)?(?:proceed|progress|continue|move)\b',
    r'\bapplication\s+(?:has\s+been\s+|was\s+)?unsuccessful\b',
    r'\bnot\s+(?:the\s+)?right\s+(?:fit|match)\b',         # "not the right fit"
    r'\bfilled\s+(?:the\s+)?(?:position|role|vacancy)\b',   # "filled the position"
    r'\bwithdr(?:awn?|ew)\b',                               # "withdrawn"
    r'\brejected\b',                                        # explicit "rejected"
    r'\bunsuccessful\b',                                    # explicit "unsuccessful"
]

def detect_status(subject: str, body: str, snippet: str = "") -> str:
    """Check subject + body + snippet for rejection signals."""
    text = (subject + " " + body[:3000] + " " + snippet).lower()
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

def fetch_all_messages(service):
    """Fetch the latest 200 matching messages."""
    response = service.users().messages().list(
        userId="me", q=GMAIL_QUERY, maxResults=200
    ).execute()
    return response.get("messages", [])


def _decode_part_data(data: str, encoding: str = "") -> str:
    if not data:
        return ""
    try:
        raw_bytes = base64.urlsafe_b64decode(data)
        if encoding.lower() == "quoted-printable":
            import quopri
            raw_bytes = quopri.decodestring(raw_bytes)
        return raw_bytes.decode("utf-8", errors="ignore")
    except Exception:
        return ""


def _strip_html(html: str) -> str:
    """Strip HTML tags and decode ALL HTML entities (named + numeric)."""
    # Remove style and script blocks entirely
    text = re.sub(r'<style[^>]*>.*?</style>', ' ', html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<script[^>]*>.*?</script>', ' ', text, flags=re.DOTALL | re.IGNORECASE)
    # Strip tags
    text = re.sub(r'<[^>]+>', ' ', text)
    # Use Python's html.unescape to handle ALL entities: &nbsp; &rsquo; &#39; &#x27; etc.
    text = html_unescape(text)
    # Replace non-breaking spaces (which unescape converts to \xa0) with regular spaces
    text = text.replace('\xa0', ' ')
    # Remove zero-width characters
    text = re.sub(r'[\u200b\u200c\u200d\ufeff]', '', text)
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def get_email_body(msg_payload):
    """Extract readable plain text from email payload."""
    mime     = msg_payload.get("mimeType", "")
    data     = msg_payload.get("body", {}).get("data", "")
    headers  = {h["name"].lower(): h["value"] for h in msg_payload.get("headers", [])}
    encoding = headers.get("content-transfer-encoding", "")

    if mime == "text/plain" and data:
        return _decode_part_data(data, encoding)

    if mime == "text/html" and data:
        return _strip_html(_decode_part_data(data, encoding))

    plain, html = "", ""
    for part in msg_payload.get("parts", []):
        part_mime = part.get("mimeType", "")
        part_data = part.get("body", {}).get("data", "")
        part_headers = {h["name"].lower(): h["value"] for h in part.get("headers", [])}
        part_encoding = part_headers.get("content-transfer-encoding", "")

        if part_mime == "text/plain" and part_data:
            plain = _decode_part_data(part_data, part_encoding)
        elif part_mime == "text/html" and part_data:
            html = _strip_html(_decode_part_data(part_data, part_encoding))
        else:
            result = get_email_body(part)
            if result and not plain:
                plain = result

    return plain or html or ""


def _extract_domain_company(email_addr: str) -> str:
    """Extract company name from an email address domain."""
    m = re.search(r'@([\w.-]+)', email_addr)
    if m:
        domain = m.group(1).lower()
        # Skip generic ATS / email service domains
        ATS_DOMAINS = {
            'talosats', 'mailgun', 'greenhouse', 'lever', 'workday',
            'icims', 'smartrecruiters', 'bamboohr', 'ashbyhq', 'jobvite',
            'recruitee', 'applytojob', 'myworkdayjobs', 'successfactors',
            'breezyhr', 'jazz', 'gmail', 'outlook', 'hotmail', 'yahoo',
            'googlemail', 'zoho', 'mail', 'noreply', 'no-reply',
        }
        company = domain.split(".")[0]
        if company.lower() in ATS_DOMAINS:
            return ""
        return company.replace("-", " ").title()
    return ""


def extract_company_from_sender(sender: str, reply_to: str = "") -> str:
    """
    Extract company name from sender/reply-to headers.

    Strategy:
      1. If display name is an email (ATS pattern like '"careers@altro.com" <no-reply@ats.com>'),
         extract domain from the display-name email.
      2. If display name is a human-readable name, strip common prefixes.
      3. Use Reply-To domain as fallback.
      4. Use From email domain as last resort.
    """
    name_match = re.match(r'^"?([^"<]+)"?\s*<', sender)
    if name_match:
        name = name_match.group(1).strip()

        # NEW: Check if display name is itself an email address
        if re.match(r'^[\w.+%-]+@[\w.-]+$', name):
            company = _extract_domain_company(name)
            if company:
                return company
            # If domain was an ATS, fall through

        # Original: strip common prefixes
        PREFIXES = [
            "careers at ", "jobs at ", "recruiting at ", "talent at ",
            "no-reply at ", "noreply at ", "hr at ", "hiring at ",
            "recruitment at ", "people at ", "team at ",
        ]
        for prefix in PREFIXES:
            if name.lower().startswith(prefix):
                name = name[len(prefix):]
                break

        # Strip trailing " Hiring Team", " Careers", " Recruitment", " Team"
        name = re.sub(
            r'\s*(?:Hiring\s+Team|Careers|Recruitment|Talent\s+(?:Team|Acquisition)|Team|HR|DoNotReply|NoReply|No-Reply)\s*$',
            '', name, flags=re.IGNORECASE
        ).strip()

        # If we still have a usable name (not empty, not just "noreply" etc.)
        if name and not re.match(r'^(?:no-?reply|noreply|donotreply|info|support|admin|system)$', name, re.IGNORECASE):
            return name

    # Fallback: try Reply-To header
    if reply_to:
        company = _extract_domain_company(reply_to)
        if company:
            return company

    # Last resort: From email domain
    company = _extract_domain_company(sender)
    return company or sender.strip()


# ── FIX #5: Job role extraction ───────────────────────────────────────────────

BLOCKLIST = re.compile(
    r'^(?:an update on your|we have received your|thank you for your|'
    r'i am looking for|what happens next|applying for the|'
    r'has been received|forwarded to the|an update|'
    r'we received your|your application|dear |hello |hi |'
    r'thanks for your|we have an update|don\'t forget|'
    r'this is to confirm|please note|we are writing|'
    r'we would like|on this occasion|after careful)',
    re.IGNORECASE
)

REF_CLEANUP = re.compile(
    r'(?:\s*[-–—]?\s*(?:ref\.?|vacancy|job|req|id|reference)[\s#:]*\d+|\s+\d{4,})\s*$',
    re.IGNORECASE
)

GENERIC_WORDS = {
    'application', 'update', 'confirmation', 'recruitment', 'feedback',
    'opportunity', 'vacancy', 'position', 'interview', 'hiring',
    'team', 'group', 'company', 'organisation', 'organization',
    'status', 'notification', 'response', 'acknowledgement',
    'acknowledgment', 'receipt', 'received', 'submitted',
    'the', 'your', 'our', 'this', 'that', 'with', 'from',
    'application feedback', 'application status', 'application update',
}


def is_valid_role(text: str) -> bool:
    text = REF_CLEANUP.sub("", text.strip(" .,()-–—:"))
    if not text or len(text) < 4 or len(text) > 80:
        return False
    if re.search(r'\d', text):
        return False
    if BLOCKLIST.match(text):
        return False
    if text.lower().strip() in GENERIC_WORDS:
        return False

    # More relaxed pattern: allow hyphenated words, slashes, parentheses
    # e.g. "Full-Stack Developer", "AI/ML Engineer", "Engineer (Backend)"
    words = text.split()
    if len(words) < 2 or len(words) > 8:
        return False

    # At least first word should be capitalized or an abbreviation
    first = words[0]
    if not (first[0].isupper() or first.isupper()):
        return False

    # At least 40% of words should be capitalised
    capitalised = sum(1 for w in words if w[0].isupper())
    if capitalised / len(words) < 0.4:
        return False

    return True


def extract_job_role(subject: str, body: str, snippet: str = "") -> str:
    """
    Extract job role from subject, then body, then snippet.
    Uses many more patterns and relaxed validation.
    """

    # ── Subject patterns ──────────────────────────────────────────────────
    subject_patterns = [
        # "Application for Data Analyst" / "Application for the role of X"
        r'\bapplication\s+for\s+(?:the\s+)?(?:role\s+of\s+)?(?:the\s+)?([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})',
        # "Junior Integration Analyst Application Feedback" — role BEFORE "Application"
        r'^([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\s+Application\b',
        # "Applying for Data Analyst role/position"
        r'\bapplying\s+for\s+(?:the\s+)?([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\s+(?:role|position|post|vacancy)\b',
        # "[Role] role/position" — title before keyword
        r'\b([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\s+(?:role|position|post|vacancy)\b',
        # "Confirmation: Data Analyst"
        r':\s*([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\s*$',
        # "Re: Data Analyst" — sometimes role is the whole subject after Re:
        r'^(?:Re:\s*)?([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\s*$',
    ]
    for pattern in subject_patterns:
        m = re.search(pattern, subject)
        if m:
            role = REF_CLEANUP.sub("", m.group(1).strip(" .,()-–—:"))
            # Remove trailing "Application" if captured
            role = re.sub(r'\s+Application$', '', role, flags=re.IGNORECASE).strip()
            if is_valid_role(role):
                return role

    # ── Body patterns ─────────────────────────────────────────────────────
    # Search body AND snippet
    search_texts = [body[:3000], snippet]

    body_patterns = [
        # "applied for the AI Automations Product Engineer position/role/job"
        r'\b(?:applied|apply|applying)\s+for\s+(?:the\s+)?([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\s+(?:role|position|post|job|vacancy)\b',
        # "apply for the role of Junior Integration Analyst"
        r'\b(?:applied|apply|applying)\s+for\s+the\s+role\s+of\s+([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\b',
        # "apply for the position of Senior Data Analyst"
        r'\b(?:applied|apply|applying)\s+for\s+the\s+(?:position|post)\s+of\s+([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\b',
        # "interest in the Junior ML Engineer position/role"
        r'\binterest\s+in\s+(?:the\s+)?([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\s+(?:role|position|post|vacancy)\b',
        # "position: Data Analyst" or "role: Senior Engineer" or "job title: X"
        r'\b(?:position|role|job\s+title|vacancy)\s*[:\-–]\s*([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\b',
        # "the Staff Data Analyst role/position"
        r'\bthe\s+([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\s+(?:position|role|post|vacancy)\b',
        # "response to the Data Analyst position"
        r'\bresponse\s+to\s+the\s+([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\s+(?:position|role|post|vacancy)\b',
        # "application for the role of X"  (body version, slightly different from subject)
        r'\bapplication\s+for\s+(?:the\s+)?(?:role|position|post)\s+of\s+([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\b',
        # "application for Data Analyst"
        r'\bapplication\s+for\s+(?:the\s+)?([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\b',
        # "regarding the Data Analyst vacancy"
        r'\bregarding\s+(?:the\s+)?([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\s+(?:role|position|post|vacancy|opening)\b',
        # "you applied for: Data Analyst"
        r'\b(?:applied|apply|applying)\s+for\s*:\s*([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\b',
        # "role of Junior Integration Analyst" (generic)
        r'\brole\s+of\s+([A-Z][A-Za-z/&\-]+(?:\s+[A-Za-z/&\-]+){1,5})\b',
    ]
    for text in search_texts:
        if not text:
            continue
        for pattern in body_patterns:
            m = re.search(pattern, text, re.IGNORECASE)
            if m:
                role = REF_CLEANUP.sub("", m.group(1).strip(" .,()-–—:"))
                role = re.sub(r'\s+Application$', '', role, flags=re.IGNORECASE).strip()
                if is_valid_role(role):
                    return role

    return "N/A"


# ── Email parsing ─────────────────────────────────────────────────────────────

def parse_email(service, msg_id: str) -> dict | None:
    msg = service.users().messages().get(
        userId="me", id=msg_id, format="full"
    ).execute()

    headers  = {h["name"].lower(): h["value"] for h in msg["payload"].get("headers", [])}
    subject  = headers.get("subject", "")
    sender   = headers.get("from", "")
    reply_to = headers.get("reply-to", "")
    date_str = headers.get("date", "")

    # Gmail snippet — pre-parsed plain text, always available
    snippet = msg.get("snippet", "")

    try:
        clean_date = re.sub(r'\s*\([^)]*\)', '', date_str).strip()
        from email.utils import parsedate_to_datetime
        date_obj = parsedate_to_datetime(clean_date).date()
    except Exception:
        date_obj = datetime.today().date()

    # FIX #4: pass reply_to to company extraction
    company  = extract_company_from_sender(sender, reply_to)
    body     = get_email_body(msg["payload"])

    # FIX #5: pass snippet as fallback for role extraction
    job_role = extract_job_role(subject, body, snippet)

    # FIX #6: pass snippet as fallback for status detection
    status   = detect_status(subject, body, snippet)

    return {
        "date": date_obj,
        "company": company,
        "job_role": job_role,
        "status": status,
    }

# ── Excel helpers ─────────────────────────────────────────────────────────────

HEADERS    = ["Date", "Company", "Job Role", "Status"]
COL_WIDTHS = [18, 25, 30, 15]

STATUS_STYLES = {
    "Pending":  {"color": "C47A1E"},
    "Rejected": {"color": "C0392B"},
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


# ── FIX #7: Dedup on (company, role) instead of just company ─────────────────

def load_existing_rows(wb: Workbook) -> dict:
    """Return {(company_lower, role_lower): row_number} for dedup."""
    ws = wb[SHEET_NAME]
    rows = {}
    for r in range(2, ws.max_row + 1):
        company_val = ws.cell(row=r, column=2).value
        role_val    = ws.cell(row=r, column=3).value
        if company_val:
            key = (
                str(company_val).strip().lower(),
                str(role_val).strip().lower() if role_val else "n/a"
            )
            rows[key] = r
    return rows


def update_status_in_row(ws, row_num: int, new_status: str):
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

            dedup_key = (
                entry["company"].strip().lower(),
                entry["job_role"].strip().lower()
            )

            if dedup_key in existing_rows:
                if entry["status"] == "Rejected":
                    row_num = existing_rows[dedup_key]
                    current = ws.cell(row=row_num, column=4).value
                    if current != "Rejected":
                        update_status_in_row(ws, row_num, "Rejected")
                        updated += 1
                        print(f"   🔄 Updated to Rejected: {entry['company']} — {entry['job_role']}")
                    else:
                        skipped += 1
                else:
                    skipped += 1
            else:
                append_row(ws, entry)
                existing_rows[dedup_key] = ws.max_row
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