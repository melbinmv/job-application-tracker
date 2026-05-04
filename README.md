# Job Application Tracker 📊

An automated job application tracker that reads acknowledgment and rejection emails from Gmail, uses AI to extract the data, and logs everything into an Excel spreadsheet — no more manual data entry.

## How it works

1. **Gmail API** scans your inbox for job application related emails using keyword filtering
2. **Gemma 3 4B** (via Google AI Studio) reads each email in batches and extracts the company name, job role, and status
3. Clean entries are written to the **Applications** sheet in Excel
4. Emails where the company or job role couldn't be extracted are saved to a **Needs Review** sheet with the full email content for manual evaluation
5. Duplicate companies are skipped — if a rejection email arrives for an existing entry, the status is automatically updated to **Rejected**

## Output

### Applications Sheet
| Date | Company | Job Role | Status |
|------|---------|----------|--------|
| 20/11/2025 | Low Carbon Contracts Company | Energy Analyst Intern | Rejected |
| 23/11/2025 | Mungos | Trainee Asset Data Analyst | Pending |
| 18/11/2025 | Webitrent | HR Systems Analyst | Pending |

### Needs Review Sheet
Emails where company or job role couldn't be extracted — includes full email content for manual review.

## Tech Stack

- Python 3.10+
- Gmail API (Google Cloud) — email fetching
- Gemma 3 4B (Google AI Studio) — AI extraction
- `google-generativeai` — Gemini/Gemma API client
- `openpyxl` — Excel file generation
- `google-auth` — Gmail authentication

## Getting Started

See [SETUP.md](SETUP.md) for full step-by-step instructions including how to set up Gmail API credentials and run the script.

## Project Structure

```
job-application-tracker/
├── sync_jobs.py        ← main script
├── requirements.txt    ← Python dependencies
├── SETUP.md            ← setup instructions
└── README.md           ← you are here
```

## Configuration

Key settings at the top of `sync_jobs.py`:

```python
BATCH_SIZE   = 5               # emails per AI call
BATCH_DELAY  = 3               # seconds between batches
GEMINI_MODEL = "gemma-3-4b-it" # AI model used
```

## Environment Variables

```bash
export GEMINI_API_KEY="your_key_here"
```

Get your free API key from [Google AI Studio](https://aistudio.google.com).

## Notes

- The script only requests **read-only** access to Gmail — it never modifies or deletes emails
- `credentials.json` and `token.pickle` are excluded from this repo for security
- Fetches the latest **200** matching emails per run
- Uses **Gemma 3 4B** via Google AI Studio free tier (14,400 requests/day)

## Changelog

### Latest — Needs Review Sheet
- Emails where company or job role could not be extracted are now saved to a separate **Needs Review** sheet with full email content for manual evaluation
- Clean entries only appear in the main Applications sheet

### Previous — AI Extraction
- Replaced regex-based extraction with **Gemma 3 4B** AI model
- Regex is now used only as a filter to identify job-related emails
- AI handles all extraction: company name, job role, and status

### Previous — Regex Extraction ([commit](https://github.com/melbinmv/job-application-tracker/commit/10f6f4f4b02e3c21a3394ae550c67eac6a710818))
- Used regex patterns to extract company name and job role from emails