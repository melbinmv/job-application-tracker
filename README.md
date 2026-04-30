# Job Application Tracker 📊

An automated job application tracker that reads acknowledgment and rejection emails from Gmail and logs them into an Excel spreadsheet — no more manual data entry.

## What it does

- Connects to your Gmail account using the Gmail API
- Scans your inbox for job application acknowledgment and rejection emails
- Extracts the **company name**, **job role**, and **date applied** from each email
- Automatically sets the status to **Pending** for acknowledgments and **Rejected** for rejection emails
- Writes everything into a neatly formatted **Excel file**
- Skips duplicates — if a company already exists, it won't add it again
- If a rejection email arrives for an existing entry, it updates the status to **Rejected**

## Output

| Date | Company | Job Role | Status |
|------|---------|----------|--------|
| 20/11/2025 | Low Carbon Contracts Company | Energy Analyst Intern | Rejected |
| 23/11/2025 | Mungos | Trainee Asset Data Analyst | Pending |
| 18/11/2025 | Webitrent | HR Systems Analyst | Pending |

## Tech Stack

- Python 3.10+
- Gmail API (Google Cloud)
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

## Notes

- The script only requests **read-only** access to Gmail — it never modifies or deletes any emails
- `credentials.json` and `token.pickle` are excluded from this repo for security — you need to generate your own (see SETUP.md)
- Fetches the latest 200 matching emails per run