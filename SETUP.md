# Job Application Tracker — Setup Guide

## What you'll need
- Python 3.10+
- A Google account (the Gmail you use for job apps)
- ~10 minutes for first-time setup

---

## Step 1 — Install dependencies

```bash
pip install -r requirements.txt
```

---

## Step 2 — Enable the Gmail API & get credentials

1. Go to [https://console.cloud.google.com](https://console.cloud.google.com)
2. Create a **new project** (name it anything, e.g. "Job Tracker")
3. In the left menu → **APIs & Services** → **Library**
4. Search for **Gmail API** → click it → click **Enable**
5. Go to **APIs & Services** → **OAuth consent screen**
   - Choose **External** → click Create
   - Fill in App name (e.g. "Job Tracker"), your email → Save and Continue
   - Skip Scopes → Save and Continue
   - Add your Gmail address as a **Test user** → Save and Continue
6. Go to **APIs & Services** → **Credentials**
   - Click **+ Create Credentials** → **OAuth client ID**
   - Application type: **Desktop app**
   - Name it anything → click **Create**
7. Click **Download JSON** on the newly created credential
8. Rename the downloaded file to **`credentials.json`**
9. Place `credentials.json` in the **same folder** as `sync_jobs.py`

---

## Step 3 — Run the script

```bash
python sync_jobs.py
```

- The first time, a browser window will open asking you to log in to Google and grant read access to Gmail.
- After you approve, a `token.pickle` file is saved so you won't need to log in again.
- The script will scan your inbox and create (or update) **`job_applications.xlsx`** in the same folder.

---

## Step 4 — Fill in the blanks

The script fills in: **Date**, **Company**, **Job Role**.

Open `job_applications.xlsx` and manually fill in:
- **Salary** — from the job posting
- **Status** — e.g. Applied, Interview, Offer, Rejected
- **Location** — city / remote

---

## Running it again later

Just run `python sync_jobs.py` again — it will skip companies already in your spreadsheet and only add new ones.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `credentials.json not found` | Make sure it's in the same folder as the script |
| `Access blocked` in browser | Add your Gmail as a Test User in OAuth consent screen |
| Company name looks wrong | Edit the company name directly in Excel; the script matches on company name to avoid duplicates |
| Token expired | Delete `token.pickle` and run again to re-authenticate |
| Script finds 0 emails | Your acknowledgment emails may use different wording — open `sync_jobs.py` and add keywords to `GMAIL_QUERY` near the top |

---

## File structure

```
job_tracker/
├── sync_jobs.py          ← main script
├── requirements.txt      ← Python dependencies
├── credentials.json      ← downloaded from Google Cloud (you add this)
├── token.pickle          ← auto-generated after first login
└── job_applications.xlsx ← your tracker (auto-created or existing)
```
