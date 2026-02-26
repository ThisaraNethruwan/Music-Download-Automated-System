# 🎵 AutoScraper — Email-to-Music Downloader

> An end-to-end Python automation pipeline that monitors a Gmail inbox, scrapes music track listings from a web portal, cross-checks them against a SQL Server database, logs everything to Google Sheets, and automatically downloads new WAV files — all without human intervention.

---

## Overview

AutoScraper is a fully automated music discovery and download pipeline built for music libraries, A&R teams, or anyone who receives regular promotional emails with links to new track listings. Instead of manually clicking through every email and download portal, AutoScraper handles the entire chain from inbox to local WAV file.

---

## How It Works

```
📧 Gmail Inbox
      ↓  (unread email from target sender)
🔗 Extract "Get Now" Link from Email HTML
      ↓
🌐 Selenium: Navigate & Log In to Web Portal
      ↓
🎵 Scrape Track List (Title + Artist)
      ↓
🗄️ SQL Server: Check if each track already exists in DB
      ↓
📊 Google Sheets: Log all tracks + DB status
      ↓
⬇️  Selenium: Download WAV for tracks NOT in DB
      ↓
📂 Update Google Sheets with local file paths
      ↓
✅ Mark email as read + log run stats
```

---

## Features

- **Automated Gmail Monitoring** — Filters unread emails from a specific sender and extracts portal links
- **Intelligent Web Scraping** — Primary DOM scraping with an automatic regex fallback for different page layouts
- **Auto Login** — Detects and fills username/password fields on the web portal automatically
- **Database Deduplication** — Queries a SQL Server `Tracks` table to skip music already in your library
- **Google Sheets Logging** — Writes every scraped track (title, artist, DB status, file path) to a spreadsheet in real time
- **Selective WAV Downloads** — Only downloads tracks flagged as "not in DB", saving time and storage
- **File Path Tracking** — After download, matches each local file back to its row in Google Sheets
- **Run Logging** — Appends execution stats (emails, URLs, tracks, downloads) to a separate log sheet per run
- **Duplicate URL Prevention** — Skips any portal URL already present in the spreadsheet
- **Download Completion Detection** — Waits for `.crdownload` temp files to disappear before proceeding

---

## Tech Stack

| Layer | Technology |
|---|---|
| Language | Python 3.x |
| Browser Automation | Selenium + ChromeDriver |
| Email Access | Google Gmail API v1 |
| Spreadsheet | Google Sheets API v4 |
| Database | Microsoft SQL Server via `pyodbc` |
| HTML Parsing | BeautifulSoup4 |
| Auth | Google OAuth 2.0 |

---

## Prerequisites

- Python 3.8+
- Google Chrome browser installed
- ChromeDriver matching your Chrome version ([download here](https://chromedriver.chromium.org/downloads))
- Microsoft SQL Server with ODBC Driver 18 installed
- A Google Cloud project with Gmail API and Google Sheets API enabled
- Access to the target music web portal

---

**3. Install dependencies**


```
google-auth
google-auth-oauthlib
google-auth-httplib2
google-api-python-client
selenium
beautifulsoup4
pyodbc
```

---

## Configuration

Open `AutoScraper.py` and update the configuration block near the top of the file:

```python
# ── Google Sheets ──────────────────────────────────────────────
SPREADSHEET_ID  = 'YOUR_SPREADSHEET_ID'   # From the Google Sheets URL
SHEET_NAME      = 'Sheet_Name'            # Main data tab name
SIGNATURE_SHEET = 'AppSignature'          # Run log tab name

# ── Web Portal Login ───────────────────────────────────────────
LOGIN_USERNAME = 'YOUR_LOGIN_USERNAME'
LOGIN_PASSWORD = 'YOUR_LOGIN_PASSWORD'
SENDER_EMAIL   = 'YOUR_SENDER_EMAIL'      # Email address to monitor

# ── SQL Server ─────────────────────────────────────────────────
SQL_SERVER           = r'YOUR_SERVER_NAME'
SQL_DATABASE         = 'YOUR_DATABASE_NAME'
SQL_TRUSTED_CONNECTION = True             # Set False if using SQL auth

# ── Downloads ──────────────────────────────────────────────────
BASE_DOWNLOAD_DIR = os.path.join(os.getcwd(), 'downloads')
```

---

## Google API Setup

**Step 1 — Create a Google Cloud Project**

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Enable **Gmail API** and **Google Sheets API** for the project

**Step 2 — Create OAuth 2.0 Credentials**

1. Navigate to **APIs & Services → Credentials**
2. Click **Create Credentials → OAuth client ID**
3. Choose **Desktop app**
4. Download the file and rename it to `credentials.json`
5. Place `credentials.json` in the same directory as `AutoScraper.py`

**Step 3 — First Run Authentication**

On the first run, a browser window will open asking you to authorise access to your Google account. After approval, a `token.json` file is automatically saved and reused for all future runs.

> Add both `credentials.json` and `token.json` to your `.gitignore`.

---

## Project Structure

```
autoscraper/
├── AutoScraper.py        # Main application script
├── credentials.json      # Google OAuth credentials (do not commit)
├── token.json            # Auto-generated auth token (do not commit)
├── downloads/            # Auto-created; WAV files stored here by run ID
│   └── <unique_id>/
│       └── track.wav
├── requirements.txt
├── .gitignore
└── README.md
```

---

## Google Sheets Structure

The script automatically creates and maintains two tabs in your spreadsheet.

**Main Sheet (`Sheet_Name`)**

| Column | Header | Description |
|---|---|---|
| A | Unique ID | 8-character UUID per email/run |
| B | URL | The scraped portal URL |
| C | Title | Track title |
| D | Artist | Artist name |
| E | In DB | `Yes` / `No` — whether track exists in SQL Server |
| F | Path | Absolute local file path after download |

**Log Sheet (`AppSignature`)**

| Column | Header | Description |
|---|---|---|
| A | Run Number | Auto-incrementing run counter |
| B | Timestamp | Date and time of run |
| C | Emails Processed | Total unread emails found |
| D | URLs Added | New portal URLs scraped |
| E | Tracks Extracted | Total tracks found across all URLs |
| F | Downloads | Number of WAV files downloaded |

---

## Running the Script

```bash
python AutoScraper.py
```

**Example console output:**

```
======================================================================
Email to Music Downloader - V3.0
======================================================================
Database connection established

 Processing Email ID: 1a2b3c4d5e6f...
   Found URL: https://portal.example.com/release/xyz
  → Setting up browser for scraping...
  → Logging in...
  → Scanning track list...
    • Midnight Drive - DJ Example [DB: No]
    • Summer Fade - The Artist [DB: Yes]
  → Reading sheet (Unique ID: a1b2c3d4)...
  → Found 1 track to download
  → Starting downloads...
   Processing: Midnight Drive
       Found 'Download WAV', clicking...
  Waiting for downloads to finalize...
  All downloads finished.
   Updated path: Midnight Drive.wav
   Complete! Verified 1 valid file(s).

FINAL SUMMARY: 1 Emails, 1 Downloads.
Database connection closed
```

---

## Workflow Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                        AutoScraper                          │
│                                                             │
│  Gmail API                                                  │
│  ┌──────────┐   unread emails    ┌───────────────────────┐  │
│  │  Inbox   │ ─────────────────► │  Extract "Get Now"    │  │
│  └──────────┘                    │  URL from HTML body   │  │
│                                  └──────────┬────────────┘  │
│                                             │               │
│  Selenium                                   ▼               │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  1. Navigate to portal URL                           │   │
│  │  2. Auto-detect & fill login form                    │   │
│  │  3. Scrape track list (DOM scan → regex fallback)    │   │
│  └──────────────────────────┬───────────────────────────┘   │
│                             │                               │
│  SQL Server                 ▼                               │
│  ┌──────────┐   check each track    ┌──────────────────┐    │
│  │  Tracks  │ ◄──────────────────── │  DB Lookup       │    │
│  │  Table   │ ──── Yes / No ──────► │  per track       │    │
│  └──────────┘                       └────────┬─────────┘    │
│                                              │              │
│  Google Sheets                               ▼              │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  Write: ID | URL | Title | Artist | In DB | Path     │   │
│  └──────────────────────────┬───────────────────────────┘   │
│                             │                               │
│  Selenium (Download Mode)   ▼                               │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  Re-login → Find rows where In DB = "No"             │   │
│  │  → Click context menu → "Download WAV"               │   │
│  │  → Wait for .crdownload files to clear               │   │
│  └──────────────────────────┬───────────────────────────┘   │
│                             │                               │
│                             ▼                               │
│              Update Sheet "Path" column                     │
│              Mark email as read                             │
│              Log run stats to AppSignature tab              │
└─────────────────────────────────────────────────────────────┘
```

---

## Troubleshooting

**`Database connection error`**
- Confirm ODBC Driver 18 for SQL Server is installed
- Verify `SQL_SERVER` and `SQL_DATABASE` values
- Check that Windows Authentication is enabled if using `SQL_TRUSTED_CONNECTION = True`

**`Error setting up Chrome driver`**
- Ensure ChromeDriver version matches your installed Chrome version
- Make sure `chromedriver` is in your system `PATH`

**Tracks not being found on the page**
- The portal's HTML structure may differ from what the CSS selectors target
- Enable visible Chrome (remove `--headless` if added) to observe what Selenium sees
- The regex fallback will attempt to extract tracks from raw page text automatically

**`No 'Get Now' link found`**
- The email HTML structure may have changed — inspect the raw email source and update the selector in `extract_press_play_url()`

**Google auth errors**
- Delete `token.json` and re-run to trigger a fresh OAuth flow
- Confirm the correct scopes are listed in your Google Cloud OAuth consent screen

---
