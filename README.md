# 🚀 B2B Lead Scraper & Cold Email Automation Pipeline

An end-to-end B2B outreach toolkit designed to scrape targeted business leads from Google Maps, discover verified company contact emails from their websites, and execute personalized cold email campaigns with automated Microsoft Outlook / IMAP Sent Items synchronization.

---

## 📌 Table of Contents

- [Overview](#overview)
- [Pipeline Architecture](#pipeline-architecture)
- [Prerequisites & Installation](#prerequisites--installation)
- [Step-by-Step Workflow](#step-by-step-workflow)
  - [Step 1: Scrape Google Maps Leads](#step-1-scrape-google-maps-leads)
  - [Step 2: Extract Website Contact Emails](#step-2-extract-website-contact-emails)
  - [Step 3: Send Personalized Emails](#step-3-send-personalized-emails)
- [Command-Line Options & Flags](#command-line-options--flags)
- [Configuration Reference](#configuration-reference)
- [Anti-Spam & Deliverability Architecture](#anti-spam--deliverability-architecture)
- [Outlook & IMAP Sent Sync Tool](#outlook--imap-sent-sync-tool)
- [Troubleshooting & FAQ](#troubleshooting--faq)

---

## 🔎 Overview

This repository automates the complete outbound sales funnel:

1. **Google Maps Scraper (`main.py` + `scraper.py`)**: Uses Playwright browser automation to scrape verified business listings based on any search query (e.g., `"Fintech companies in Gurgaon"` or `"BPO companies in Delhi"`). Collects business name, phone number, website, physical address, rating, and review count.
2. **Website Email Extractor (`email_extractor.py`)**: Traverses each discovered company website through multiple fallback strategies (HTML parsing, Cloudflare email de-obfuscation, contact page discovery, and headless Playwright rendering for JavaScript SPAs) to find real contact emails.
3. **Personalized Mail-Merge Engine (`send_email.py` / `send_emails.py`)**: Merges recipient business names into battle-tested cold email templates. Sends through direct SSL SMTP with randomized human delay, checks against `sent_log.csv` to prevent duplicate sends, and automatically syncs all sent emails to the server's `Sent` and `Sent Items` folders for instant visibility in Microsoft Outlook.

---

## 🏗 Pipeline Architecture

```text
┌───────────────────────────────┐
│       Google Maps UI          │  Playwright (Chromium)
│  (Targeted Search Query)      │
└───────────────┬───────────────┘
                │
                ▼
┌───────────────────────────────┐
│     Step 1: main.py           │──► Output: businesses.xlsx / businesses.csv
│     (Lead Generation)         │    (Name, Phone, Website, Address, Rating)
└───────────────┬───────────────┘
                │
                ▼
┌───────────────────────────────┐
│  Step 2: email_extractor.py   │──► Output: businesses_with_emails.xlsx
│  (Multi-Stage Email Hunter)   │    (Static HTML + De-obfuscation + SPA render)
└───────────────┬───────────────┘
                │
                ▼
┌───────────────────────────────┐
│   Step 3: send_email.py       │──► Safe Preview: drafts/*.txt
│   (Mail Merge & Delivery)     │──► Tracking: sent_log.csv
└───────────────┬───────────────┘
                │
        ┌───────┴────────────────────────┐
        ▼                                ▼
┌──────────────────────┐      ┌────────────────────────┐
│  Direct SSL SMTP     │      │   IMAP Dual Folder     │
│  (mail.nurovec.com)  │      │   Sync to 'Sent Items' │
└──────────────────────┘      └────────────────────────┘
```

---

## ⚙️ Prerequisites & Installation

### 1. System Requirements
- Python 3.9+ (Windows recommended for Outlook COM integration)
- Playwright Chromium browser binaries

### 2. Setup Virtual Environment
Open your terminal (PowerShell or Command Prompt) in the project directory:

```bash
# Create virtual environment
python -m venv venv

# Activate on Windows
.\venv\Scripts\activate

# Activate on Linux/macOS
source venv/bin/activate
```

### 3. Install Dependencies
Install the required packages using the bundled requirement files:

```bash
pip install -r requirements.txt
pip install -r requirements_email_extractor.txt
pip install -r requirements_send_email.txt
```

### 4. Install Playwright Browsers
Download the Chromium browser required for scraping and JavaScript rendering:

```bash
playwright install chromium
```

---

## 🚀 Step-by-Step Workflow

### Step 1: Scrape Google Maps Leads

Run the interactive Google Maps scraper:

```bash
python main.py
```

1. Enter your search query when prompted (e.g., `Fintech companies in Gurgaon`).
2. Enter how many businesses you wish to collect (e.g., `50`).
3. A visible Chromium window will launch, automatically scroll through the results feed to load all cards, visit each place profile, and extract details.
4. If Google presents a verification puzzle or CAPTCHA, the console pauses and prompts you to solve it manually in the opened browser window, then press `Enter` to resume.

**Outputs generated:**
- `businesses.xlsx`
- `businesses.csv`

---

### Step 2: Extract Website Contact Emails

Run the email extraction engine to find contact emails from company websites:

```bash
python email_extractor.py
```

The extractor processes `businesses.xlsx` (or any selected file) and runs a 5-tier extraction strategy:
1. **Homepage fetch:** Quick HTTP/HTTPS request with SSL fallback.
2. **De-obfuscation:** Scans `mailto:` links, Cloudflare email protection (`data-cfemail`), and string de-obfuscation (`name [at] domain [dot] com`).
3. **Contact Link Discovery:** Follows internal links like `/contact`, `/about-us`, `/support`, `/reach-us`.
4. **Guessed Paths:** Probes standard endpoint variations if no direct links are exposed.
5. **Headless SPA Rendering:** If static requests fail, launches headless Chromium via Playwright to evaluate JavaScript-rendered web applications.
6. **Filtering:** Automatically strips out placeholder and image extensions (`.png`, `.jpg`, `.svg`, `example.com`, etc.).

**Outputs generated:**
- `businesses_with_emails.xlsx`
- `businesses_with_emails.csv`

---

### Step 3: Send Personalized Emails

Execute personalized cold outreach to the extracted leads.

#### A. Preview Drafts Safely (Dry Run)
Before sending any real emails, verify your merged templates by generating local draft files:

```bash
python send_email.py --dry-run
```
*This writes `.txt` drafts for every valid lead into the `drafts/` directory without contacting anyone.*

#### B. Send Live Outreach
Send live emails using direct authenticated SSL SMTP:

```bash
python send_email.py
```

- Prompts securely for your email password on the first run (or reads from `EMAIL_PASSWORD` environment variable).
- Applies randomized delays (8–15 seconds) between emails to prevent spam detection and mail-server throttling.
- Appends each successful dispatch to `sent_log.csv`.
- Automatically copies the sent email to both the server's `Sent` and `Sent Items` IMAP folders.

---

## ⌨️ Command-Line Options & Flags

| Command | Description |
| :--- | :--- |
| `python send_email.py` | Sends emails to new recipients from `businesses_with_emails.xlsx` |
| `python send_email.py <file.xlsx>` | Sends emails using a specific input file (e.g. `Bpo comapnies.xlsx`) |
| `python send_email.py --dry-run` | Generates text drafts in `drafts/` without sending any real emails |
| `python send_email.py --resend` | Ignores `sent_log.csv` and resends to all valid rows in the spreadsheet |
| `python send_emails.py [args]` | Alias pointing to `send_email.py` with identical arguments |
| `python sync_to_outlook.py` | Syncs historical emails from server `Sent` to Outlook's `Sent Items` folder |

---

## 🛠 Configuration Reference

All settings can be customized directly in [send_email.py](file:///d:/WebScrapper(url_link)/send_email.py):

| Variable | Default Value | Purpose |
| :--- | :--- | :--- |
| `DEFAULT_INPUT_FILE` | `"businesses_with_emails.xlsx"` | Default spreadsheet read on launch |
| `SEND_METHOD` | `"direct"` | `"direct"` (SSL SMTP + IMAP sync) or `"outlook_desktop"` (Windows COM) |
| `SENDER_NAME` | `"Offheman"` | Sender display name |
| `SENDER_EMAIL` | `"ansh.kumar@nurovec.com"` | From address used for outbound messages |
| `MAIL_SERVER` | `"mail.nurovec.com"` | Mail host for SMTP (465) and IMAP (993) |
| `DELAY_BETWEEN_EMAILS` | `(8, 15)` | Min/Max delay in seconds between consecutive emails |
| `TEMPLATE_CHOICE` | `"bulk"` | Active template: `"bulk"`, `"sample"`, or `"ergonomic"` |
| `CATALOGUE_URL` | OneDrive link | Product catalogue link populated into templates |
| `SENT_LOG_FILE` | `"sent_log.csv"` | Local CSV ledger preventing duplicate deliveries |

### Environment Variables
You can avoid entering your password in terminal prompts by exporting it:

```powershell
# PowerShell
$env:EMAIL_PASSWORD="YourSecretPassword"
python send_email.py
```

```bash
# Bash
export EMAIL_PASSWORD="YourSecretPassword"
python send_email.py
```

---

## 🛡 Anti-Spam & Deliverability Architecture

1. **Clean Plain-Text Layout:** Avoids heavy HTML tracking elements, promotional images, and buzzwords that trigger spam scoring filters (e.g., rSPAM, SpamAssassin).
2. **Dynamic Company Personalization:** Injects `{company_name}` into the subject line and body for individualized messaging.
3. **Smart Duplicate Prevention:** Tracks recipient domains and addresses in `sent_log.csv`. If an email was previously contacted, it is skipped unless `--resend` is passed.
4. **File-Lock Resiliency:** If `sent_log.csv` is locked open in Microsoft Excel, `send_email.py` buffers the log to a temporary lock-file and merges it automatically once Excel closes.
5. **Invalid Address Filtering:** Automatically filters out test accounts and invalid formats (such as `sample@mail.com` or syntax errors) before dispatch.
6. **Throttled Human Cadence:** Random sleep intervals between 8 and 15 seconds replicate human sending patterns and protect your domain reputation.

---

## 📬 Outlook & IMAP Sent Sync Tool

When sending through external SMTP, emails often show up only on the webmail server and not inside desktop Outlook folders.

To solve this:
- `send_email.py` uses IMAP `APPEND` to save a copy directly into both `'Sent'` and `'Sent Items'` immediately upon sending.
- If you have previously sent emails that need retroactive syncing to Outlook, run:

```bash
python sync_to_outlook.py
```

**Viewing in Outlook (new):**
1. Switch to your folder list in Outlook.
2. Select **Sent Items** (or **Sent Items_0**).
3. Click **Sync / Update Folder** in the View ribbon to refresh.

---

## ❓ Troubleshooting & FAQ

### Q1: The Google Maps scraper only collects 10 records and stops.
**Solution:** This was caused by Google's obfuscated class marker triggering premature end detection. Update your code or ensure you have the latest `scraper.py`, which uses direct search URLs and card `scroll_into_view` triggers.

### Q2: Google Maps shows a CAPTCHA or "Unusual Traffic".
**Solution:** The scraper is non-headless by default. When a CAPTCHA appears, solve it in the visible Chromium window, then return to your terminal and press `Enter` to resume execution.

### Q3: `Permission denied: 'sent_log.csv'` warning.
**Solution:** This occurs when `sent_log.csv` is open in Microsoft Excel while sending. The script will safely write logs to a temporary file (`sent_log_buffer_*.csv`) and automatically merge them when Excel is closed.

### Q4: Emails bounced with `550 Classified as rSPAM`.
**Solution:** Use the included `"bulk"` template in `send_email.py`. Keep links to a single reputable source (like OneDrive or your corporate domain), avoid aggressive promotional language, and ensure your domain has valid SPF/DKIM DNS records.

---

## 📄 License
Internal proprietary tool for business lead generation and outreach.