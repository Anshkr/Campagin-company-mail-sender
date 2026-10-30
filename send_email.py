"""
send_email.py (and send_emails.py)
----------------------------------
Mail-merge script: takes an Excel file (must have Business Name + Email
columns) and personalizes an email template for every company, replacing
{company_name} with that row's actual name.

SAFE BY DEFAULT or LIVE:
- If SEND_EMAILS = False (or pass --dry-run), it writes personalized .txt drafts
  into the drafts/ folder so you can preview them.
- If SEND_EMAILS = True, it sends the emails using your chosen SEND_METHOD.

TWO SENDING METHODS (set SEND_METHOD below):

  "outlook_desktop" (Recommended for Microsoft 365 / Outlook users on Windows)
      Drives your Outlook desktop application directly via COM.
      Does NOT require SMTP authentication enabled by tenant admin.
      Requirements: Windows + Outlook desktop app + pywin32.

  "smtp"
      Sends via smtplib over SMTP (e.g. smtp.office365.com, smtp.gmail.com).
      Requires mailbox to have Authenticated SMTP or App Password enabled.

Run:
    python send_email.py
    python send_email.py It.xlsx
    python send_email.py --resend         (resends emails even if already in sent_log.csv)
    python send_email.py --dry-run        (generates drafts only, does not send)
"""

import csv
import email.utils
import getpass
import imaplib
import os
import re
import smtplib
import sys
import time
import random
from email.mime.text import MIMEText

import pandas as pd

# ----------------------------------------------------------------------
# CONFIG - edit these before running
# ----------------------------------------------------------------------

# Default input file. If this file does not exist, the script checks
# for businesses_with_emails.xlsx, businesses.xlsx, etc.
DEFAULT_INPUT_FILE = "businesses_with_emails.xlsx"

SENDER_NAME = "Offheman"
YOUR_COMPANY = "Offheman"
YOUR_PHONE = "92174 60742"
YOUR_WEBSITE = "offheman.in"
CATALOGUE_URL = "https://1drv.ms/b/c/9f20d4e9b9dd1dda/IQBYrYx5SWdtSoC6pWyNPR7gAZOz50FQE64AmkwEEb4bJD8?e=jd0dNI"

# Pick which template to use: "bulk" (recommended), "sample", or "ergonomic"
TEMPLATE_CHOICE = "bulk"

# Set to True to send emails, or False to generate drafts only
SEND_EMAILS = True

# Set to True to ignore sent_log.csv and resend emails (can also pass --resend)
IGNORE_SENT_LOG = False

# How to send:
#   "direct" (RECOMMENDED: Direct mail server delivery with IMAP sent sync)
#   "outlook_desktop" (Classic Outlook desktop app)
SEND_METHOD = "direct"

# Sender email address
SENDER_EMAIL = "ansh.kumar@nurovec.com"

# Mail server settings for nurovec.com:
MAIL_SERVER = "mail.nurovec.com"
SMTP_PORT = 465  # SSL port
IMAP_PORT = 993  # SSL port
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", os.environ.get("SMTP_PASSWORD", None))

DELAY_BETWEEN_EMAILS = (8, 15)  # Natural human delay between emails to avoid spam throttling
SENT_LOG_FILE = "sent_log.csv"

# ----------------------------------------------------------------------
# Anti-Spam Optimized Templates (No personal name revealed)
# ----------------------------------------------------------------------

TEMPLATES = {
    "bulk": {
        "subject": "Office chair for {company_name}",
        "body": """Hi team,

I was looking at {company_name} and wanted to reach out briefly.

We work with growing businesses across Gurgaon supplying ergonomic office chairs and workstation seating (through {your_company}).

If you're currently reviewing office seating or planning a workspace refresh in the coming months, I'd be glad to share our catalogue and corporate pricing based on your team size.

Would you be open to a quick 2-minute chat this week?

Best,
Team {your_company}
{your_website}
Ph: {your_phone}

View our full catalogue: {catalogue_url}""",
    },
    "sample": {
        "subject": "{company_name} office seating / Gurgaon",
        "body": """Hi team,

I'm reaching out from {your_company}. We help businesses across Gurgaon furnish their office spaces with durable, comfortable ergonomic chairs.

Rather than sending long pitches, we prefer letting teams test our chairs firsthand. If you are reviewing office furniture or expanding your team at {company_name}, I can arrange a demo chair for your office to test out.

Would you like me to send over our product details or arrange a sample for your team?

Best regards,
Team {your_company}
{your_website}
Ph: {your_phone}

View our full catalogue: {catalogue_url}""",
    },
    "ergonomic": {
        "subject": "ergonomic seating for {company_name}",
        "body": """Hi team,

I noticed {company_name} is based in the Gurgaon area and wanted to connect briefly.

We supply commercial ergonomic chairs designed specifically for teams spending long hours at their desks, helping improve daily comfort and workplace wellness.

If your team is considering new chairs or upgrading your current setup this quarter, let me know and I'd be happy to send over our models and bulk corporate rates.

Regards,
Team {your_company}
{your_website}
Ph: {your_phone}

View our full catalogue: {catalogue_url}""",
    },
}


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------

DUMMY_EMAIL_PATTERNS = {
    "sample@mail.com", "example@example.com", "test@test.com", "user@domain.com",
    "name@domain.com", "email@domain.com", "info@example.com", "admin@domain.com",
    "yourname@email.com", "someone@example.com", "test@example.com",
}

DUMMY_DOMAINS = {
    "example.com", "example.org", "example.net", "domain.com", "mysite.com",
    "yourdomain.com", "sentry.io", "wixpress.com", "wordpress.com",
}


def is_valid_real_email(candidate: str) -> bool:
    if not candidate:
        return False
    cand_lower = candidate.strip().lower()
    if cand_lower in DUMMY_EMAIL_PATTERNS:
        return False
    if "@" not in cand_lower:
        return False
    domain = cand_lower.split("@")[-1]
    if domain in DUMMY_DOMAINS or domain.endswith(".invalid") or domain.endswith(".example"):
        return False
    return True


def extract_valid_email(value) -> str | None:
    """Safely extracts and cleans an email from a cell value. Handles multiple emails, quotes, and whitespace."""
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.upper() == "N/A" or text.lower() == "nan":
        return None

    # Handle multiple emails separated by comma or semicolon
    parts = re.split(r"[,;\s]+", text)
    email_pattern = re.compile(r"^[^@\s]+@[^@\s]+\.[a-zA-Z]{2,}$")
    for part in parts:
        cleaned = part.strip().strip("<>\"'.,;:")
        if email_pattern.match(cleaned) and is_valid_real_email(cleaned):
            return cleaned

    # Fallback regex search
    match = re.search(r"[a-zA-Z0-9.\_%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
    if match:
        candidate = match.group(0).strip()
        if is_valid_real_email(candidate):
            return candidate

    return None


def safe_filename(name: str) -> str:
    return re.sub(r"[^\w\-]+", "_", name).strip("_")[:80] or "company"


SENT_LOG_BUFFER_FILE = "sent_log_buffer.csv"


def merge_sent_log_buffer():
    """Merges any buffered records into the main sent_log.csv if it was previously locked."""
    if not os.path.exists(SENT_LOG_BUFFER_FILE):
        return
    try:
        with open(SENT_LOG_BUFFER_FILE, "r", newline="", encoding="utf-8") as bf:
            buffer_rows = list(csv.reader(bf))
        if not buffer_rows:
            os.remove(SENT_LOG_BUFFER_FILE)
            return

        file_exists = os.path.exists(SENT_LOG_FILE)
        with open(SENT_LOG_FILE, "a", newline="", encoding="utf-8") as mf:
            writer = csv.writer(mf)
            if not file_exists:
                writer.writerow(["Business Name", "Email"])
            for row in buffer_rows:
                if row and len(row) >= 2 and row != ["Business Name", "Email"]:
                    writer.writerow(row)
        os.remove(SENT_LOG_BUFFER_FILE)
    except Exception:
        pass


def load_sent_log() -> set:
    merge_sent_log_buffer()
    sent = set()
    for fname in (SENT_LOG_FILE, SENT_LOG_BUFFER_FILE):
        if os.path.exists(fname):
            try:
                with open(fname, newline="", encoding="utf-8") as f:
                    for row in csv.DictReader(f):
                        if "Email" in row and row["Email"]:
                            sent.add(str(row["Email"]).strip().lower())
            except Exception:
                pass
    return sent


def append_sent_log(business_name: str, email: str):
    # Try merging any old buffer first
    merge_sent_log_buffer()

    # Attempt 1: Write directly to main SENT_LOG_FILE with retry
    for attempt in range(3):
        try:
            file_exists = os.path.exists(SENT_LOG_FILE)
            with open(SENT_LOG_FILE, "a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                if not file_exists:
                    writer.writerow(["Business Name", "Email"])
                writer.writerow([business_name, email])
            return
        except PermissionError:
            time.sleep(0.3)
        except Exception:
            break

    # Attempt 2: If Excel locked sent_log.csv, write to backup buffer so no record is lost
    try:
        buf_exists = os.path.exists(SENT_LOG_BUFFER_FILE)
        with open(SENT_LOG_BUFFER_FILE, "a", newline="", encoding="utf-8") as bf:
            writer = csv.writer(bf)
            if not buf_exists:
                writer.writerow(["Business Name", "Email"])
            writer.writerow([business_name, email])
        print(f"   -> [Saved to buffer: please close '{SENT_LOG_FILE}' in Excel]")
    except Exception as e:
        print(f"   -> WARNING: Could not update sent log: {e}")


def personalize(template: dict, company_name: str) -> tuple:
    subject = template["subject"].format(
        company_name=company_name,
        sender_name=SENDER_NAME,
        your_company=YOUR_COMPANY,
        your_phone=YOUR_PHONE,
        your_website=YOUR_WEBSITE,
        catalogue_url=CATALOGUE_URL,
    )
    body = template["body"].format(
        company_name=company_name,
        sender_name=SENDER_NAME,
        your_company=YOUR_COMPANY,
        your_phone=YOUR_PHONE,
        your_website=YOUR_WEBSITE,
        catalogue_url=CATALOGUE_URL,
    )
    return subject, body


# ----------------------------------------------------------------------
# Sending method 1: Direct Mail Server (Recommended for Outlook (new))
# Sends via mail.nurovec.com and automatically syncs to Sent Items folder!
# ----------------------------------------------------------------------

# Set to True to receive a copy in your Inbox (guaranteed to appear in Outlook (new) instantly!)
BCC_SELF = False

_cached_sent_folders = None


def get_imap_sent_folders(imap) -> list:
    """
    Finds all valid Sent folders on mail.nurovec.com.
    Ensures both 'Sent' (native server folder) and 'Sent Items' (Outlook folder) exist.
    """
    global _cached_sent_folders
    if _cached_sent_folders:
        return _cached_sent_folders

    detected = []
    
    # 1. Check if native 'Sent' folder exists
    try:
        s, _ = imap.select('"Sent"', readonly=True)
        if s == "OK":
            detected.append("Sent")
    except Exception:
        pass

    # 2. Check if 'Sent Items' exists, or try to create it so Outlook (new) can map to it
    try:
        s, _ = imap.select('"Sent Items"', readonly=True)
        if s == "OK":
            detected.append("Sent Items")
        else:
            # Create Sent Items on server if supported
            create_status, _ = imap.create('"Sent Items"')
            if create_status == "OK":
                detected.append("Sent Items")
    except Exception:
        pass

    _cached_sent_folders = detected if detected else ["Sent"]
    return _cached_sent_folders


def save_to_imap_sent(msg_bytes: bytes):
    """Syncs a copy of the sent message directly into server Sent folders."""
    global EMAIL_PASSWORD
    if not EMAIL_PASSWORD:
        return
    try:
        imap = imaplib.IMAP4_SSL(MAIL_SERVER, IMAP_PORT, timeout=15)
        imap.login(SENDER_EMAIL, EMAIL_PASSWORD)
        target_folders = get_imap_sent_folders(imap)

        synced = []
        for folder in target_folders:
            folder_clean = folder.strip('"')
            status, resp = imap.append(f'"{folder_clean}"', "\\Seen", imaplib.Time2Internaldate(time.time()), msg_bytes)
            if status == "OK":
                synced.append(folder_clean)

        if synced:
            print(f"   -> Saved copy to server folder(s): {synced}")
        imap.logout()
    except Exception as e:
        print(f"   -> Note: IMAP sync warning: {e}")


def sync_drafts_to_imap():
    """Uploads saved drafts from drafts/ directly into Outlook (new) Sent Items without sending live emails."""
    global EMAIL_PASSWORD
    if not EMAIL_PASSWORD:
        EMAIL_PASSWORD = getpass.getpass(f"Password for {SENDER_EMAIL}: ")

    draft_dir = "drafts"
    if not os.path.exists(draft_dir):
        print(f"ERROR: Directory '{draft_dir}' does not exist.")
        return

    draft_files = [f for f in os.listdir(draft_dir) if f.endswith(".txt")]
    if not draft_files:
        print(f"No .txt draft files found in '{draft_dir}'.")
        return

    print(f"\nConnecting to {MAIL_SERVER} IMAP as {SENDER_EMAIL}...")
    try:
        imap = imaplib.IMAP4_SSL(MAIL_SERVER, IMAP_PORT, timeout=15)
        imap.login(SENDER_EMAIL, EMAIL_PASSWORD)
        target_folders = get_imap_sent_folders(imap)
        print(f"Found server Sent folders: {target_folders}")

        count = 0
        for df_name in draft_files:
            df_path = os.path.join(draft_dir, df_name)
            try:
                with open(df_path, "r", encoding="utf-8") as f:
                    content = f.read()

                lines = content.splitlines()
                to_addr = ""
                subj = ""
                body_start = 0
                for i, line in enumerate(lines):
                    if line.startswith("To: "):
                        to_addr = line[4:].strip()
                    elif line.startswith("Subject: "):
                        subj = line[9:].strip()
                    elif line == "":
                        body_start = i + 1
                        break
                body_text = "\n".join(lines[body_start:])

                msg = MIMEText(body_text, "plain", "utf-8")
                msg["Subject"] = subj
                msg["From"] = SENDER_EMAIL
                msg["To"] = to_addr
                msg["Reply-To"] = SENDER_EMAIL
                msg["Date"] = email.utils.formatdate(localtime=True)

                msg_bytes = msg.as_bytes()
                for folder in target_folders:
                    folder_clean = folder.strip('"')
                    imap.append(f'"{folder_clean}"', "\\Seen", imaplib.Time2Internaldate(time.time()), msg_bytes)
                print(f"   -> Synced '{df_name}' ({to_addr}) to Outlook Sent Items")
                count += 1
            except Exception as e:
                print(f"   -> Could not sync '{df_name}': {e}")

        imap.logout()
        print("\n" + "=" * 60)
        print(f"Done! Successfully synced {count} drafts into Outlook Sent Items.")
        print("In Outlook (new), click the 'Sync' button in the 'View' tab to refresh!")
        print("=" * 60)
    except Exception as e:
        print(f"ERROR syncing to IMAP: {e}")



def send_email_direct(to_email: str, subject: str, body: str):
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = email.utils.formataddr((SENDER_NAME, SENDER_EMAIL))
    msg["To"] = to_email
    msg["Reply-To"] = email.utils.formataddr((SENDER_NAME, SENDER_EMAIL))
    msg["Date"] = email.utils.formatdate(localtime=True)
    msg["Message-ID"] = email.utils.make_msgid(domain="nurovec.com")

    recipients = [to_email]
    if BCC_SELF and SENDER_EMAIL.lower() != to_email.lower():
        recipients.append(SENDER_EMAIL)

    with smtplib.SMTP_SSL(MAIL_SERVER, SMTP_PORT, timeout=30) as server:
        server.login(SENDER_EMAIL, EMAIL_PASSWORD)
        server.send_message(msg, to_addrs=recipients)

    # Automatically sync to server Sent folders:
    save_to_imap_sent(msg.as_bytes())


# ----------------------------------------------------------------------
# Sending method 2: Outlook desktop automation (Windows only)
# ----------------------------------------------------------------------

_outlook_app = None
_outlook_ns = None
_sender_account = None
_sender_account_checked = False


def get_outlook_session():
    """
    Connects to Outlook Application and initializes MAPI Namespace.
    Attaches to running Outlook or starts one, and logs onto default profile.
    """
    global _outlook_app, _outlook_ns
    if _outlook_app is not None and _outlook_ns is not None:
        return _outlook_app, _outlook_ns

    try:
        import win32com.client
    except ImportError:
        raise RuntimeError("pywin32 is not installed. Install it with: pip install pywin32")

    # Try attaching to already open Outlook first, fallback to Dispatch
    try:
        _outlook_app = win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        try:
            _outlook_app = win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            raise RuntimeError(
                f"Failed to start Outlook COM object ({e}). "
                "Make sure Outlook desktop app is installed and open, or switch SEND_METHOD to 'smtp'."
            )

    try:
        _outlook_ns = _outlook_app.GetNamespace("MAPI")
        _outlook_ns.Logon("", "", False, False)
    except Exception as e:
        # Logon error is often non-fatal if session is already active
        pass

    return _outlook_app, _outlook_ns


def find_sender_account(outlook_ns):
    """
    Look up the Outlook account object matching SENDER_EMAIL.
    Strictly enforces that SENDER_EMAIL must exist in Outlook.
    Never falls back to personal or default accounts.
    """
    global _sender_account, _sender_account_checked
    if _sender_account_checked:
        return _sender_account
    _sender_account_checked = True

    try:
        accounts = outlook_ns.Accounts
    except Exception as e:
        raise RuntimeError(f"Could not read Outlook accounts: {e}")

    account_list = []
    for i in range(1, accounts.Count + 1):
        acc = accounts.Item(i)
        addr = getattr(acc, "SmtpAddress", None) or getattr(acc, "DisplayName", "")
        account_list.append((acc, str(addr).strip().lower()))

    target = SENDER_EMAIL.strip().lower()
    for acc, addr in account_list:
        if addr == target:
            _sender_account = acc
            print(f"   -> Verified Outlook Account: {acc.DisplayName} ({addr})")
            return _sender_account

    # STRICT: Do not use default or personal accounts!
    detected = [a[1] for a in account_list]
    raise RuntimeError(
        f"\n"
        f"======================================================================\n"
        f"ERROR: '{SENDER_EMAIL}' is NOT configured in Outlook!\n"
        f"Detected accounts in your Outlook app: {detected}\n"
        f"\n"
        f"To prevent sending from your personal account ({detected[0] if detected else 'default'}),\n"
        f"the script has stopped.\n"
        f"\n"
        f"Please add '{SENDER_EMAIL}' into your Outlook Desktop app:\n"
        f"  1. Open Outlook -> Click 'File' (top-left) -> 'Add Account'\n"
        f"  2. Enter: {SENDER_EMAIL}\n"
        f"  3. Follow the setup prompts with your password\n"
        f"======================================================================"
    )


def send_email_outlook_desktop(to_email: str, subject: str, body: str):
    app, ns = get_outlook_session()
    account = find_sender_account(ns)  # Strictly verifies SENDER_EMAIL exists

    mail = app.CreateItem(0)  # 0 = olMailItem
    mail.To = to_email
    mail.Subject = subject
    mail.Body = body
    mail.SendUsingAccount = account

    try:
        mail.SentOnBehalfOfName = SENDER_EMAIL
        mail.ReplyRecipients.Add(SENDER_EMAIL)
    except Exception:
        pass

    mail.Send()
    try:
        ns.SendAndReceive(False)
    except Exception:
        pass


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------

def resolve_input_file(arg_file: str | None) -> str:
    """Finds the most suitable input file."""
    if arg_file and os.path.exists(arg_file):
        return arg_file

    candidates = [DEFAULT_INPUT_FILE, "businesses_with_emails.xlsx", "businesses.xlsx", "businesses.csv"]
    for c in candidates:
        if os.path.exists(c):
            return c

    return DEFAULT_INPUT_FILE


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Finds a matching column name in a case-insensitive manner."""
    lower_map = {str(col).strip().lower(): col for col in df.columns}
    for candidate in candidates:
        if candidate.lower() in lower_map:
            return lower_map[candidate.lower()]
    return None


def main():
    global EMAIL_PASSWORD, SEND_EMAILS, IGNORE_SENT_LOG, SEND_METHOD

    # Process CLI flags
    args = sys.argv[1:]
    input_file_arg = None
    for arg in args:
        if arg.lower() in ("--sync-sent", "--sync-drafts", "--sync"):
            sync_drafts_to_imap()
            return
        elif arg.lower() in ("--dry-run", "-d"):
            SEND_EMAILS = False
        elif arg.lower() in ("--resend", "--force", "-f"):
            IGNORE_SENT_LOG = True
        elif arg.lower() in ("--direct", "--smtp"):
            SEND_METHOD = "direct"
        elif arg.lower() in ("--outlook", "--desktop"):
            SEND_METHOD = "outlook_desktop"
        elif not arg.startswith("-"):
            input_file_arg = arg

    input_file = resolve_input_file(input_file_arg)

    print("=" * 60)
    print(" Mail Merge - Personalize & Send Emails")
    print(f" Sender email: {SENDER_EMAIL}")
    print(f" Input file:   {input_file}")
    print(f" Send method:  {'Direct via mail.nurovec.com (syncs to Outlook (new) Sent Items)' if SEND_METHOD in ('direct', 'smtp') else 'Outlook Desktop (Classic)'}")
    print(f" Mode:         {'SENDING LIVE' if SEND_EMAILS else 'DRY RUN (drafts only, nothing will be sent)'}")
    if IGNORE_SENT_LOG:
        print(" Sent log:     IGNORED (--resend mode: will send to previously sent emails)")
    print("=" * 60)

    if SEND_METHOD not in ("direct", "smtp", "outlook_desktop"):
        print(f"ERROR: Unknown SEND_METHOD '{SEND_METHOD}'. Use 'direct' or 'outlook_desktop'.")
        return

    # Verify Outlook connection early if using outlook_desktop
    if SEND_METHOD == "outlook_desktop":
        try:
            print("Connecting to Outlook Desktop...")
            app, ns = get_outlook_session()
            account = find_sender_account(ns)
            print("Outlook connection ready.")
        except Exception as e:
            print(f"\nERROR connecting to Outlook: {e}")
            return

    # Verify credentials early if using direct mail server
    if SEND_EMAILS and SEND_METHOD in ("direct", "smtp"):
        if not EMAIL_PASSWORD:
            EMAIL_PASSWORD = getpass.getpass(f"Password for {SENDER_EMAIL}: ")
        print(f"Connecting to {MAIL_SERVER} as {SENDER_EMAIL}...")
        try:
            test_server = smtplib.SMTP_SSL(MAIL_SERVER, SMTP_PORT, timeout=15)
            test_server.login(SENDER_EMAIL, EMAIL_PASSWORD)
            test_server.quit()
            print("Credentials verified! Ready to send.")
        except Exception as e:
            print(f"\nERROR: Login failed for {SENDER_EMAIL}: {e}")
            print("Please verify your password and try again.")
            return

    # Load data
    try:
        if input_file.endswith(".csv"):
            df = pd.read_csv(input_file)
        else:
            df = pd.read_excel(input_file)
    except Exception as e:
        print(f"ERROR: Could not read '{input_file}': {e}")
        return

    # Detect required columns
    name_col = find_column(df, ["Business Name", "Company Name", "Company", "Name", "Business"])
    email_col = find_column(df, ["Email", "Email Address", "E-mail", "Mail", "Emails"])

    if not name_col or not email_col:
        print("ERROR: File must have business name and email columns.")
        print(f"Found columns: {list(df.columns)}")
        print("Expected columns such as 'Business Name' and 'Email'.")
        return

    template = TEMPLATES.get(TEMPLATE_CHOICE)
    if template is None:
        print(f"ERROR: Unknown TEMPLATE_CHOICE '{TEMPLATE_CHOICE}'. Use bulk / sample / ergonomic.")
        return

    already_sent = load_sent_log() if (SEND_EMAILS and not IGNORE_SENT_LOG) else set()
    os.makedirs("drafts", exist_ok=True)

    total = len(df)
    valid_emails_in_file = 0
    all_matched_already_sent = 0

    # Quick pre-scan to check if all emails were previously sent
    for _, row in df.iterrows():
        raw_email = row.get(email_col)
        cleaned_email = extract_valid_email(raw_email)
        if cleaned_email:
            valid_emails_in_file += 1
            if cleaned_email.lower() in already_sent:
                all_matched_already_sent += 1

    if SEND_EMAILS and valid_emails_in_file > 0 and all_matched_already_sent == valid_emails_in_file and not IGNORE_SENT_LOG:
        print("\n" + "!" * 60)
        print(f"NOTICE: All {valid_emails_in_file} valid emails in '{input_file}' are already recorded as sent in '{SENT_LOG_FILE}'.")
        print("To prevent accidental duplicate emails, no emails were sent.")
        print("")
        print("To resend to these businesses:")
        print("  - Run with --resend:   python send_email.py --resend")
        print("  - Or delete/rename:   sent_log.csv")
        print("  - Or set IGNORE_SENT_LOG = True in the script config")
        print("!" * 60)
        return

    skipped_no_email = 0
    skipped_already_sent = 0
    processed = 0

    print(f"\nProcessing {total} rows from '{input_file}'...\n")

    for idx, row in df.iterrows():
        company_name = str(row.get(name_col, "")).strip() or f"Company {idx + 1}"
        raw_email = row.get(email_col)
        email = extract_valid_email(raw_email)

        if not email:
            skipped_no_email += 1
            continue

        if SEND_EMAILS and not IGNORE_SENT_LOG and email.lower() in already_sent:
            skipped_already_sent += 1
            continue

        subject, body = personalize(template, company_name)

        # Always save a draft copy for your records
        draft_path = os.path.join("drafts", f"{safe_filename(company_name)}.txt")
        try:
            with open(draft_path, "w", encoding="utf-8") as f:
                f.write(f"To: {email}\nSubject: {subject}\n\n{body}")
        except Exception:
            pass

        if SEND_EMAILS:
            print(f"[{idx + 1}/{total}] Sending to {company_name} <{email}> ...")
            try:
                if SEND_METHOD in ("direct", "smtp"):
                    send_email_direct(email, subject, body)
                else:
                    send_email_outlook_desktop(email, subject, body)
                append_sent_log(company_name, email)
                print("   -> Sent successfully")
            except Exception as e:
                print(f"   -> FAILED: {e}")
            time.sleep(random.uniform(*DELAY_BETWEEN_EMAILS))
        else:
            print(f"[{idx + 1}/{total}] Draft saved for {company_name} <{email}>")

        processed += 1

    if SEND_EMAILS and SEND_METHOD == "outlook_desktop" and processed > 0:
        print("\nTransmitting queued emails via Outlook Desktop...")
        try:
            _, ns = get_outlook_session()
            ns.SendAndReceive(False)
            for i in range(1, ns.SyncObjects.Count + 1):
                try:
                    ns.SyncObjects.Item(i).Start()
                except Exception:
                    pass
            # Wait up to 5 seconds for Outlook background spooler to empty the Outbox
            outbox = ns.GetDefaultFolder(4)
            for _ in range(5):
                if outbox.Items.Count == 0:
                    break
                time.sleep(1)
            print("   -> Outbox transmission cycle completed.")
        except Exception:
            pass

    print("\n" + "=" * 60)
    print(f"Done. Processed: {processed}, Skipped (no/invalid email): {skipped_no_email}, "
          f"Skipped (already sent): {skipped_already_sent}")
    if not SEND_EMAILS:
        print("This was a DRY RUN. Check the 'drafts' folder, then run 'python send_email.py' to send.")
    print("=" * 60)


if __name__ == "__main__":
    main()
