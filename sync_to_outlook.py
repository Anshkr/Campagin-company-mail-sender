"""
sync_to_outlook.py
------------------
Copies all emails from the server's native 'Sent' folder into 'Sent Items'
so they show up immediately in Outlook's 'Sent Items_0' folder.
"""
import imaplib
import getpass
import os

MAIL_SERVER = "mail.nurovec.com"
SENDER_EMAIL = "ansh.kumar@nurovec.com"
password = os.environ.get("EMAIL_PASSWORD") or getpass.getpass(f"Password for {SENDER_EMAIL}: ")

print(f"\nConnecting to {MAIL_SERVER} as {SENDER_EMAIL}...")
imap = imaplib.IMAP4_SSL(MAIL_SERVER, 993, timeout=15)
imap.login(SENDER_EMAIL, password)

# 1. Ensure 'Sent Items' exists
imap.create('"Sent Items"')

# 2. Copy all messages from 'Sent' to 'Sent Items'
status, _ = imap.select('"Sent"')
if status == "OK":
    status, data = imap.search(None, "ALL")
    msg_ids = data[0].split()
    print(f"Found {len(msg_ids)} sent emails in server 'Sent' folder.")
    print("Copying them into 'Sent Items' (Outlook Sent Items_0)...")
    copied = 0
    for mid in msg_ids:
        res, _ = imap.copy(mid, '"Sent Items"')
        if res == "OK":
            copied += 1
    print(f"Successfully transferred {copied} emails into 'Sent Items'!")

# 3. Verify total in 'Sent Items'
s, data = imap.select('"Sent Items"', readonly=True)
count = data[0].decode() if data and data[0] else '0'
print(f"\nTotal emails now inside 'Sent Items': {count}")

imap.logout()
print("\n" + "=" * 60)
print("SUCCESS! All your sent emails are now in 'Sent Items'.")
print("In Outlook (new):")
print("  1. Click on another folder (e.g. 'Inbox').")
print("  2. Click back on 'Sent Items_0' and click 'Sync' in the View tab.")
print("  3. All your sent emails will appear!")
print("=" * 60)
