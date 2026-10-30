"""
send_emails.py
--------------
Compatibility entrypoint forwarding directly to send_email.py.
Both `python send_email.py` and `python send_emails.py` are supported.
"""

import send_email

if __name__ == "__main__":
    send_email.main()