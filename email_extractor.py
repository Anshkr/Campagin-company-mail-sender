"""
email_extractor.py
-------------------
Reads an existing Excel file (columns: Business Name, Phone Number, Website,
Address, Rating, Number of Reviews, Google Maps URL) and, for every row that
has a real Website, tries to scrape a contact email address from that site.

Rows whose Website is "N/A" / empty are skipped and get Email = "N/A".
If nothing is found after all strategies below, Email is also "N/A".

Strategies used, in order, per website:
  1. Fetch the homepage (with https/http + SSL-retry fallback).
  2. Look for emails via: mailto: links, Cloudflare-obfuscated emails
     (data-cfemail / /cdn-cgi/l/email-protection#...), plain text regex,
     and "name [at] domain [dot] com"-style de-obfuscation.
  3. If nothing found, discover real contact/about links from the
     homepage's own navigation (not just guessed URLs) and check those.
  4. If still nothing, guess a handful of common contact-page paths.
  5. If STILL nothing (likely a JS-rendered site), fall back to rendering
     the homepage with a real (headless) browser via Playwright and
     re-scan the rendered HTML. This step is slower, so it only runs
     when the fast methods above fail.

Run:
    python email_extractor.py

Output:
    businesses_with_emails.csv
    businesses_with_emails.xlsx
"""

import random
import re
import time
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
import urllib3
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ----------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------

REQUEST_TIMEOUT = 12
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# Guessed contact-style pages, tried if link-discovery on the homepage
# doesn't turn up anything.
GUESSED_PATHS = [
    "contact", "contact-us", "contactus", "contact_us", "contact.html",
    "contact-us.html", "about", "about-us", "support", "get-in-touch",
    "reach-us", "connect", "enquiry", "inquiry",
]

MAX_STATIC_PAGES_PER_SITE = 6  # cap requests per website for the fast path

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9.\_%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")

LINK_KEYWORDS = [
    "contact", "about", "support", "enquiry", "inquiry",
    "connect", "reach", "touch", "help",
]

BLOCKED_DOMAINS = {
    "example.com", "domain.com", "yourdomain.com", "email.com",
    "sentry.io", "wixpress.com", "godaddy.com", "schema.org",
    "w3.org", "gstatic.com", "googleapis.com", "cloudflare.com",
}

BLOCKED_TLDS = {
    "png", "jpg", "jpeg", "gif", "svg", "webp", "css", "js",
    "woff", "woff2", "ttf", "eot", "ico", "bmp",
}


# ----------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------

def is_missing(value) -> bool:
    if value is None:
        return True
    text = str(value).strip()
    return text == "" or text.upper() == "N/A" or text.lower() == "nan"


def normalize_url(url: str) -> str:
    url = url.strip()
    if not url.lower().startswith(("http://", "https://")):
        url = "http://" + url
    return url


def site_domain_of(url: str) -> str:
    return urlparse(url).netloc.lower().replace("www.", "")


def is_valid_email(email: str, site_domain: str) -> bool:
    email = email.lower().strip().strip(".")
    if "@" not in email:
        return False
    local, _, domain = email.partition("@")
    tld = domain.rsplit(".", 1)[-1] if "." in domain else ""

    if domain in BLOCKED_DOMAINS or tld in BLOCKED_TLDS:
        return False
    if len(local) < 2 or len(domain) < 4 or len(local) > 64:
        return False
    if ".." in email or local.startswith((".", "-")) or local.endswith((".", "-")):
        return False
    if local in {"sentry", "wixpress", "no-reply", "noreply"} and domain != site_domain:
        return False
    return True


def decode_cf_email(cfemail: str):
    """Decode Cloudflare's email-obfuscation hex encoding."""
    try:
        r = int(cfemail[:2], 16)
        return "".join(
            chr(int(cfemail[i:i + 2], 16) ^ r) for i in range(2, len(cfemail), 2)
        )
    except Exception:
        return None


def deobfuscate_text(text: str) -> str:
    """Turn 'name [at] domain [dot] com' style text into 'name@domain.com'."""
    t = text
    for pattern, repl in [
        (r"\s*[\[\(\{]\s*at\s*[\]\)\}]\s*", "@"),
        (r"\s+at\s+", "@"),
        (r"\s*[\[\(\{]\s*dot\s*[\]\)\}]\s*", "."),
        (r"\s+dot\s+", "."),
    ]:
        t = re.sub(pattern, repl, t, flags=re.IGNORECASE)
    return t


# ----------------------------------------------------------------------
# Fetching (with https/http + SSL-error fallback)
# ----------------------------------------------------------------------

def _get(url: str, session: requests.Session):
    try:
        resp = session.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, allow_redirects=True)
    except requests.exceptions.SSLError:
        try:
            resp = session.get(
                url, headers=HEADERS, timeout=REQUEST_TIMEOUT,
                allow_redirects=True, verify=False,
            )
        except requests.RequestException:
            return None
    except requests.RequestException:
        return None

    if resp.status_code >= 400:
        return None
    return resp


def fetch(url: str, session: requests.Session):
    resp = _get(url, session)
    if resp is not None:
        return resp
    # Try the other scheme (http<->https) before giving up
    parsed = urlparse(url)
    alt_scheme = "https" if parsed.scheme == "http" else "http"
    alt_url = url.replace(parsed.scheme + "://", alt_scheme + "://", 1)
    return _get(alt_url, session)


# ----------------------------------------------------------------------
# Email extraction from a page's HTML
# ----------------------------------------------------------------------

def extract_emails_from_html(html: str, site_domain: str) -> list:
    found = set()
    soup = BeautifulSoup(html, "html.parser")

    # 1) mailto: links
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.lower().startswith("mailto:"):
            addr = href[7:].split("?")[0].strip()
            if addr:
                found.add(addr)

    # 2) Cloudflare-obfuscated emails (data-cfemail attribute anywhere)
    for tag in soup.find_all(attrs={"data-cfemail": True}):
        decoded = decode_cf_email(tag["data-cfemail"])
        if decoded:
            found.add(decoded)

    # 2b) Cloudflare-obfuscated emails via /cdn-cgi/l/email-protection#hash
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "cdn-cgi/l/email-protection#" in href:
            token = href.split("#")[-1]
            decoded = decode_cf_email(token)
            if decoded:
                found.add(decoded)

    # 3) Visible text - direct regex
    text = soup.get_text(" ")
    for m in EMAIL_REGEX.findall(text):
        found.add(m)

    # 4) Visible text - de-obfuscated ("at"/"dot" spelled out)
    for m in EMAIL_REGEX.findall(deobfuscate_text(text)):
        found.add(m)

    # 5) Raw HTML source (catches emails in data-* attributes, JSON-LD, etc.)
    for m in EMAIL_REGEX.findall(html):
        found.add(m)

    valid = [e for e in found if is_valid_email(e, site_domain)]
    # Prefer an email on the business's own domain over third-party ones
    valid.sort(key=lambda e: 0 if e.lower().endswith("@" + site_domain) else 1)
    return valid


def discover_contact_links(html: str, base_url: str, site_domain: str) -> list:
    """Find real contact/about-style links from the homepage's own nav,
    instead of only guessing common paths."""
    soup = BeautifulSoup(html, "html.parser")
    found, seen = [], set()

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href or href.startswith("#") or href.lower().startswith(("mailto:", "tel:", "javascript:")):
            continue
        link_text = a.get_text(" ").strip().lower()
        haystack = href.lower() + " " + link_text

        if any(k in haystack for k in LINK_KEYWORDS):
            abs_url = urljoin(base_url, href)
            parsed = urlparse(abs_url)
            if parsed.scheme not in ("http", "https"):
                continue
            link_domain = parsed.netloc.lower().replace("www.", "")
            if link_domain and link_domain != site_domain:
                continue  # stay on the same site
            if abs_url not in seen:
                seen.add(abs_url)
                found.append(abs_url)

    return found[:6]


# ----------------------------------------------------------------------
# Per-website scraping
# ----------------------------------------------------------------------

def scrape_email_static(website: str, session: requests.Session):
    """Fast path: plain HTTP requests, no JS rendering."""
    homepage_url = normalize_url(website)
    domain = site_domain_of(homepage_url)
    if not homepage_url.endswith("/"):
        homepage_url += "/"

    pages_checked = 0

    resp = fetch(homepage_url, session)
    pages_checked += 1
    discovered_links = []

    if resp is not None:
        emails = extract_emails_from_html(resp.text, domain)
        if emails:
            return emails[0]
        discovered_links = discover_contact_links(resp.text, homepage_url, domain)

    candidates = discovered_links + [urljoin(homepage_url, p) for p in GUESSED_PATHS]
    seen = {homepage_url}
    for url in candidates:
        if pages_checked >= MAX_STATIC_PAGES_PER_SITE:
            break
        if url in seen:
            continue
        seen.add(url)

        resp = fetch(url, session)
        pages_checked += 1
        if resp is None:
            continue

        emails = extract_emails_from_html(resp.text, domain)
        if emails:
            return emails[0]

        time.sleep(random.uniform(0.4, 0.9))

    return "N/A"


def scrape_email_js_fallback(website: str, browser_context, discovered_url: str = None):
    """Slow path: render with a real headless browser, for JS-heavy sites
    where the email never appears in the raw HTML."""
    homepage_url = normalize_url(website)
    domain = site_domain_of(homepage_url)

    urls_to_try = [homepage_url]
    if discovered_url:
        urls_to_try.append(discovered_url)

    for url in urls_to_try:
        try:
            page = browser_context.new_page()
            page.goto(url, timeout=20000, wait_until="networkidle")
            html = page.content()
            page.close()
        except Exception:
            continue

        emails = extract_emails_from_html(html, domain)
        if emails:
            return emails[0]

    return "N/A"


def scrape_email_for_website(website: str, session: requests.Session, browser_context) -> str:
    email = scrape_email_static(website, session)
    if email != "N/A":
        return email

    # Fast path failed - likely a JS-rendered site. Try rendering it.
    return scrape_email_js_fallback(website, browser_context)


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------

def main():
    print("=" * 60)
    print(" Email Extractor - scrapes contact emails from Website column")
    print("=" * 60)

    input_path = input(
        "Enter path to your Excel file (default: businesses.xlsx): "
    ).strip() or "businesses.xlsx"

    try:
        df = pd.read_excel(input_path)
    except Exception as e:
        print(f"Could not read '{input_path}': {e}")
        return

    if "Website" not in df.columns:
        print("This file has no 'Website' column - cannot continue.")
        return

    emails = []
    session = requests.Session()
    total = len(df)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        browser_context = browser.new_context(
            user_agent=HEADERS["User-Agent"],
            viewport={"width": 1366, "height": 900},
        )

        for idx, row in df.iterrows():
            website = row.get("Website")
            name = row.get("Business Name", f"Row {idx + 1}")

            if is_missing(website):
                print(f"[{idx + 1}/{total}] {name}: no website -> Email = N/A")
                emails.append("N/A")
                continue

            print(f"[{idx + 1}/{total}] {name}: scraping {website} ...")
            try:
                email = scrape_email_for_website(str(website), session, browser_context)
            except Exception as e:
                print(f"   -> error: {e}")
                email = "N/A"

            print(f"   -> {email}")
            emails.append(email)

            time.sleep(random.uniform(1.0, 2.0))  # delay between different websites

        browser.close()

    df["Email"] = emails

    df.to_csv("businesses_with_emails.csv", index=False, encoding="utf-8-sig")
    df.to_excel("businesses_with_emails.xlsx", index=False)

    found_count = sum(1 for e in emails if e != "N/A")
    print("\n" + "=" * 60)
    print(f"Done. Found emails for {found_count}/{total} businesses.")
    print("Saved to:")
    print("  - businesses_with_emails.csv")
    print("  - businesses_with_emails.xlsx")
    print("=" * 60)


if __name__ == "__main__":
    main()