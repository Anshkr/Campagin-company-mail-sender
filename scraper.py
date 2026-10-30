"""
scraper.py
----------
Core Google Maps scraping logic using Playwright.

Design notes:
- We drive real Google Maps (maps.google.com) in a *visible* (non-headless)
  Chromium browser. This is genuine browser automation against the real
  Google Maps UI - nothing is faked or mocked.
- Google's DOM/class names change often, so every field extraction tries
  several fallback selectors before giving up and returning "N/A".
- If Google shows a CAPTCHA / "unusual traffic" page, the script pauses
  and waits for YOU to solve it manually in the visible browser window,
  then press Enter in the terminal to continue. It never tries to bypass it.
"""

import random
import time
import urllib.parse

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError


# ----------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------

def _sleep(a=1.0, b=2.2):
    """Random human-like delay."""
    time.sleep(random.uniform(a, b))


def _first_text(page, selectors):
    """Try a list of CSS selectors, return the first non-empty inner_text."""
    for sel in selectors:
        try:
            el = page.locator(sel).first
            if el.count() > 0:
                txt = el.inner_text(timeout=1500).strip()
                if txt:
                    return txt
        except Exception:
            continue
    return "N/A"


def _first_attr(page, selectors, attr):
    """Try a list of CSS selectors, return the first non-empty attribute value."""
    for sel in selectors:
        try:
            el = page.locator(sel).first
            if el.count() > 0:
                val = el.get_attribute(attr, timeout=1500)
                if val:
                    return val.strip()
        except Exception:
            continue
    return "N/A"


class GoogleMapsScraper:
    def __init__(self, headless: bool = False, slow_mo: int = 50):
        self.headless = headless
        self.slow_mo = slow_mo

    # ------------------------------------------------------------------
    # CAPTCHA handling
    # ------------------------------------------------------------------
    def _check_and_wait_for_captcha(self, page):
        """
        If Google shows a CAPTCHA / 'unusual traffic' interstitial, pause
        execution and let the human solve it in the visible browser window.
        """
        try:
            content = page.content().lower()
        except Exception:
            return

        captcha_markers = [
            "unusual traffic",
            "recaptcha",
            "/sorry/",
            "our systems have detected",
            "verify you're a human",
            "verify you are a human",
        ]

        if any(marker in content for marker in captcha_markers) or "/sorry/" in page.url:
            print("\n" + "=" * 60)
            print("CAPTCHA / verification detected!")
            print("Please solve it manually in the opened Chromium window.")
            input("Once solved, press ENTER here to continue...")
            print("=" * 60 + "\n")
            _sleep(1, 2)

    # ------------------------------------------------------------------
    # Consent handling (EU/EEA cookie consent screen)
    # ------------------------------------------------------------------
    def _handle_consent(self, page):
        consent_texts = ["Accept all", "I agree", "Reject all", "Accept"]
        for text in consent_texts:
            try:
                btn = page.get_by_role("button", name=text, exact=False)
                if btn.count() > 0:
                    btn.first.click(timeout=2000)
                    _sleep(0.5, 1.2)
                    return
            except Exception:
                continue

    # ------------------------------------------------------------------
    # Step 1: search Google Maps and collect result URLs
    # ------------------------------------------------------------------
    def _collect_listing_urls(self, page, query: str, target_count: int):
        print(f"Opening Google Maps for query: '{query}' ...")
        # Direct navigation is faster and more reliable than typing in the search box
        search_url = f"https://www.google.com/maps/search/{urllib.parse.quote_plus(query)}"
        page.goto(search_url, timeout=60000)
        _sleep(3.0, 5.0)

        self._handle_consent(page)
        self._check_and_wait_for_captcha(page)

        # Wait for either the results feed or a single place page
        try:
            page.wait_for_selector("div[role='feed'], div.m6QErb, a[href*='/maps/place/']", timeout=12000)
        except Exception:
            pass

        # The scrollable results panel (feed of business cards)
        feed_selectors = [
            "div[role='feed']",
            "div.m6QErb[aria-label]",
            "div[aria-label*='Results for']",
            "div.m6QErb",
        ]
        feed = None
        for sel in feed_selectors:
            try:
                loc = page.locator(sel).first
                if loc.count() > 0:
                    feed = loc
                    break
            except Exception:
                continue

        # Single-result case: Maps sometimes jumps straight to one business
        if feed is None:
            current_url = page.url
            if "/maps/place/" in current_url:
                print("Single business result detected.")
                return [current_url]
            raise RuntimeError(
                "Could not find the results panel. Google's layout may have "
                "changed, or no results were found for this query."
            )

        # Scroll the feed to load more listings until we reach target_count or end of list
        collected = []
        seen = set()
        stagnant_rounds = 0
        max_stagnant_rounds = 6

        print(f"Scrolling to load at least {target_count} results...")

        while len(collected) < target_count and stagnant_rounds < max_stagnant_rounds:
            # Extract links across multiple possible Google Maps card selectors
            anchors = feed.locator("a[href*='/maps/place/'], a.hfpxzc")
            n = anchors.count()
            for i in range(n):
                try:
                    href = anchors.nth(i).get_attribute("href")
                except Exception:
                    href = None
                if href:
                    clean_url = href.split("?")[0] if "?" in href else href
                    if clean_url not in seen:
                        seen.add(clean_url)
                        seen.add(href)
                        collected.append(href)

            before = len(collected)

            # Scroll the last card into view - this triggers Google Maps to fetch the next batch
            try:
                cards = feed.locator("div.Nv2PK, div[role='article']")
                card_count = cards.count()
                if card_count > 0:
                    cards.last.scroll_into_view_if_needed(timeout=2000)
            except Exception:
                pass

            # Scroll using DOM scrollTop and mouse wheel inside feed
            try:
                feed.evaluate("(el) => { el.scrollTop = el.scrollHeight; el.scrollBy(0, 4000); }")
            except Exception:
                pass

            try:
                page.mouse.wheel(0, 4000)
            except Exception:
                pass

            # Sleep enough time for Google Maps to trigger AJAX and render the new batch
            _sleep(2.5, 3.5)
            self._check_and_wait_for_captcha(page)

            # Check if we've reached the explicit end of results
            # Note: Do not use generic class selectors like span.HlvSq as they match non-terminal elements!
            end_markers = [
                "text=You've reached the end of the list",
                "text=You've reached the end",
                "text=No more results",
            ]
            reached_end = False
            for marker in end_markers:
                try:
                    if page.locator(marker).first.count() > 0:
                        reached_end = True
                        break
                except Exception:
                    pass

            if reached_end:
                # One final link harvest before exiting
                anchors = feed.locator("a[href*='/maps/place/'], a.hfpxzc")
                for i in range(anchors.count()):
                    try:
                        href = anchors.nth(i).get_attribute("href")
                        if href:
                            clean_url = href.split("?")[0] if "?" in href else href
                            if clean_url not in seen:
                                seen.add(clean_url)
                                seen.add(href)
                                collected.append(href)
                    except Exception:
                        pass
                print(f"Reached the end of Google Maps results. Found {len(collected)} total listings.")
                break

            if len(collected) == before:
                stagnant_rounds += 1
                # Wiggle scroll: scroll up slightly then down again to re-trigger observer
                try:
                    feed.evaluate("(el) => el.scrollBy(0, -500)")
                    _sleep(1.0, 1.5)
                    feed.evaluate("(el) => { el.scrollTop = el.scrollHeight; el.scrollBy(0, 4000); }")
                    page.keyboard.press("PageDown")
                    _sleep(1.5, 2.0)
                except Exception:
                    pass
            else:
                stagnant_rounds = 0
                print(f"   -> Loaded {len(collected)}/{target_count} business listings...")

        print(f"Total collected: {len(collected)} unique business links (target: {target_count}).")
        return collected[:target_count]

    # ------------------------------------------------------------------
    # Step 2: visit each listing and extract details
    # ------------------------------------------------------------------
    def _extract_business(self, page, url: str) -> dict:
        page.goto(url, timeout=60000)
        _sleep(1.5, 2.5)
        self._check_and_wait_for_captcha(page)

        # Wait for the business name heading to render
        try:
            page.wait_for_selector("h1", timeout=8000)
        except PWTimeoutError:
            pass

        name = _first_text(page, ["h1.DUwDvf", "h1.fontHeadlineLarge", "h1"])

        # Rating + review count usually live together, e.g. "4.5 (1,234)"
        rating = _first_text(
            page,
            [
                "div.F7nice span[aria-hidden='true']",
                "span.MW4etd",
                "div.fontDisplayLarge",
            ],
        )

        reviews = "N/A"
        try:
            review_el = page.locator("div.F7nice span[aria-label*='review']").first
            if review_el.count() > 0:
                aria = review_el.get_attribute("aria-label") or ""
                reviews = "".join(ch for ch in aria if ch.isdigit()) or "N/A"
        except Exception:
            pass
        if reviews == "N/A":
            raw = _first_text(page, ["span.UY7F9"])
            reviews = raw.strip("()") if raw != "N/A" else "N/A"

        address = "N/A"
        try:
            el = page.locator("button[data-item-id='address']").first
            if el.count() > 0:
                aria = el.get_attribute("aria-label") or ""
                address = aria.replace("Address:", "").strip() or "N/A"
        except Exception:
            pass

        phone = "N/A"
        try:
            el = page.locator("button[data-item-id^='phone:tel:']").first
            if el.count() > 0:
                aria = el.get_attribute("aria-label") or ""
                phone = aria.replace("Phone:", "").strip() or "N/A"
        except Exception:
            pass

        website = "N/A"
        try:
            el = page.locator("a[data-item-id='authority']").first
            if el.count() > 0:
                href = el.get_attribute("href")
                if href:
                    website = href.strip()
        except Exception:
            pass

        return {
            "Business Name": name,
            "Phone Number": phone,
            "Website": website,
            "Address": address,
            "Rating": rating,
            "Number of Reviews": reviews,
            "Google Maps URL": url,
        }

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------
    def run(self, query: str, count: int) -> list:
        results = []
        seen_keys = set()

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=self.headless, slow_mo=self.slow_mo)
            context = browser.new_context(
                viewport={"width": 1366, "height": 900},
                locale="en-US",
            )
            page = context.new_page()

            try:
                urls = self._collect_listing_urls(page, query, count)
            except Exception as e:
                print(f"Error while collecting listings: {e}")
                urls = []

            for idx, url in enumerate(urls, start=1):
                print(f"[{idx}/{len(urls)}] Scraping: {url}")
                try:
                    data = self._extract_business(page, url)
                except Exception as e:
                    print(f"  -> Failed to extract this listing: {e}")
                    continue

                # Deduplicate by name + address
                key = (data["Business Name"].lower(), data["Address"].lower())
                if key in seen_keys:
                    print("  -> Duplicate, skipping.")
                    continue
                seen_keys.add(key)

                results.append(data)
                print(f"  -> {data['Business Name']} | {data['Phone Number']} | {data['Rating']}")

                _sleep(1.5, 3.0)

                if len(results) >= count:
                    break

            browser.close()

        return results