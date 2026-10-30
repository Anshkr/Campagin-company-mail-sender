"""
main.py
-------
Entry point for the Google Maps business scraper.

Run:
    python main.py

You'll be asked for a search query (e.g. "finance company in Gurgaon")
and how many businesses to collect. The script then drives a real,
visible Chromium browser via Playwright to search Google Maps and
extract business details, printing them to the terminal and saving
them to businesses.csv and businesses.xlsx.
"""

import sys

import pandas as pd

from scraper import GoogleMapsScraper


def get_query() -> str:
    query = input("Enter your Google search query: ").strip()
    while not query:
        query = input("Please enter a non-empty query: ").strip()
    return query


def get_count() -> int:
    while True:
        raw = input("How many businesses do you want to collect? ").strip()
        try:
            count = int(raw)
            if count > 0:
                return count
            print("Please enter a positive number.")
        except ValueError:
            print("Please enter a valid integer.")


def print_results_table(results: list):
    if not results:
        print("\nNo results collected.")
        return

    print("\n" + "=" * 100)
    print(f"Collected {len(results)} business(es):")
    print("=" * 100)
    for i, biz in enumerate(results, start=1):
        print(f"\n{i}. {biz['Business Name']}")
        print(f"   Phone:    {biz['Phone Number']}")
        print(f"   Website:  {biz['Website']}")
        print(f"   Address:  {biz['Address']}")
        print(f"   Rating:   {biz['Rating']}  ({biz['Number of Reviews']} reviews)")
        print(f"   Maps URL: {biz['Google Maps URL']}")
    print("\n" + "=" * 100)


def save_results(results: list):
    if not results:
        print("Nothing to save.")
        return

    df = pd.DataFrame(
        results,
        columns=[
            "Business Name",
            "Phone Number",
            "Website",
            "Address",
            "Rating",
            "Number of Reviews",
            "Google Maps URL",
        ],
    )

    df.to_csv("businesses.csv", index=False, encoding="utf-8-sig")
    df.to_excel("businesses.xlsx", index=False)

    print("\nSaved results to:")
    print("  - businesses.csv")
    print("  - businesses.xlsx")


def main():
    print("=" * 60)
    print(" Google Maps Business Scraper")
    print("=" * 60)

    query = get_query()
    count = get_count()

    scraper = GoogleMapsScraper(headless=False)

    try:
        results = scraper.run(query, count)
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
        sys.exit(1)

    print_results_table(results)
    save_results(results)


if __name__ == "__main__":
    main()