#!/usr/bin/env python3
# make_quick_guide_pdf.py — render quick_guide.html to a one-page A4 PDF.
"""Turn the printable one-page guide (`quick_guide.html`, 繁中) into
`quick_guide.pdf` using headless Google Chrome via Playwright.

    python make_quick_guide_pdf.py

Needs the `playwright` package and Google Chrome (channel="chrome"); no browser
download required — only `python -m pip install playwright`.

After editing quick_guide.html, re-run this; the PDF inherits the page's
`@page { size: A4 }` rule so it should stay a single page.
"""
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))


def count_pages(pdf_path):
    """Page count from the PDF page tree (/Count N), for a sanity check."""
    with open(pdf_path, "rb") as fh:
        data = fh.read()
    counts = [int(m) for m in re.findall(rb"/Count\s+(\d+)", data)]
    return max(counts) if counts else 0


def main():
    src = os.path.join(HERE, "quick_guide.html")
    out = os.path.join(HERE, "quick_guide.pdf")

    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=True)
        page = browser.new_page()
        page.goto("file://" + src, wait_until="load")
        page.emulate_media(media="print")
        page.pdf(path=out, prefer_css_page_size=True, print_background=True,
                 format="A4",
                 margin={"top": "0", "right": "0", "bottom": "0", "left": "0"})
        browser.close()

    pages = count_pages(out)
    print("%s  (%d KB, %d page%s)"
          % (out, os.path.getsize(out) // 1024, pages, "" if pages == 1 else "s"))
    if pages != 1:
        print("WARNING: expected a single page — trim quick_guide.html.")


if __name__ == "__main__":
    main()
