#!/usr/bin/env python3
# make_quick_guide_pdf.py — render the printable one-page guides to PDF.
"""Turn the printable one-page guides into PDFs with headless Google Chrome
via Playwright:

    quick_guide.html     -> quick_guide.pdf      (繁中)
    quick_guide.en.html  -> quick_guide.en.pdf   (English)

    python make_quick_guide_pdf.py                # both
    python make_quick_guide_pdf.py quick_guide.en.html

Needs the `playwright` package and Google Chrome (channel="chrome"); no browser
download required — only `python -m pip install playwright`.

After editing either HTML file, re-run this; the PDFs inherit the page's
`@page { size: A4 }` rule so each should stay a single page.
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT = ["quick_guide.html", "quick_guide.en.html"]


def count_pages(pdf_path):
    """Page count from the PDF page tree (/Count N), for a sanity check."""
    with open(pdf_path, "rb") as fh:
        data = fh.read()
    counts = [int(m) for m in re.findall(rb"/Count\s+(\d+)", data)]
    return max(counts) if counts else 0


def render(page, html_name):
    src = os.path.join(HERE, html_name)
    out = os.path.join(HERE, os.path.splitext(html_name)[0] + ".pdf")
    page.goto("file://" + src, wait_until="load")
    page.emulate_media(media="print")
    page.pdf(path=out, prefer_css_page_size=True, print_background=True,
             format="A4",
             margin={"top": "0", "right": "0", "bottom": "0", "left": "0"})
    pages = count_pages(out)
    print("%-22s -> %-22s (%d KB, %d page%s)"
          % (html_name, os.path.basename(out), os.path.getsize(out) // 1024,
             pages, "" if pages == 1 else "s"))
    if pages != 1:
        print("  WARNING: expected a single page — trim %s." % html_name)


def main():
    names = sys.argv[1:] or DEFAULT

    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=True)
        page = browser.new_page()
        for name in names:
            render(page, name)
        browser.close()


if __name__ == "__main__":
    main()
