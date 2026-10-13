#!/usr/bin/env python3
# make_ui_hymn_demo_gif.py — record the full hymn workflow to a looping GIF.
"""Drive the running app through the 詩歌敬拜 (Hymns) workflow and write a
looping GIF for the user manual:

    1. open step 2 and expand 🔍 歌詞庫搜尋
    2. search the lyric library (default term 中國啊)
    3. what the 結果 row shows
    4. 👁 看全首歌詞 pop-up
    5. ⬇️ 匯入整首詩歌 — lyrics + metadata land in the editor
    6. 🎨 背景圖 — pick a background thumbnail
    7. ▶ 預覽這首詩歌的 slides → 🎬 產生這首詩歌的 slide

    # start the app first, then:
    python make_ui_hymn_demo_gif.py
    python make_ui_hymn_demo_gif.py --term "讓我們 愛" --width 1100

Needs the `playwright` package and Google Chrome (channel="chrome"); no browser
download required — only `python -m pip install playwright`.
"""
import argparse
import os
import sys
from io import BytesIO

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
os.chdir(HERE)


def expand(page, text):
    """Open the first <details> whose summary contains `text`."""
    loc = page.locator('summary:has-text("%s")' % text).first
    if loc.count():
        loc.scroll_into_view_if_needed()
        loc.click()
        page.wait_for_timeout(1200)


def click(page, text, nth=0):
    loc = page.locator('button:has-text("%s")' % text).nth(nth)
    if loc.count():
        loc.scroll_into_view_if_needed()
        loc.click()
        return True
    return False


def grab(page, shots, tag):
    page.wait_for_timeout(600)
    shots.append((tag, page.screenshot()))


def capture(page, args):
    shots = []
    click(page, "2 詩歌")
    page.wait_for_timeout(args.settle)
    page.evaluate("window.scrollTo(0, 0)")
    expand(page, "歌詞庫搜尋")
    grab(page, shots, "search-open")

    box = page.get_by_label(args.search_label).first
    box.click()
    box.fill(args.term)
    box.press("Enter")
    page.wait_for_timeout(3000)
    box.scroll_into_view_if_needed()
    grab(page, shots, "search-results")

    if click(page, "看全首歌詞"):
        page.wait_for_timeout(1500)
        grab(page, shots, "whole-song")
        page.keyboard.press("Escape")
        page.wait_for_timeout(1200)

    if click(page, "匯入整首詩歌"):
        page.wait_for_timeout(3000)
        ta = page.locator('textarea[placeholder*="貼上整首歌詞"]').first
        val = ta.input_value() if ta.count() else ""
        print("imported lyric chars: %d" % len(val))
        ta.scroll_into_view_if_needed()
        grab(page, shots, "imported")

    expand(page, "背景圖")
    bg = page.locator('details:has(summary:has-text("背景圖"))').first
    cbs = bg.locator('[data-testid="stCheckbox"]')
    print("bg checkboxes: %d" % cbs.count())
    if cbs.count() > args.bg_index:
        cbx = cbs.nth(args.bg_index)
        cbx.scroll_into_view_if_needed()
        cbx.click()
        page.wait_for_timeout(2500)
    grab(page, shots, "background")

    expand(page, "預覽這首詩歌的 slide")
    if click(page, "產生這首詩歌的 slide"):
        page.wait_for_timeout(6000)
    page.locator('summary:has-text("預覽這首詩歌的 slide")').first.scroll_into_view_if_needed()
    grab(page, shots, "preview")
    return shots


def write_gif(shots, path, width, seconds, colors):
    pals = []
    for _tag, b in shots:
        f = Image.open(BytesIO(b)).convert("RGB")
        h = round(f.height * width / f.width)
        pals.append(f.resize((width, h)).quantize(colors=colors,
                                                  method=Image.MEDIANCUT))
    ms = int(round(seconds * 1000))
    pals[0].save(path, save_all=True, append_images=pals[1:], duration=ms,
                 loop=0, optimize=True, disposal=2)


def main():
    parser = argparse.ArgumentParser(
        description="Record the full hymn workflow to a looping GIF.")
    parser.add_argument("--url", default="http://localhost:8501")
    parser.add_argument("--out", default="user_guide_assets/section_hymns.gif")
    parser.add_argument("--term", default="中國啊", help="lyric search term")
    parser.add_argument("--search-label", default="搜尋歌詞")
    parser.add_argument("--bg-label", default="選用")
    parser.add_argument("--bg-index", type=int, default=1,
                        help="which background thumbnail to pick (default 1)")
    parser.add_argument("--width", type=int, default=1000)
    parser.add_argument("--seconds", type=float, default=1.6,
                        help="seconds per frame (default 1.6)")
    parser.add_argument("--colors", type=int, default=128)
    parser.add_argument("--view-width", type=int, default=1440)
    parser.add_argument("--view-height", type=int, default=900)
    parser.add_argument("--settle", type=int, default=3000)
    parser.add_argument("--timeout", type=int, default=60000)
    parser.add_argument("--channel", default="chrome")
    args = parser.parse_args()

    from playwright.sync_api import sync_playwright

    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(channel=args.channel, headless=True)
        page = browser.new_page(viewport={"width": args.view_width,
                                          "height": args.view_height},
                                device_scale_factor=2)
        page.goto(args.url, wait_until="domcontentloaded", timeout=args.timeout)
        page.wait_for_timeout(args.settle + 3000)
        shots = capture(page, args)
        browser.close()

    write_gif(shots, args.out, args.width, args.seconds, args.colors)
    print("frames: %s" % ", ".join(t for t, _b in shots))
    print("%d frames  ->  %s  (%d KB)"
          % (len(shots), args.out, os.path.getsize(args.out) // 1024))


if __name__ == "__main__":
    main()
