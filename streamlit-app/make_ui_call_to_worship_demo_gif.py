#!/usr/bin/env python3
# make_ui_call_to_worship_demo_gif.py — record the 宣召經文 workflow to a GIF.
"""Drive the running app through wizard step 1 (宣召經文 · Call to Worship) and
write a looping GIF for the user manual, showing the whole capture flow:

    1. open step 1 · 宣召經文
    2. 書卷 dropdown — type to filter the list, then pick 詩篇
    3. the 出處 auto-fills with the FHL book code; top it up to 詩 100:1-5
    4. 從 fhl.net 聖經網輸入 — the verses land in 內容

Every frame shows a 👆 click marker on the control that was just used, so the
demo reads like a walkthrough.

Note: on step 1 the fhl.net button also saves the passage to the week store, so
point this at an app whose store is local (no Google Sheets credentials) if you
do not want the demo to touch a live sheet.

    # start the app first, then:
    python make_ui_call_to_worship_demo_gif.py
    python make_ui_call_to_worship_demo_gif.py --ref "詩篇 23:1-6" --book 詩篇

Needs the `playwright` package and Google Chrome (channel="chrome"); no browser
download required — only `python -m pip install playwright`.
"""
import argparse
import os
import sys
from io import BytesIO

from PIL import Image

from ui_gif_cursor import draw_click, quantize_with_markers

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
os.chdir(HERE)

_DSF = 2                       # device_scale_factor used for the page
_PENDING = {"xy": None}        # click point for the next grab()


def _mark(loc):
    """Remember a locator's centre so the next grab() can flag the click."""
    try:
        bb = loc.bounding_box()
    except Exception:
        bb = None
    if bb:
        _PENDING["xy"] = (bb["x"] + bb["width"] / 2, bb["y"] + bb["height"] / 2)


def grab(page, shots, tag):
    page.wait_for_timeout(600)
    img = Image.open(BytesIO(page.screenshot())).convert("RGB")
    xy = _PENDING.get("xy")
    if xy:
        img = draw_click(img, xy[0] * _DSF, xy[1] * _DSF, scale=_DSF)
        _PENDING["xy"] = None
    shots.append((tag, img))


def capture(page, args):
    shots = []
    # make sure we are on step 1 · 宣召經文 (clicking it when already there is a
    # no-op, so this is safe either way)
    page.evaluate("window.scrollTo(0, 0)")
    step = page.locator('button:has-text("%s")' % args.step_label).first
    if step.count():
        _mark(step)
        step.click()
    page.wait_for_timeout(args.settle)
    page.evaluate("window.scrollTo(0, 0)")
    if step.count():
        _mark(step)                        # re-mark at the top, after settling
    grab(page, shots, "step-open")

    # 1) 書卷 dropdown — type to filter the virtual list
    sb = page.get_by_label(args.book_label).first
    sb.scroll_into_view_if_needed()
    sb.click()
    if args.book_filter:
        sb.fill(args.book_filter)
        page.wait_for_timeout(1300)
    _mark(sb)
    grab(page, shots, "book-search")

    # pick 詩篇 — flag the option while the list is still open, then click it
    opt = page.locator('[role="option"]:has-text("%s")' % args.book).first
    if opt.count():
        _mark(opt)
        grab(page, shots, "book-pick")
        opt.click()
        page.wait_for_timeout(2000)
    picked = page.get_by_label(args.ref_label).first
    print("出處 after book pick: %r"
          % (picked.input_value() if picked.count() else ""))
    if picked.count():
        picked.scroll_into_view_if_needed()
        _mark(picked)                      # hand on 出處, now showing the code
    grab(page, shots, "book-picked")

    # 2) the pick filled 出處 with the book code ("詩"); complete it to 詩 100:1-5
    ref = page.get_by_label(args.ref_label).first
    ref.scroll_into_view_if_needed()
    ref.click()
    ref.fill(args.ref)
    ref.press("Enter")
    page.wait_for_timeout(1500)
    print("出處 typed          : %r" % ref.input_value())
    ref.scroll_into_view_if_needed()
    _mark(ref)
    grab(page, shots, "ref-typed")

    # 3) load the verses from fhl.net into 內容
    btn = page.locator('button:has-text("%s")' % args.load_label).first
    if btn.count():
        btn.scroll_into_view_if_needed()
        _mark(btn)
        btn.click()
        page.wait_for_timeout(6000)
    ta = page.locator("textarea").first
    chars = ta.input_value() if ta.count() else ""
    print("verses chars: %d" % len(chars))
    if ta.count():
        ta.scroll_into_view_if_needed()
        page.wait_for_timeout(400)
    if btn.count():
        bb = btn.bounding_box()
        # the page scrolled to the verses, so re-aim the hand at the button's
        # *new* position; only draw it if the button is still on screen.
        if bb and -20 <= bb["y"] <= args.view_height - 20:
            _mark(btn)
        print("fetch button y after load: %s"
              % ("%.0f" % bb["y"] if bb else "n/a"))
    grab(page, shots, "loaded")
    return shots


def fit(img, size, bg=(255, 255, 255)):
    """Scale `img` to `size[0]` wide and centre it on a canvas of `size`."""
    if img.size == size:
        return img
    h = max(1, round(img.height * size[0] / img.width))
    canvas = Image.new("RGB", size, bg)
    canvas.paste(img.resize((size[0], h)), (0, max(0, (size[1] - h) // 2)))
    return canvas


def write_gif(frames, path, width, seconds, colors):
    """Downscale every frame to `width`, pad to a common size, write a GIF."""
    first = frames[0][1]
    size = (width, round(first.height * width / first.width))
    pals = [quantize_with_markers(fit(img, size), colors)
            for _tag, img in frames]
    pals[0].save(path, save_all=True, append_images=pals[1:],
                 duration=int(round(seconds * 1000)), loop=0,
                 optimize=True, disposal=2)


def main():
    parser = argparse.ArgumentParser(
        description="Record the 宣召經文 (Call to Worship) capture flow to a GIF.")
    parser.add_argument("--url", default="http://localhost:8501")
    parser.add_argument("--out", default="user_guide_assets/section_call_to_worship.gif")
    parser.add_argument("--step-label", default="1 宣召經文")
    parser.add_argument("--book", default="詩篇", help="book to pick from 書卷")
    parser.add_argument("--book-filter", default="詩",
                        help="text typed into 書卷 to filter the list")
    parser.add_argument("--book-label", default="書卷")
    parser.add_argument("--ref-label", default="出處")
    parser.add_argument("--ref", default="詩 100:1-5", help="出處 to fetch")
    parser.add_argument("--load-label", default="從 fhl.net 聖經網輸入")
    parser.add_argument("--width", type=int, default=1000)
    parser.add_argument("--seconds", type=float, default=3.0,
                        help="seconds per frame (default 3)")
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
                                device_scale_factor=_DSF)
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
