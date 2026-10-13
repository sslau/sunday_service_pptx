#!/usr/bin/env python3
# make_ui_scripture_demo_gif.py — record the scripture flow to a looping GIF.
"""Drive the app's 📜 經文投影片（獨立流程）and write a looping GIF for the user
manual — this is the scripture page that can actually *download* a deck:

    1. open the standalone 經文投影片 flow
    2. pick 書卷 from the book dropdown (filtered by typing)
    3. type the 出處 verses, e.g. 詩篇 34:1-10
    4. 從 fhl.net 聖經網輸入 — the verses land in 內容
    4. 📜 製成經文投影片 pptx → ⬇️ 下載經文投影片 pptx
    5. rasterize every slide of the downloaded .pptx (the finished deck)

    # start the app first, then:
    python make_ui_scripture_demo_gif.py
    python make_ui_scripture_demo_gif.py --ref "羅馬書 12:1-8" --book 羅馬書

Needs the `playwright` package and Google Chrome (channel=\"chrome\"); no browser
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


def click(page, text, nth=0):
    loc = page.locator('button:has-text("%s")' % text).nth(nth)
    if loc.count():
        loc.scroll_into_view_if_needed()
        _mark(loc)
        loc.click()
        return True
    return False


def grab(page, shots, tag):
    page.wait_for_timeout(600)
    img = Image.open(BytesIO(page.screenshot())).convert("RGB")
    xy = _PENDING.get("xy")
    if xy:
        img = draw_click(img, xy[0] * _DSF, xy[1] * _DSF, scale=_DSF)
        _PENDING["xy"] = None
    shots.append((tag, img))


def download_slides(page, args, shots):
    """Click the scripture download button, then rasterize the .pptx."""
    loc = page.locator('button:has-text("%s")' % args.dl_label).first
    if not loc.count():
        print("no download button; skipping slide frames")
        return []
    loc.scroll_into_view_if_needed()
    if shots is not None:                      # frame with the hand on the button
        _mark(loc)
        grab(page, shots, "download")
    with page.expect_download(timeout=args.timeout) as dl_info:
        loc.click()
    path = dl_info.value.path()
    with open(path, "rb") as fh:
        data = fh.read()
    from pptx import Presentation

    from video_deck import render_pptx_slide

    count = min(len(Presentation(BytesIO(data)).slides), args.max_slides)
    print("scripture pptx: %d KB, %d slides" % (len(data) // 1024, count))
    return [render_pptx_slide(data, i, args.width) for i in range(count)]


def capture(page, args):
    shots = []
    # open the standalone 📜 經文投影片 flow from the main toolbar
    click(page, args.open_label)
    page.wait_for_timeout(args.settle)
    page.evaluate("window.scrollTo(0, 0)")
    grab(page, shots, "flow-open")

    # 1) pick the book from the 書卷 dropdown (type to filter the virtual list)
    sb = page.get_by_label(args.book_label).first
    sb.scroll_into_view_if_needed()
    _mark(sb)
    sb.click()
    sb.fill(args.book)
    page.wait_for_timeout(1200)
    grab(page, shots, "book-dropdown")

    opt = page.locator('[role="option"]:has-text("%s")' % args.book).first
    _mark(opt)
    opt.click()
    page.wait_for_timeout(2000)
    grab(page, shots, "book-picked")

    # 2) type the chapter:verse into 出處
    ref = page.get_by_label(args.ref_label).first
    ref.scroll_into_view_if_needed()
    _mark(ref)
    ref.click()
    ref.fill(args.ref)
    ref.press("Enter")
    page.wait_for_timeout(1500)
    grab(page, shots, "ref-typed")

    # 3) load the verses from fhl.net
    if click(page, args.load_label):
        page.wait_for_timeout(6000)
    ta = page.locator("textarea").first
    chars = ta.input_value() if ta.count() else ""
    print("verses chars: %d" % len(chars))
    ta.scroll_into_view_if_needed()
    grab(page, shots, "loaded")

    # 4) compile the scripture .pptx
    if click(page, args.build_label):
        page.wait_for_timeout(5000)
    loc = page.locator('button:has-text("%s")' % args.dl_label).first
    if loc.count():
        loc.scroll_into_view_if_needed()
    grab(page, shots, "generated")

    # 5) download it and animate every slide of the finished deck
    for i, img in enumerate(download_slides(page, args, shots), 1):
        shots.append(("slide-%d" % i, img))
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
    first = frames[0][1]
    size = (width, round(first.height * width / first.width))
    pals = [quantize_with_markers(fit(img, size), colors)
            for _tag, img in frames]
    pals[0].save(path, save_all=True, append_images=pals[1:],
                 duration=int(round(seconds * 1000)), loop=0,
                 optimize=True, disposal=2)


def main():
    parser = argparse.ArgumentParser(
        description="Record the scripture flow (input → fhl.net → download).")
    parser.add_argument("--url", default="http://localhost:8501")
    parser.add_argument("--out", default="user_guide_assets/section_scripture.gif")
    parser.add_argument("--book", default="詩篇", help="book to pick in 書卷")
    parser.add_argument("--ref", default="詩篇 34:1-10", help="出處 to fetch")
    parser.add_argument("--book-label", default="書卷")
    parser.add_argument("--ref-label", default="出處")
    parser.add_argument("--open-label", default="獨立流程")
    parser.add_argument("--load-label", default="從 fhl.net 聖經網輸入")
    parser.add_argument("--build-label", default="製成經文投影片")
    parser.add_argument("--dl-label", default="下載經文投影片 pptx")
    parser.add_argument("--max-slides", type=int, default=10)
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
