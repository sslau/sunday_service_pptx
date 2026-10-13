#!/usr/bin/env python3
# make_ui_hymn_demo_gif.py — record the full hymn workflow to a looping GIF.
"""Drive the running app through the 詩歌敬拜 (Hymns) workflow and write a
looping GIF for the user manual:

    1. open step 2 and expand 🔍 歌詞庫搜尋
    2. search the lyric library (default term 每一天)
    3. open the 結果 dropdown to show every matching song, then pick one
    4. the 👁 看全首歌詞 pop-up (opens automatically when a result is picked)
    5. ⬇️ 匯入整首詩歌 — lyrics + metadata land in the editor
    6. 🎨 背景圖 — pick a background thumbnail
    7. ▶ 預覽這首詩歌的 slides → 🎬 產生這首詩歌的 slide
    8. download that hymn's .pptx and rasterize every slide (the finished deck)

    # start the app first, then:
    python make_ui_hymn_demo_gif.py
    python make_ui_hymn_demo_gif.py --term "讓我們 愛" --width 1100 --seconds 3

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


def expand(page, text):
    """Open the first <details> whose summary contains `text`."""
    loc = page.locator('summary:has-text("%s")' % text).first
    if loc.count():
        loc.scroll_into_view_if_needed()
        _mark(loc)
        loc.click()
        page.wait_for_timeout(1200)


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
    """Click the hymn download button, then rasterize the .pptx it returns."""
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

    count = len(Presentation(BytesIO(data)).slides)
    count = min(count, args.max_slides)
    print("hymn pptx: %d KB, %d slides" % (len(data) // 1024, count))
    return [render_pptx_slide(data, i, args.width) for i in range(count)]


def capture(page, args):
    shots = []
    click(page, "2 詩歌")
    page.wait_for_timeout(args.settle)
    page.evaluate("window.scrollTo(0, 0)")
    expand(page, "歌詞庫搜尋")
    grab(page, shots, "search-open")

    box = page.get_by_label(args.search_label).first
    box.scroll_into_view_if_needed()
    _mark(box)
    box.click()
    box.fill(args.term)
    box.press("Enter")
    page.wait_for_timeout(3000)
    box.scroll_into_view_if_needed()
    grab(page, shots, "search-results")

    # open the 結果 dropdown so every matching song is listed, then pick one
    sel = page.locator('[data-testid="stSelectbox"]').filter(
        has=page.locator('label:has-text("%s")' % args.results_label)).first
    trig = sel.locator('[data-baseweb="select"]').first
    trig.scroll_into_view_if_needed()
    _mark(trig)
    trig.click()
    page.wait_for_timeout(1200)
    hit, picked = None, ""
    pick = page.locator('[role="option"]').filter(has_text=args.pick)
    if args.pick_book:
        pick = pick.filter(has_text=args.pick_book)
    if pick.count():
        hit = pick.first
        hit.scroll_into_view_if_needed()
        picked = hit.inner_text().splitlines()[0]
        _mark(hit)
    grab(page, shots, "results-list")
    if hit is not None:
        hit.scroll_into_view_if_needed()
        hit.click()
        page.wait_for_timeout(3000)
    chosen = sel.inner_text().splitlines()[-1].strip() if sel.count() else ""
    print("picked: %s  ->  selected: %s" % (picked[:60], chosen[:60]))
    if args.pick not in chosen:                 # fall back: keyboard-filter
        sel.locator('[data-baseweb="select"]').first.click()
        page.wait_for_timeout(600)
        page.keyboard.type(args.pick)
        page.wait_for_timeout(1200)
        page.locator('[role="option"]:has-text("%s")' % args.pick).first.click()
        page.wait_for_timeout(3000)
        chosen = sel.inner_text().splitlines()[-1].strip() if sel.count() else ""
        print("re-selected: %s" % chosen[:60])

    # picking a result pops the full-lyrics dialog open by itself
    dlg = page.locator('[data-testid="stDialog"], [role="dialog"]')
    if dlg.count():
        grab(page, shots, "whole-song")
        page.keyboard.press("Escape")
        page.wait_for_timeout(1200)
    elif click(page, "看全首歌詞"):
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
    if cbs.count() > args.bg_index:
        cbs.nth(args.bg_index).scroll_into_view_if_needed()
        _mark(cbs.nth(args.bg_index))
        cbs.nth(args.bg_index).click()
        page.wait_for_timeout(2500)
    grab(page, shots, "background")

    expand(page, "預覽這首詩歌的 slide")
    if click(page, "產生這首詩歌的 slide"):
        page.wait_for_timeout(6000)
    page.locator('summary:has-text("預覽這首詩歌的 slide")').first.scroll_into_view_if_needed()
    grab(page, shots, "preview")

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
        description="Record the full hymn workflow to a looping GIF.")
    parser.add_argument("--url", default="http://localhost:8501")
    parser.add_argument("--out", default="user_guide_assets/section_hymns.gif")
    parser.add_argument("--term", default="每一天", help="lyric search term")
    parser.add_argument("--search-label", default="搜尋歌詞")
    parser.add_argument("--results-label", default="結果",
                        help="label of the song-results dropdown")
    parser.add_argument("--pick", default="88每一天88",
                        help="title of the result to select")
    parser.add_argument("--pick-book", default="美樂頌",
                        help="collection shown in the result label")
    parser.add_argument("--bg-index", type=int, default=1,
                        help="which background thumbnail to pick (default 1)")
    parser.add_argument("--dl-label", default="下載這首詩歌的 pptx",
                        help="download button whose .pptx becomes the slides")
    parser.add_argument("--max-slides", type=int, default=8,
                        help="cap the rasterized deck frames (default 8)")
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
