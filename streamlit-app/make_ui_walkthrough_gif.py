#!/usr/bin/env python3
# make_ui_walkthrough_gif.py — record the app's wizard UI to a looping GIF.
"""Drive the running app with a headless browser, screenshot each wizard step,
and write a looping GIF (3 s per frame) for the user manual's 速覽 section.

    # start the app first, then:
    python make_ui_walkthrough_gif.py                   # localhost:8501 -> user_guide_assets/
    python make_ui_walkthrough_gif.py --url https://… --width 1280 --seconds 2.5
    python make_ui_walkthrough_gif.py --no-tab2         # skip the 製作成投影片 tab

Needs the `playwright` package and Google Chrome.  Playwright drives the system
Chrome (channel="chrome"), so `playwright install` / a browser download is NOT
required; only the pip package:

    python -m pip install playwright
"""
import argparse
import os

from PIL import Image

# Wizard step buttons, in order (labels as shown at the top of the edit page).
STEPS = ["1 宣召經文", "2 詩歌", "3 獻詩", "4 讀經", "5 講道",
         "6 詩歌回應", "7 聖餐", "8 家事分享", "9 家事MP4"]
TAB2 = "製作成投影片"


def click(page, label):
    """Click the first button whose visible text contains `label`."""
    loc = page.locator('button:has-text("%s")' % label).first
    if loc.count():
        loc.click()
        return True
    return False


def capture(args):
    """Return a list of RGB screenshots, one per wizard step (+ tab 2)."""
    from playwright.sync_api import sync_playwright

    shots = []
    with sync_playwright() as p:
        browser = p.chromium.launch(channel=args.channel, headless=True)
        page = browser.new_page(viewport={"width": args.view_width,
                                          "height": args.view_height},
                                device_scale_factor=2)
        page.goto(args.url, wait_until="domcontentloaded", timeout=args.timeout)
        page.wait_for_timeout(args.settle)          # let Streamlit finish booting
        targets = [(s, None) for s in STEPS]
        if not args.no_tab2:
            targets.append((TAB2, "tab2"))
        for label, key in targets:
            click(page, label)                      # step 1 is the default view too
            page.wait_for_timeout(args.settle)
            page.evaluate("window.scrollTo(0, 0)")
            page.wait_for_timeout(300)
            shots.append(page.screenshot())
        browser.close()
    return [Image.open(__import__("io").BytesIO(b)).convert("RGB") for b in shots]


def write_gif(frames, path, width, seconds, colors):
    """Downscale to `width` and write a looping GIF, `seconds` per frame."""
    pals = []
    for f in frames:
        h = round(f.height * width / f.width)
        pals.append(f.resize((width, h)).quantize(colors=colors,
                                                  method=Image.MEDIANCUT))
    pals[0].save(path, save_all=True, append_images=pals[1:],
                 duration=int(round(seconds * 1000)), loop=0,
                 optimize=True, disposal=2)


def main():
    parser = argparse.ArgumentParser(
        description="Record the app's wizard UI to a looping GIF.")
    parser.add_argument("--url", default="http://localhost:8501",
                        help="app URL (default: http://localhost:8501)")
    parser.add_argument("--outdir", default="user_guide_assets",
                        help="output directory (default: user_guide_assets)")
    parser.add_argument("--name", default="ui_walkthrough.gif",
                        help="output filename (default: ui_walkthrough.gif)")
    parser.add_argument("--channel", default="chrome",
                        help="Playwright browser channel (default: chrome)")
    parser.add_argument("--view-width", type=int, default=1440)
    parser.add_argument("--view-height", type=int, default=900)
    parser.add_argument("--width", type=int, default=1200,
                        help="GIF frame width in px (default: 1200)")
    parser.add_argument("--seconds", type=float, default=3.0,
                        help="seconds per frame (default: 3)")
    parser.add_argument("--colors", type=int, default=128,
                        help="GIF palette size per frame (default: 128)")
    parser.add_argument("--settle", type=int, default=2500,
                        help="ms to wait after each click (default: 2500)")
    parser.add_argument("--timeout", type=int, default=60000,
                        help="page load timeout in ms (default: 60000)")
    parser.add_argument("--no-tab2", action="store_true",
                        help="do not include the 製作成投影片 tab")
    args = parser.parse_args()

    frames = capture(args)
    os.makedirs(args.outdir, exist_ok=True)
    out = os.path.join(args.outdir, "ui_walkthrough.gif")
    write_gif(frames, out, args.width, args.seconds, args.colors)
    print("%d frames  ->  %s  (%d KB)"
          % (len(frames), out, os.path.getsize(out) // 1024))


if __name__ == "__main__":
    main()
