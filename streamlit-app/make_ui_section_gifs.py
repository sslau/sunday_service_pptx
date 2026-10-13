#!/usr/bin/env python3
# make_ui_section_gifs.py — record each app section's UI to a looping GIF.
"""Drive the running app with a headless browser, scroll through one wizard
step at a time, and write a looping GIF per section for the user manual.

This is the UI-screenshot counterpart of make_section_gifs.py: instead of
rasterizing the generated slides, it captures what the app actually shows for
that step (top of the panel, scrolling down to the bottom).

    # start the app first, then:
    python make_ui_section_gifs.py                    # localhost:8501 -> user_guide_assets/
    python make_ui_section_gifs.py --sections hymns   # just one section
    python make_ui_section_gifs.py --width 1100 --seconds 0.7

Needs the `playwright` package and Google Chrome.  It drives the system Chrome
(channel="chrome"), so `playwright install` / a browser download is NOT
required; only the pip package:

    python -m pip install playwright
"""
import argparse
import os

from PIL import Image

# (wizard step button label, section key -> section_<key>.gif, 標題)
SECTIONS = [
    ("1 宣召經文", "psalm"),
    ("2 詩歌", "hymns"),
    ("4 讀經", "scripture"),
    ("5 講道", "sermon"),
    ("6 詩歌回應", "response"),
    ("8 家事分享", "announcements"),
]

# Find the largest scrollable element (Streamlit's main content pane) and the
# vertical offsets to screenshot from top to bottom.
SCROLLER_JS = """() => {
    let best = null, bd = 0;
    for (const el of document.querySelectorAll('html,body,div,section,main')) {
        const d = el.scrollHeight - el.clientHeight;
        if (d > bd && el.clientHeight > 200) { bd = d; best = el; }
    }
    return best || document.scrollingElement;
}"""


def click(page, label):
    loc = page.locator('button:has-text("%s")' % label).first
    if loc.count():
        loc.click()
        return True
    return False


def scroll_offsets(sh, ch, max_frames):
    span = max(0, sh - ch)
    if span <= 4:
        return [0]
    step = max(150, int(ch * 0.35))
    n = min(max_frames, max(2, -(-span // step) + 1))
    return [round(span * i / (n - 1)) for i in range(n)]


def section_frames(page, label, args):
    """Click a step and return screenshots scrolled from top to bottom."""
    click(page, label)
    page.wait_for_timeout(args.settle)
    el = page.evaluate_handle(SCROLLER_JS).as_element()
    if el is None:
        page.evaluate("window.scrollTo(0, 0)")
        return [page.screenshot()]
    info = el.evaluate("e => ({h: e.clientHeight, sh: e.scrollHeight})")
    shots = []
    for y in scroll_offsets(info["sh"], info["h"], args.max_frames):
        el.evaluate("(e, y) => { e.scrollTop = y; }", y)
        page.wait_for_timeout(args.settle_frame)
        shots.append(page.screenshot())
    el.evaluate("e => { e.scrollTop = 0; }")
    return shots


def write_gif(shots, path, width, seconds, colors):
    """Downscale to `width` and write a looping GIF, holding the last frame."""
    pals = []
    for b in shots:
        f = Image.open(__import__("io").BytesIO(b)).convert("RGB")
        h = round(f.height * width / f.width)
        pals.append(f.resize((width, h)).quantize(colors=colors,
                                                  method=Image.MEDIANCUT))
    ms = int(round(seconds * 1000))
    durations = [ms] * (len(pals) - 1) + [ms + 1500]      # linger on the bottom
    pals[0].save(path, save_all=True, append_images=pals[1:], duration=durations,
                 loop=0, optimize=True, disposal=2)


def main():
    parser = argparse.ArgumentParser(
        description="Record each app section's UI to a looping GIF.")
    parser.add_argument("--url", default="http://localhost:8501")
    parser.add_argument("--outdir", default="user_guide_assets")
    parser.add_argument("--sections", nargs="*",
                        help="section keys to build (default: all)")
    parser.add_argument("--width", type=int, default=1000)
    parser.add_argument("--seconds", type=float, default=0.7,
                        help="seconds per scroll frame (default: 0.7)")
    parser.add_argument("--colors", type=int, default=128)
    parser.add_argument("--view-width", type=int, default=1440)
    parser.add_argument("--view-height", type=int, default=900)
    parser.add_argument("--max-frames", type=int, default=8)
    parser.add_argument("--settle", type=int, default=2500,
                        help="ms to wait after clicking a step (default: 2500)")
    parser.add_argument("--settle-frame", type=int, default=280,
                        help="ms to wait after each scroll (default: 280)")
    parser.add_argument("--timeout", type=int, default=60000)
    parser.add_argument("--channel", default="chrome")
    args = parser.parse_args()

    wanted = set(args.sections) if args.sections else None
    todo = [(lbl, key) for lbl, key in SECTIONS if not wanted or key in wanted]
    if not todo:
        raise SystemExit("no matching sections for %s" % args.sections)

    from playwright.sync_api import sync_playwright

    os.makedirs(args.outdir, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(channel=args.channel, headless=True)
        page = browser.new_page(viewport={"width": args.view_width,
                                          "height": args.view_height},
                                device_scale_factor=2)
        page.goto(args.url, wait_until="domcontentloaded", timeout=args.timeout)
        page.wait_for_timeout(args.settle)
        for label, key in todo:
            shots = section_frames(page, label, args)
            out = os.path.join(args.outdir, "section_%s.gif" % key)
            write_gif(shots, out, args.width, args.seconds, args.colors)
            print("%-9s %2d frames  ->  %s  (%d KB)"
                  % (key, len(shots), out, os.path.getsize(out) // 1024))
        browser.close()


if __name__ == "__main__":
    main()
