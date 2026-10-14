#!/usr/bin/env python3
# make_ui_sermon_demo_gif.py — record the 講道 · OneDrive picker to a GIF.
"""Drive wizard step 5 (講道信息 · Sermon) and write a looping GIF for the user
manual that shows the OneDrive file picker:

    1. open step 5 · 講道信息 and the 🌐 從 OneDrive 讀取講道投影片 panel
    2. the file list appears — pick sermons_20260920.pptx
    3. the picked file is selected (ready for ⚡ 從 OneDrive 讀取並儲存)

Every frame shows a 👆 click marker on the control that was used.

The OneDrive listing normally needs Microsoft Graph credentials, so this script
starts its own throwaway app instance whose `onedrive.py` is replaced by a stub
that returns a canned file list, with a local (non-Google-Sheets) store.  No
real drive and no live sheet is touched.

    python make_ui_sermon_demo_gif.py
    python make_ui_sermon_demo_gif.py --file sermons_20260920.pptx

Pass `--url http://localhost:8501` to capture an already-running app instead
(then the names are whatever that app lists).

Needs the `playwright` package and Google Chrome (channel="chrome"); no browser
download required — only `python -m pip install playwright`.
"""
import argparse
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import urllib.request
from io import BytesIO

from PIL import Image

from ui_gif_cursor import draw_click, quantize_with_markers

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
os.chdir(HERE)

_DSF = 2                       # device_scale_factor used for the page
_PENDING = {"xy": None}        # click point for the next grab()

# --- throwaway app instance -------------------------------------------------
# A drop-in `onedrive.py` that lists a fixed set of decks (no network), plus a
# secrets.toml that switches the store to local and "configures" OneDrive.
_STUB_ONEDRIVE = '''\
"""Stub onedrive.py used only by make_ui_sermon_demo_gif.py (no network)."""
_FILES = [
    {"id": "f0", "name": "sermon_2026.09.20.pptx", "size": 118000,
     "folder": False},
    {"id": "f1", "name": "sermons_20260920.pptx", "size": 121000,
     "folder": False},
    {"id": "f2", "name": "sermons_20260913.pptx", "size": 120000,
     "folder": False},
    {"id": "f3", "name": "sermons_20261004.pptx", "size": 122000,
     "folder": False},
]


def _date_variants(date):
    s = str(date)
    out = {s, s.replace(".", "-"), s.replace(".", "")}
    parts = s.split(".")
    if len(parts) == 3 and parts[0] and parts[1] and parts[2]:
        out.add("%s.%s%s" % (parts[0], parts[1], parts[2]))
    return out


def folders(cfg, folder_path):
    return [{"id": "sub", "name": "2026.0920", "folder": True}]


def _date_subfolder(cfg, folder_path, date):
    return {"id": "sub", "name": "2026.0920", "folder": True}


def _list_items(cfg, drive_path):
    return list(_FILES)


def list_files(cfg, folder_path):
    return list(_FILES)


def files_in_date_folder(cfg, folder_path, date):
    return list(_FILES)


def find_mp4s(cfg, folder_path=None, limit=300):
    return []


def download_to(cfg, file_id, target_path):
    return False, "demo stub（示範用，不會真的下載）"
'''

_STUB_SECRETS = '''\
[store]
spreadsheet = ""

[onedrive]
tenant_id = "demo-tenant"
client_id = "demo-client"
client_secret = "demo-secret"
site_host = "demo.sharepoint.com"
site_path = "/sites/Demo"
folder_path = "Documents/崇拜投影片"
'''


def build_instance(app_dir):
    """A temp copy of the app (symlinked) with a stubbed OneDrive + local store."""
    tmp = tempfile.mkdtemp(prefix="sermon_demo_")
    for name in os.listdir(app_dir):
        if name in ("data", ".streamlit", "onedrive.py", "__pycache__"):
            continue
        os.symlink(os.path.join(app_dir, name), os.path.join(tmp, name))
    os.makedirs(os.path.join(tmp, "data"), exist_ok=True)
    std = os.path.join(tmp, ".streamlit")
    os.makedirs(std, exist_ok=True)
    cfg = os.path.join(app_dir, ".streamlit", "config.toml")
    if os.path.isfile(cfg):
        shutil.copy(cfg, os.path.join(std, "config.toml"))
    with open(os.path.join(tmp, "onedrive.py"), "w", encoding="utf-8") as fh:
        fh.write(_STUB_ONEDRIVE)
    with open(os.path.join(std, "secrets.toml"), "w", encoding="utf-8") as fh:
        fh.write(_STUB_SECRETS)
    return tmp


def _free_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def start_server(app_dir, timeout=90):
    """Launch the stubbed instance; returns (process, url, tmpdir)."""
    tmp = build_instance(app_dir)
    port = _free_port()
    url = "http://localhost:%d" % port
    log = open(os.path.join(tmp, "server.log"), "w", encoding="utf-8")
    proc = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", "app.py",
         "--server.port", str(port), "--server.headless", "true",
         "--server.fileWatcherType", "none"],
        cwd=tmp, stdout=log, stderr=subprocess.STDOUT)
    end = time.time() + timeout
    while time.time() < end:
        if proc.poll() is not None:
            raise SystemExit("app exited early; see %s/server.log" % tmp)
        try:
            with urllib.request.urlopen(url, timeout=3) as resp:
                if resp.status == 200:
                    return proc, url, tmp
        except Exception:
            time.sleep(0.5)
    proc.terminate()
    raise SystemExit("app did not come up; see %s/server.log" % tmp)


# --- capture ----------------------------------------------------------------
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
    # go to step 5 · 講道信息 (clicking it is safe even if already there)
    page.evaluate("window.scrollTo(0, 0)")
    step = page.locator('button:has-text("%s")' % args.step_label).first
    if step.count():
        step.click()
    page.wait_for_timeout(args.settle)

    # 1) the ready panel — flag the OneDrive panel to open
    exp = page.locator('[data-testid="stExpander"] summary:has-text("%s")'
                       % args.expander_label).first
    if not exp.count():
        raise SystemExit("could not find the %s panel" % args.expander_label)
    exp.scroll_into_view_if_needed()
    page.wait_for_timeout(300)
    _mark(exp)
    grab(page, shots, "step-open")

    # 2) open the panel -> the file list
    exp.scroll_into_view_if_needed()
    exp.click()
    page.wait_for_timeout(1200)
    opt = page.locator('[data-testid="stRadio"] label:has-text("%s")'
                       % args.file).first
    if not opt.count():
        raise SystemExit("file %r not found in the OneDrive list" % args.file)
    opt.scroll_into_view_if_needed()
    page.wait_for_timeout(400)
    _mark(opt)
    grab(page, shots, "onedrive-open")

    # pick the file
    opt.click()
    page.wait_for_timeout(2000)
    opt = page.locator('[data-testid="stRadio"] label:has-text("%s")'
                       % args.file).first
    opt.scroll_into_view_if_needed()
    _mark(opt)
    grab(page, shots, "file-picked")
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
        description="Record the 講道 (Sermon) OneDrive picker to a GIF.")
    parser.add_argument("--url", default=None,
                        help="capture an already-running app instead of "
                             "starting the built-in stubbed instance")
    parser.add_argument("--out",
                        default="user_guide_assets/section_sermon.gif")
    parser.add_argument("--step-label", default="5 講道")
    parser.add_argument("--expander-label", default="從 OneDrive 讀取講道投影片")
    parser.add_argument("--file", default="sermons_20260920.pptx",
                        help="file to pick in the 從 OneDrive list")
    parser.add_argument("--width", type=int, default=1000)
    parser.add_argument("--seconds", type=float, default=3.0,
                        help="seconds per frame (default 3)")
    parser.add_argument("--colors", type=int, default=128)
    parser.add_argument("--view-width", type=int, default=1440)
    parser.add_argument("--view-height", type=int, default=900)
    parser.add_argument("--settle", type=int, default=3500)
    parser.add_argument("--timeout", type=int, default=60000)
    parser.add_argument("--channel", default="chrome")
    args = parser.parse_args()

    from playwright.sync_api import sync_playwright

    proc = tmp = None
    url = args.url
    if not url:
        proc, url, tmp = start_server(HERE)
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(channel=args.channel, headless=True)
            page = browser.new_page(
                viewport={"width": args.view_width,
                          "height": args.view_height},
                device_scale_factor=_DSF)
            page.goto(url, wait_until="domcontentloaded", timeout=args.timeout)
            page.wait_for_timeout(args.settle + 3000)
            shots = capture(page, args)
            browser.close()
    finally:
        if proc is not None:
            proc.terminate()
            try:
                proc.wait(timeout=10)
            except Exception:
                proc.kill()
        if tmp:
            shutil.rmtree(tmp, ignore_errors=True)

    write_gif(shots, args.out, args.width, args.seconds, args.colors)
    print("frames: %s" % ", ".join(t for t, _b in shots))
    print("%d frames  ->  %s  (%d KB)"
          % (len(shots), args.out, os.path.getsize(args.out) // 1024))


if __name__ == "__main__":
    main()
