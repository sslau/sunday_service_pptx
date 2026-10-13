#!/usr/bin/env python3
# make_section_gifs.py — render each service section's slides to a looping GIF.
"""Build one section deck per 崇拜 section (宣召經文 / 詩歌敬拜 / 讀經 / 講道 /
詩歌回應 / 家事分享), rasterize every slide, and write a looping GIF (3 s per
slide) so the user manual can embed them.

    python make_section_gifs.py                    # demo_week.json -> user_guide_assets/
    python make_section_gifs.py --date 2026.10.18  # pull a saved week from the store
    python make_section_gifs.py --week other.json --width 1280 --seconds 4

Rendering reuses the app's own pipeline — deck_builder.build_section_pptx_bytes
(what the per-section 產生…投影片 buttons call) and video_deck.render_pptx_slide
(the MP4 rasterizer) — so the GIFs match the built decks exactly.  Pure Pillow,
no LibreOffice or system ffmpeg.
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
os.chdir(HERE)

from PIL import Image  # noqa: E402

import deck_builder  # noqa: E402
import model  # noqa: E402

# (section key accepted by build_section_pptx_bytes, 標題, output filename)
SECTIONS = [
    ("psalm", "宣召經文", "section_psalm.gif"),
    ("hymns", "詩歌敬拜", "section_hymns.gif"),
    ("scripture", "讀經", "section_scripture.gif"),
    ("sermon", "講道", "section_sermon.gif"),
    ("response", "詩歌回應", "section_response.gif"),
    ("announcements", "家事分享", "section_announcements.gif"),
]


def load_week(args):
    """Week dict from --week <json>, --date <YYYY.MM.DD> (store), or demo_week.json."""
    if args.week:
        with open(args.week, "r", encoding="utf-8") as fh:
            return json.load(fh)
    if args.date:
        import tomllib

        from store import get_store

        secrets = {}
        secrets_path = os.path.join(HERE, ".streamlit", "secrets.toml")
        if os.path.isfile(secrets_path):
            with open(secrets_path, "rb") as fh:
                secrets = tomllib.load(fh)
        week = get_store(secrets).get(args.date, fresh=True)
        if not week:
            raise SystemExit("no week saved for %s" % args.date)
        return week
    with open(os.path.join(HERE, "demo_week.json"), encoding="utf-8") as fh:
        return json.load(fh)


def render_frames(pptx_bytes, count, width):
    """Every slide of a .pptx as a list of RGB Pillow images."""
    from video_deck import render_pptx_slide

    return [render_pptx_slide(pptx_bytes, i, width) for i in range(count)]


def write_gif(frames, path, seconds=3.0, colors=128):
    """Write `frames` as a looping GIF, `seconds` per slide, loop forever."""
    pals = [f.convert("RGB").quantize(colors=colors, method=Image.MEDIANCUT)
            for f in frames]
    pals[0].save(path, save_all=True, append_images=pals[1:],
                 duration=int(round(seconds * 1000)), loop=0,
                 optimize=True, disposal=2)


def main():
    parser = argparse.ArgumentParser(
        description="Render each service section's slides to a looping GIF.")
    parser.add_argument("--week", help="path to a week JSON")
    parser.add_argument("--date", help="date to pull from the store, e.g. 2026.10.18")
    parser.add_argument("--outdir", default="user_guide_assets",
                        help="output directory (default: user_guide_assets)")
    parser.add_argument("--width", type=int, default=960,
                        help="frame width in px, 16:9 (default: 960)")
    parser.add_argument("--seconds", type=float, default=3.0,
                        help="seconds per slide (default: 3)")
    parser.add_argument("--colors", type=int, default=128,
                        help="GIF palette size per frame (default: 128)")
    args = parser.parse_args()

    week = model.normalize_week(load_week(args))
    os.makedirs(args.outdir, exist_ok=True)
    print("source: %s" % (args.week or ("store %s" % args.date if args.date
                                        else "demo_week.json")))
    for section, label, filename in SECTIONS:
        data, count = deck_builder.build_section_pptx_bytes(section, week)
        if count < 1:
            print("%-13s %-8s — no slides, skipped" % (section, label))
            continue
        frames = render_frames(data, count, args.width)
        out = os.path.join(args.outdir, filename)
        write_gif(frames, out, args.seconds, args.colors)
        print("%-13s %-8s %2d slides  ->  %s  (%d KB)"
              % (section, label, count, out, os.path.getsize(out) // 1024))


if __name__ == "__main__":
    main()
