"""Render announcement slides (a subset) into an MP4, pure Python.

Runs on the Streamlit server with no System dependencies: each announcement
slide is drawn with Pillow to mirror the pptx layout produced by
generate_deck.build_announcements_section (solid navy bg / full-bleed image,
36pt centered title, left-aligned auto-fitted body), then encoded to H.264
MP4 via imageio + imageio-ffmpeg (pip-only; the wheel bundles its own ffmpeg
binary).  A CJK font is bundled under fonts/ so Chinese text renders even on
a bare Linux server.
"""
import io
import os

import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageOps
from pptx import Presentation
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR

Image.MAX_IMAGE_PIXELS = None  # trusted conference uploads — silence the bomb check

HERE = os.path.dirname(os.path.abspath(__file__))

_SLIDE_W_IN = 13.333
_SLIDE_H_IN = 7.5
_MARGIN_IN = 0.83
_BG = "#1F3A6E"           # mirrors _new_slide default
_TEXT = "#FFFFFF"
_PT_SCALE = 96.0 / (72.0 * _SLIDE_W_IN)  # px-per-pt-per-frame-width

_FONT_CANDIDATES = [
    os.path.join(HERE, "fonts", "NotoSansCJKsc-Regular.otf"),
    os.path.join(HERE, "fonts", "NotoSansSC-Regular.ttf"),
    "/System/Library/Fonts/PingFang.ttc",           # macOS
    "/System/Library/Fonts/Hiragino Sans GB.ttc",   # macOS
    "/System/Library/Fonts/Songti.ttc",             # macOS
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",  # Linux
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "C:/Windows/Fonts/msyh.ttc",                    # Windows
]

_font_path_cache = [None]


def _cjk_font_path():
    """First existing CJK font (bundled fonts/ preferred), cached."""
    if _font_path_cache[0] is None:
        _font_path_cache[0] = next(
            (p for p in _FONT_CANDIDATES if os.path.isfile(p)), None)
    return _font_path_cache[0]


def _load_font(pt, width):
    """PIL font for a given pptx pt size at frame width `width` (px).
    16:9 slide 13.333in wide: 1pt = width/13.333/72 px."""
    px = max(1, round(pt * width / _SLIDE_W_IN / 72.0 * 96.0))
    path = _cjk_font_path()
    if path is None:
        return ImageFont.load_default()
    return ImageFont.truetype(path, px)


_NAMED_FONT_PATHS = {
    "DFKAI-SB": [
        os.path.expanduser("~/Library/Fonts/kaiu.ttf"),
        "/Library/Fonts/DFKai-SB.ttf",
        "/System/Library/Fonts/Supplemental/DFKai-SB.ttc",
    ],
    "KAIU": [
        os.path.expanduser("~/Library/Fonts/kaiu.ttf"),
        "/Library/Fonts/DFKai-SB.ttf",
    ],
}
_font_cache = {}


def _font_for(font_name, px):
    """Best font for a pptx run (its declared name if resolvable locally, else
    the bundled CJK fallback), cached by (name, size)."""
    key = (str(font_name or "").upper(), int(px))
    if key in _font_cache:
        return _font_cache[key]
    cands = list(_NAMED_FONT_PATHS.get(key[0], [])) + _FONT_CANDIDATES
    font = None
    for p in cands:
        if os.path.isfile(p):
            try:
                font = ImageFont.truetype(p, int(px))
                break
            except Exception:
                continue
    if font is None:
        font = ImageFont.load_default()
    _font_cache[key] = font
    return font


def _run_font(run, width):
    """px font for a pptx run at frame width `width`; falls back to
    18pt/#FFFFFF when size is unset."""
    try:
        pt = run.font.size.pt
    except Exception:
        pt = 18
    return _font_for(run.font.name or None,
                     max(1, round(pt * width * _PT_SCALE)))


def _render_pptx_text(img, tf, box, sc, width):
    """Draw a text frame's paragraphs into Pillow using the run geometry."""
    try:
        m_l, m_r = (tf.margin_left or 0) * sc, (tf.margin_right or 0) * sc
        m_t = (tf.margin_top or 0) * sc
    except Exception:
        m_l = m_r = m_t = 0
    x0, y0 = box[0] + m_l, box[1] + m_t
    maxw = (box[2] - box[0]) - m_l - m_r
    try:
        anchor = tf.vertical_anchor
    except Exception:
        anchor = None

    wrapped = []
    for p in tf.paragraphs:
        runs = [r for r in p.runs if r.text]
        if not runs:
            wrapped.append([])
            continue
        line, cur = [], 0.0
        for r in runs:
            font = _run_font(r, width)
            for ch in r.text:
                cw = font.getlength(ch)
                if cur + cw > maxw and cur > 0:
                    wrapped.append(line)
                    line, cur = [], 0.0
                line.append((ch, r, font))
                cur += cw
        wrapped.append(line)

    line_hs = []
    for ln in wrapped:
        lh = 0
        for _ch, _r, font in ln:
            lh = max(lh, font.size)
        line_hs.append(max(1, round(lh * 1.3)))
    block_h = sum(line_hs)

    y = y0
    if anchor in (MSO_ANCHOR.MIDDLE, MSO_ANCHOR.BOTTOM):
        avail = (box[3] - box[1]) - m_t
        if anchor == MSO_ANCHOR.MIDDLE:
            y = y0 + max(0, (avail - block_h) / 2)
        else:
            y = y0 + max(0, avail - block_h)

    for ln, lh in zip(wrapped, line_hs):
        if not ln:
            y += lh
            continue
        try:
            algn = ln[0][1].paragraph.alignment
        except Exception:
            algn = None
        total = sum(f.getlength(ch) for ch, _r, f in ln)
        if algn == 2:          # CENTER
            x = x0 + (maxw - total) / 2
        elif algn == 3:        # RIGHT
            x = x0 + maxw - total
        else:
            x = x0
        for ch, r, font in ln:
            try:
                if r.font.color and r.font.color.type is not None:
                    col = str(r.font.color.rgb)
                else:
                    col = None
            except Exception:
                col = None
            fill = "#%s" % col if col else _TEXT
            draw = ImageDraw.Draw(img)
            draw.text((x, y), ch, font=font, fill=fill)
            x += font.getlength(ch)
        y += lh


def render_pptx_slide(data, index, width=1920):
    """Rasterize one slide of an uploaded .pptx at `width` px, honouring the
    slide geometry: solid fills, pictures at their real boxes, and every text
    box/run (font, size, color, alignment, vertical anchor). Pure python-pptx
    + Pillow — no LibreOffice required."""
    prs = Presentation(io.BytesIO(data))
    slide = prs.slides[index]
    sw, sh = float(prs.slide_width), float(prs.slide_height)
    if width % 2:
        width += 1
    height = max(1, round(width * sh / sw))
    sc = width / sw
    img = Image.new("RGB", (width, height), _BG)
    fills, pics, texts = [], [], []
    for shp in slide.shapes:
        try:
            box = (shp.left * sc, shp.top * sc,
                   (shp.left + shp.width) * sc, (shp.top + shp.height) * sc)
        except Exception:
            continue
        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                pics.append((box, shp.image.blob))
            except Exception:
                pass
        elif shp.has_text_frame and shp.text_frame.text.strip():
            texts.append((shp, box))
        else:
            try:
                if shp.fill.type == MSO_FILL_TYPE.SOLID:
                    fills.append((box[0], box[1], box[2], box[3],
                                  str(shp.fill.fore_color.rgb)))
            except Exception:
                pass
    for box in fills:
        if box[4]:
            ImageDraw.Draw(img).rectangle(box[:4], fill="#%s" % box[4])
    for box, blob in pics:
        try:
            im = Image.open(io.BytesIO(blob)).convert("RGB")
            im = im.resize((max(1, round(box[2] - box[0])),
                            max(1, round(box[3] - box[1]))),
                           Image.LANCZOS)
            img.paste(im, (round(box[0]), round(box[1])))
        except Exception:
            pass
    for shp, box in texts:
        _render_pptx_text(img, shp.text_frame, box, sc, width)
    return img


def _encode_mp4(frames, fps, progress=None):
    """H.264-encode `frames` into an MP4 (bytes).  When `progress` is given it
    is called with a float 0..1 after each frame is appended."""
    if not frames:
        raise ValueError("沒有可產製的投影片")
    import imageio
    buf = io.BytesIO()
    total = len(frames)
    writer = imageio.get_writer(buf, format="mp4", fps=float(fps),
                                codec="libx264", macro_block_size=1)
    try:
        for n, f in enumerate(frames, 1):
            writer.append_data(f)
            if progress:
                progress(n / total)
    finally:
        writer.close()
    return buf.getvalue()


def build_images_mp4(images, secs_per_slide=5.0, fps=15.0, width=1920,
                     progress=None):
    """Encode a list of already-final images (PNG/JPEG/WEBP bytes, one per
    slide) into an MP4 — pixel-exact by definition, no re-drawing."""
    if width % 2:
        width += 1
    count_per = max(1, round(float(secs_per_slide) * float(fps)))
    frames = []
    for blob in images:
        im = Image.open(io.BytesIO(blob)).convert("RGB")
        frame = ImageOps.fit(im, (width, max(1, round(width * _SLIDE_H_IN
                                                      / _SLIDE_W_IN))),
                             Image.LANCZOS)
        arr = np.asarray(frame)
        frames.extend([arr] * count_per)
    return _encode_mp4(frames, fps, progress)


def build_pptx_slides_mp4(data, indices, secs_per_slide=5.0, fps=15.0,
                          width=1920, progress=None):
    """Encode selected slides of an uploaded .pptx — rasterized from the file
    itself (exact image, not re-drawn text) — into an MP4."""
    count_per = max(1, round(float(secs_per_slide) * float(fps)))
    frames = []
    for i in indices:
        frame = render_pptx_slide(data, i, width)
        arr = np.asarray(frame)
        frames.extend([arr] * count_per)
    return _encode_mp4(frames, fps, progress)


def _margin_px(width):
    return _MARGIN_IN * width / _SLIDE_W_IN


def _announcement_frame(a, width, image_bytes=None):
    """Draw one announcement slide as a (w×h) RGB frame, mirroring the deck
    layout.  `image_bytes` renders the full-bleed image announcement."""
    height = round(width * _SLIDE_H_IN / _SLIDE_W_IN)
    if image_bytes:
        im = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        return ImageOps.fit(im, (width, height), Image.LANCZOS)

    img = Image.new("RGB", (width, height), _BG)
    draw = ImageDraw.Draw(img)
    margin = _margin_px(width)
    box_w = width - 2 * margin

    title = (a.get("title") or "").strip()
    if title:
        size = 36
        font = _load_font(size, width)
        while draw.textlength(title, font=font) > box_w and size > 24:
            size -= 2
            font = _load_font(size, width)
        # title box top 0.5in, height 1.0in -> vertical centre at 1.0in
        cy = 1.0 * width / _SLIDE_W_IN
        draw.text((width / 2, cy), title, font=font, fill=_TEXT, anchor="mm")

    body = [ln for ln in (a.get("body") or []) if str(ln).strip()]
    if body:
        body = [str(ln) for ln in body]
        size = 36
        line_h = font = None
        max_h = 5.6 * width / _SLIDE_W_IN
        while size >= 24:
            font = _load_font(size, width)
            line_h = round(font.size * 1.25)
            fits_w = all(draw.textlength(ln, font=font) <= box_w for ln in body)
            if fits_w and len(body) * line_h <= max_h:
                break
            size -= 1
        y = 1.6 * width / _SLIDE_W_IN
        for ln in body:
            draw.text((margin, y), ln, font=font, fill=_TEXT)
            y += line_h
    return img


def announcements_from_pptx(data):
    """Parse an announcements .pptx into the same list-of-dicts shape used by
    build_announcements_mp4 (one dict per slide): title = first non-empty text
    frame, body = the remaining paragraphs as lines, plus the first raster
    picture as `image_bytes` when the slide is image-type.

    Pure python-pptx — no LibreOffice needed — so an uploaded deck of any date
    can be turned into a video directly.
    """
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    out = []
    prs = Presentation(io.BytesIO(data))
    for slide in prs.slides:
        title, body, img = "", [], None
        texts = []
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and img is None:
                try:
                    img = shape.image.blob
                except Exception:
                    img = None
            if getattr(shape, "has_text_frame", False) and \
                    shape.text_frame.text.strip():
                texts.append(shape.text_frame.text)
        if texts:
            title = texts[0].strip()
            for tx in texts[1:]:
                body.extend(ln.strip() for ln in tx.splitlines()
                            if ln.strip())
        out.append({
            "title": title, "body": body, "image": None,
            "image_name": "", "image_bytes": img,
        })
    return out


def build_announcements_mp4(announcements, indices, secs_per_slide=5.0,
                            fps=15.0, width=1920, images=None,
                            progress=None):
    """Encode selected announcement slides into an MP4 (bytes).

    indices   : list of announcement positions to include (subset OK)
    images    : {index: bytes} for image-type announcements (else text slide)
    secs_per_slide / fps : pacing of the video
    """
    count_per = max(1, round(float(secs_per_slide) * float(fps)))
    frames = []
    for i in indices:
        a = announcements[i]
        frame = _announcement_frame(
            a, width,
            (images or {}).get(i, a.get("image_bytes")))
        arr = np.asarray(frame)
        frames.extend([arr] * count_per)
    return _encode_mp4(frames, fps, progress)