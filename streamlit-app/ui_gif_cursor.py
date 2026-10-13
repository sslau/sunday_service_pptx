"""Draw a click indicator (hand + ripple) onto UI screenshots for the GIFs.

The tutorial GIFs call `draw_click()` on a screenshot to show *where* the demo
just clicked: a colour hand whose fingertip sits on the click point, plus a red
ripple ring.  macOS ships Apple Color Emoji, but Pillow only accepts its bitmap
strike sizes, so the requested size is snapped to a supported one.
"""
from PIL import Image, ImageDraw, ImageFont

_EMOJI_FONT = "/System/Library/Fonts/Apple Color Emoji.ttc"
_FALLBACK_FONT = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"
_EMOJI_SIZES = (160, 96, 64, 48, 40, 32, 20)
_RIPPLE = (255, 59, 48)
_CACHE = {}


def _snap(size):
    return min(_EMOJI_SIZES, key=lambda s: abs(s - size))


def hand(size=96, char="👆"):
    """An RGBA pointing-hand image (fingertip at the top-centre)."""
    size = _snap(size)
    key = (char, size)
    if key in _CACHE:
        return _CACHE[key]
    pad = size // 2
    img = None
    try:                                   # colour emoji
        font = ImageFont.truetype(_EMOJI_FONT, size)
        canvas = Image.new("RGBA", (size + 2 * pad, size + 2 * pad), (0, 0, 0, 0))
        ImageDraw.Draw(canvas).text((pad, pad), char, font=font,
                                    embedded_color=True)
        bb = canvas.getbbox()
        if bb:
            img = canvas.crop(bb)
    except Exception:
        img = None
    if img is None:                        # monochrome fallback + white halo
        try:
            font = ImageFont.truetype(_FALLBACK_FONT, size)
            ch = "☝"                        # U+261D white up pointing index
            canvas = Image.new("RGBA", (size + 2 * pad, size + 2 * pad), (0, 0, 0, 0))
            d = ImageDraw.Draw(canvas)
            for dx, dy in ((-2, 0), (2, 0), (0, -2), (0, 2),
                           (-2, -2), (2, 2), (-2, 2), (2, -2)):
                d.text((pad + dx, pad + dy), ch, font=font,
                       fill=(255, 255, 255, 255))
            d.text((pad, pad), ch, font=font, fill=(20, 20, 20, 255))
            bb = canvas.getbbox()
            if bb:
                img = canvas.crop(bb)
        except Exception:
            img = None
    if img is None:
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    _CACHE[key] = img
    return img


def draw_click(img, x, y, scale=1.0, char="👆"):
    """Return `img` with a hand+ripple whose fingertip is at (x, y).

    `scale` is the screenshot's device-pixel ratio (streamlit GIFs use 2)."""
    x, y = int(round(x)), int(round(y))
    out = img.convert("RGBA")
    r = int(round(24 * scale))
    ring = Image.new("RGBA", out.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ring)
    d.ellipse([x - r, y - r, x + r, y + r], fill=_RIPPLE + (60,))
    d.ellipse([x - r, y - r, x + r, y + r], outline=_RIPPLE + (235,),
              width=max(2, int(round(3 * scale))))
    out.alpha_composite(ring)
    h = hand(int(round(48 * scale)), char)
    out.alpha_composite(h, (x - h.width // 2, y))   # fingertip on the point
    return out.convert("RGB")


def marker_colours(size=96, n=24):
    """Palette entries (RGB tuples) that keep the click marker alive in a GIF.

    A GIF is limited to 256 colours and median-cut allocates them by area, so a
    small overlay like the hand is otherwise quantized away.  Reserving these
    colours for the frame palette keeps it visible."""
    h = hand(size)
    canvas = Image.new("RGB", h.size, (255, 255, 255))
    canvas.paste(h, (0, 0), h)
    q = canvas.quantize(colors=n, method=Image.MEDIANCUT)
    pal = q.getpalette()[: n * 3]
    rgbs = [_RIPPLE] + [(pal[i], pal[i + 1], pal[i + 2])
                        for i in range(0, n * 3, 3)]
    seen, out = set(), []
    for c in rgbs:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out


def quantize_with_markers(img, colours=256, size=96):
    """Quantize `img` to `colours`, reserving room for the click marker."""
    extra = marker_colours(size)
    k = len(extra)
    base = img.quantize(colors=max(2, colours - k), method=Image.MEDIANCUT)
    pal = list(base.getpalette()[: (colours - k) * 3])
    for rgb in extra:
        pal.extend(rgb)
    pimg = Image.new("P", (1, 1))
    pimg.putpalette(pal)
    return img.quantize(palette=pimg, dither=Image.NONE)
