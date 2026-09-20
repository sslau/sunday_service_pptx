# model.py — week data model: defaults, parsing, normalization.
"""Weekly config shape (the dict generate_deck.build/plan consume):

    {
      "date": "2026.09.20",
      "psalm":        {ref, ref_size, font_size, verses[]},
      "hymn_font_size": 44,
      "hymn_font_min":  36,
      "hymn_margin_in": 0.83,
      "hymns": [ {title, subtitle, source, refrain[], refrain_after_every_verse,
                  verses[ [line...], ... ]} ],
      "scripture":    {ref, ref_size, font_size, max_lines, verses[]},
      "sermon":       {title, title_2, scripture, ref_size, speaker, date, slides[]},
      "response":     {title, subtitle, source, refrain[], refrain_after_every_verse,
                       verses[ [line...], ... ]},   # 詩歌回應（單首）
      "announcements":[ {title, body[], image, image_name} ],
      "communion":    false,
    }

The generator falls back to repo media/ images for header slots, so the week
dict need not carry any image paths.  Announcement "image" entries are
session-only (the uploaded file bytes are not persisted); image_name is kept
so the editor remembers which announcement is picture-based.
"""
import copy
import json
import os
import re

DATE_RE = re.compile(r"^\d{4}\.\d{2}\.\d{2}$")

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_FILE = os.path.join(HERE, "sample_week.json")
_default = None


def default_week(date):
    """A fresh week for `date`, seeded from sample_week.json."""
    global _default
    if _default is None:
        with open(DEFAULT_FILE, "r", encoding="utf-8") as fh:
            _default = json.load(fh)
    week = copy.deepcopy(_default)
    week["date"] = date
    week.setdefault("sermon", {})
    week.setdefault("response", {})
    week.setdefault("announcements", [])
    week.setdefault("communion", False)
    return week


def split_lines(text):
    """Text area -> list of non-empty, right-trimmed lines."""
    out = []
    for line in (text or "").split("\n"):
        line = line.rstrip()
        if line.strip():
            out.append(line)
    return out


def _num(value, default):
    try:
        n = int(round(float(value)))
        return n if n > 0 else default
    except (TypeError, ValueError):
        return default


def normalize_week(w):
    """Coerce an arbitrary dict into a safe week shape."""
    w = dict(w or {})
    w["date"] = str(w.get("date") or "").strip()
    psalm = dict(w.get("psalm") or {})
    psalm["ref"] = str(psalm.get("ref") or "")
    psalm["ref_size"] = _num(psalm.get("ref_size"), 48)
    psalm["font_size"] = _num(psalm.get("font_size"), 44)
    psalm["verses"] = [str(v) for v in (psalm.get("verses") or []) if str(v).strip()]
    w["psalm"] = psalm

    w["hymn_font_size"] = _num(w.get("hymn_font_size"), 44)
    w["hymn_font_min"] = _num(w.get("hymn_font_min"), 36)
    w["hymn_margin_in"] = float(w.get("hymn_margin_in") or 0.83)

    hymns = []
    for h in (w.get("hymns") or []):
        h = dict(h or {})
        hymn = {
            "title": str(h.get("title") or ""),
            "subtitle": str(h.get("subtitle") or ""),
            "source": str(h.get("source") or ""),
            "refrain": [str(x) for x in (h.get("refrain") or []) if str(x).strip()] or None,
            "refrain_after_every_verse": bool(h.get("refrain_after_every_verse", True)),
            "verses": [[str(x) for x in stanza if str(x).strip()]
                       for stanza in (h.get("verses") or [])],
        }
        hymn["verses"] = [s for s in hymn["verses"] if s]
        hymns.append(hymn)
    w["hymns"] = hymns

    scripture = dict(w.get("scripture") or {})
    scripture["ref"] = str(scripture.get("ref") or "")
    scripture["ref_size"] = _num(scripture.get("ref_size"), 48)
    scripture["font_size"] = _num(scripture.get("font_size"), 44)
    scripture["max_lines"] = _num(scripture.get("max_lines"), 4)
    scripture["verses"] = [str(v) for v in (scripture.get("verses") or []) if str(v).strip()]
    w["scripture"] = scripture

    sermon = dict(w.get("sermon") or {})
    for key, default in (("title", ""), ("title_2", ""), ("scripture", ""),
                         ("speaker", ""), ("date", w.get("date", ""))):
        sermon[key] = str(sermon.get(key) or default)
    sermon["ref_size"] = _num(sermon.get("ref_size"), 28)
    slides = []
    for s in (sermon.get("slides") or []):
        s = dict(s or {})
        slides.append({
            "title": str(s.get("title") or ""),
            "body": [str(x) for x in (s.get("body") or []) if str(x).strip()],
        })
    sermon["slides"] = slides
    w["sermon"] = sermon

    response = dict(w.get("response") or {})
    response["title"] = str(response.get("title") or "")
    response["subtitle"] = str(response.get("subtitle") or "")
    response["source"] = str(response.get("source") or "")
    response["refrain"] = [str(x) for x in (response.get("refrain") or [])
                           if str(x).strip()] or None
    response["refrain_after_every_verse"] = bool(
        response.get("refrain_after_every_verse", True))
    response["verses"] = [[str(x) for x in stanza if str(x).strip()]
                          for stanza in (response.get("verses") or [])]
    response["verses"] = [s for s in response["verses"] if s]
    w["response"] = response

    announcements = []
    for a in (w.get("announcements") or []):
        a = dict(a or {})
        announcements.append({
            "title": str(a.get("title") or ""),
            "body": [str(x) for x in (a.get("body") or []) if str(x).strip()],
            "image": a.get("image"),          # path/None — session-only
            "image_name": str(a.get("image_name") or ""),
        })
    w["announcements"] = announcements
    w["communion"] = bool(w.get("communion"))
    return w


def startswith_key_in(state, prefix):
    """True if any session key begins with `prefix`."""
    return any(k.startswith(prefix) for k in state)