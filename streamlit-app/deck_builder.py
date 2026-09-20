# deck_builder.py — turn a week dict into a downloadable .pptx in memory.
"""Reuses the existing generator (generate_deck.build) as a library; no
subprocess, no local bridge.  Uploaded decks/images are written to temp files
and handed to the generator as absolute paths.

    from deck_builder import build_deck, count_deck

    png_bytes, n = build_deck(week, songs_pptx=<UploadedFile|None>,
                              announcements_pptx=..., image_overrides=...)
"""
import copy
import io
import os
import tempfile

from generate_deck import (build, build_section_deck, plan)
from pptx import Presentation

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "blank_16x9.pptx")

# Per-section standalone decks saved by the web app are kept under
# data/decks/<date>/ so a later full build can reuse them automatically.
SAVE_DIR = os.path.join(HERE, "data", "decks")
SECTION_FILENAMES = {
    "songs": "songs_slides_{date}.pptx",
    "sermon": "sermon_{date}.pptx",
    "announcements": "announcements_{date}.pptx",
}
# section -> (import flag, file key) read by generate_deck.build
SECTION_CFG = {
    "songs": ("songs_import", "songs_file"),
    "sermon": ("sermon_import", "sermon_file"),
    "announcements": ("announcements_import", "announcements_file"),
}


def section_path(section, date):
    """Canonical save path for a section deck (may not exist yet)."""
    tmpl = SECTION_FILENAMES.get(section)
    if not tmpl or not date:
        return None
    return os.path.join(SAVE_DIR, str(date), tmpl.format(date=date))


def saved_section_path(section, date):
    """Path of a previously saved section deck, or None."""
    path = section_path(section, date)
    return path if path and os.path.isfile(path) else None


def save_section_deck(section, date, data):
    """Persist a compiled section deck so a later full build can reuse it.
    Returns the saved path."""
    path = section_path(section, date)
    if not path:
        return None
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(data)
    return path

# Header/image slots the generator resolves from the config (keys are the
# exact names generate_deck.build reads).  Exposed as optional build-time
# uploads; unset keys fall back to repo media/ defaults.
IMAGE_SLOTS = [
    "welcome_image", "call_image", "worship_header", "prayer_header",
    "lords_prayer", "scripture_header", "sermon_header", "doxology",
    "benediction", "announce_header", "closing", "communion", "apostles_creed",
]


def _write_upload(upload, suffix):
    path = tempfile.NamedTemporaryFile(
        delete=False, suffix=suffix or ".tmp", prefix="svc_").name
    with open(path, "wb") as fh:
        fh.write(upload.getvalue())
    return path


def _ext(upload):
    if upload is None:
        return ""
    _, ext = os.path.splitext(upload.name or "")
    return ext if ext else ".bin"


def resolve_deck_path(name):
    """Resolve a deck filename/hint against the app folder, data/downloads,
    and ~/Downloads (handy when running locally next to the generator)."""
    if not name or not str(name).strip():
        return None
    name = str(name).strip()
    if os.path.isabs(name) and os.path.isfile(name):
        return name
    dirs = [HERE, os.path.join(HERE, "data", "downloads"),
            os.path.expanduser("~/Downloads")]
    candidates = [name, os.path.basename(name)]
    for cand in candidates:
        for d in dirs:
            path = os.path.normpath(os.path.join(d, cand))
            if os.path.isfile(path):
                return path
    return None


def _resolve_upload(upload):
    """Accept an UploadedFile, or a filename/hint string looked up on disk."""
    if upload is None:
        return None
    if isinstance(upload, str):
        return resolve_deck_path(upload)
    return upload


def _set_pptx(cfg, flag_key, file_key, upload):
    up = _resolve_upload(upload)
    if up is None:
        return
    cfg[flag_key] = True
    if isinstance(up, str):
        cfg[file_key] = up
    else:
        cfg[file_key] = _write_upload(up, _ext(up))


def _config(week, songs_pptx, announcements_pptx, sermon_pptx,
            image_overrides, announcement_images, use_saved=True):
    cfg = copy.deepcopy(week)
    _set_pptx(cfg, "songs_import", "songs_file", songs_pptx)
    _set_pptx(cfg, "announcements_import", "announcements_file",
              announcements_pptx)
    _set_pptx(cfg, "sermon_import", "sermon_file", sermon_pptx)
    if use_saved:
        date = cfg.get("date")
        for section, (flag, fkey) in SECTION_CFG.items():
            if not cfg.get(flag):
                saved = saved_section_path(section, date)
                if saved:
                    cfg[flag] = True
                    cfg[fkey] = saved
    for key, upload in (image_overrides or {}).items():
        if upload is not None:
            cfg[key] = _write_upload(upload, _ext(upload))
    for idx, upload in (announcement_images or {}).items():
        if upload is not None and idx < len(cfg.get("announcements", [])):
            cfg["announcements"][idx]["image"] = _write_upload(upload,
                                                               _ext(upload))
    return cfg


def build_deck(week, songs_pptx=None, announcements_pptx=None,
               sermon_pptx=None, image_overrides=None,
               announcement_images=None, use_saved=True):
    """Build the full deck for `week`; returns (pptx_bytes, slide_count).
    Any of the three section decks (songs/sermon/announcements) may be given
    as an UploadedFile *or* a filename/hint string resolved via
    resolve_deck_path(); each is merged in place of web-compiled content.
    With `use_saved`, sections not given explicitly fall back to a deck saved
    earlier for this date (data/decks/<date>/)."""
    os.chdir(HERE)  # generator resolves media/ fallbacks relative to cwd
    cfg = _config(week, songs_pptx, announcements_pptx, sermon_pptx,
                  image_overrides, announcement_images, use_saved)
    pres = Presentation(TEMPLATE)
    build(pres, cfg)
    buf = io.BytesIO()
    pres.save(buf)
    return buf.getvalue(), plan(cfg)


def count_deck(week, songs_pptx=None, announcements_pptx=None,
               sermon_pptx=None, use_saved=True):
    """Slide count the generator would produce (no build)."""
    os.chdir(HERE)
    cfg = _config(week, songs_pptx, announcements_pptx, sermon_pptx, {}, {},
                  use_saved)
    return plan(cfg)


def build_section_pptx_bytes(section, week, announcement_images=None):
    """Compile one section (hymns|sermon|announcements|scripture) as a
    standalone .pptx from the web-editor content, so it can be downloaded and
    shared / merged.  ``scripture`` yields verse slides only (讀經), a separate
    workflow.  Returns (pptx_bytes, slide_count)."""
    os.chdir(HERE)
    cfg = _config(week, None, None, None, {}, announcement_images,
                  use_saved=False)
    pres = build_section_deck(section, cfg, TEMPLATE)
    buf = io.BytesIO()
    pres.save(buf)
    return buf.getvalue(), len(pres.slides._sldIdLst)