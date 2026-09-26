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
import json
import os
import re
import tempfile

from generate_deck import (_blank_layout, build, build_section_deck, plan)  # noqa: E402
from pptx import Presentation

_PPTX_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "blank_16x9.pptx")

# Per-section standalone decks saved by the web app are kept under
# data/decks/<date>/ so a later full build can reuse them automatically.
SAVE_DIR = os.path.join(HERE, "data", "decks")
SECTION_FILENAMES = {
    "songs": "songs_slides_{date}.pptx",
    "sermon": "sermon_{date}.pptx",
    "announcements": "announcements_{date}.pptx",
    "offering": "offering_{date}.pptx",
    "response": "response_{date}.pptx",
}
# section -> (import flag, file key) read by generate_deck.build
SECTION_CFG = {
    "songs": ("songs_import", "songs_file"),
    "sermon": ("sermon_import", "sermon_file"),
    "announcements": ("announcements_import", "announcements_file"),
    "offering": ("offering_import", "offering_file"),
    "response": ("response_import", "response_file"),
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


def clear_section_deck(section, date):
    """Delete a previously saved section deck (if any). Used to undo an
    accidental upload/persist. Returns True when a file was removed."""
    path = saved_section_path(section, date)
    if path:
        os.remove(path)
        return True
    return False


# Original uploaded/source filenames are kept in a small JSON next to the
# saved decks so the status page can still show them after a re-login
# (session state is cleared, but the deck files on disk survive).
def upload_names_path(date):
    if not date:
        return None
    return os.path.join(SAVE_DIR, str(date), "upload_names.json")


def save_upload_name(section, date, name):
    """Durably record the original filename a section deck was saved from."""
    path = upload_names_path(date)
    if not path or not name:
        return
    names = {}
    try:
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as fh:
                names = json.load(fh) or {}
    except (OSError, ValueError, TypeError):
        names = {}
    if not isinstance(names, dict):
        names = {}
    names[section] = str(name)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(names, fh, ensure_ascii=False, indent=1)


def load_upload_names(date):
    """Mapping section -> original filename for a date, or {}."""
    path = upload_names_path(date)
    if not path or not os.path.isfile(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            names = json.load(fh)
        return names if isinstance(names, dict) else {}
    except (OSError, ValueError, TypeError):
        return {}


def clear_upload_name(section, date):
    """Drop a section's recorded original filename (keeps the deck itself)."""
    names = load_upload_names(date)
    if section in names:
        del names[section]
        path = upload_names_path(date)
        try:
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(names, fh, ensure_ascii=False, indent=1)
        except OSError:
            pass


def video_mp4_path(date):
    """Canonical save path for a generated 家事MP4 (may not exist yet)."""
    if not date:
        return None
    return os.path.join(SAVE_DIR, str(date),
                        "announcements_{date}.mp4".format(date=date))


def saved_video_mp4(date):
    """Path of a previously generated 家事MP4, or None."""
    path = video_mp4_path(date)
    return path if path and os.path.isfile(path) else None


def save_video_mp4(date, data):
    """Persist a generated 家事MP4 so a later reload/build can reuse it
    (session_state alone is wiped on restart).  Returns the saved path."""
    path = video_mp4_path(date)
    if not path:
        return None
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(data)
    return path


def clear_video(date):
    """Delete a previously generated 家事MP4 and its 併入第一頁 meta
    (if any).  Returns True when the mp4 was removed."""
    removed = False
    path = saved_video_mp4(date)
    if path:
        try:
            os.remove(path)
            removed = True
        except OSError:
            pass
    meta = _video_meta_path(date)
    if meta and os.path.isfile(meta):
        try:
            os.remove(meta)
        except OSError:
            pass
    return removed


def _video_meta_path(date):
    if not date:
        return None
    return os.path.join(SAVE_DIR, str(date), "video_first.json")


def save_video_meta(date, first_page):
    """Persist the 併入第一頁 checkbox value next to the generated MP4."""
    path = _video_meta_path(date)
    if not path:
        return None
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump({"first_page": bool(first_page)}, fh)
    return path


def saved_video_meta(date):
    """Previously persisted 併入第一頁 value, or None if unknown."""
    path = _video_meta_path(date)
    if not path or not os.path.isfile(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return bool(json.load(fh).get("first_page"))
    except (OSError, ValueError, TypeError):
        return None

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
            image_overrides, announcement_images, use_saved=True,
            offering_pptx=None, response_pptx=None):
    cfg = copy.deepcopy(week)
    _set_pptx(cfg, "songs_import", "songs_file", songs_pptx)
    _set_pptx(cfg, "announcements_import", "announcements_file",
              announcements_pptx)
    _set_pptx(cfg, "sermon_import", "sermon_file", sermon_pptx)
    _set_pptx(cfg, "offering_import", "offering_file", offering_pptx)
    _set_pptx(cfg, "response_import", "response_file", response_pptx)
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
               announcement_images=None, use_saved=True, offering_pptx=None,
               response_pptx=None):
    """Build the full deck for `week`; returns (pptx_bytes, slide_count).
    songs/sermon/announcements/offering/response decks may be given as
    UploadedFile *or* a filename/hint string resolved via resolve_deck_path();
    each is merged in place of web-compiled content.  With `use_saved`,
    sections not given explicitly fall back to a deck saved earlier for this
    date (data/decks/<date>/)."""
    os.chdir(HERE)  # generator resolves media/ fallbacks relative to cwd
    cfg = _config(week, songs_pptx, announcements_pptx, sermon_pptx,
                  image_overrides, announcement_images, use_saved,
                  offering_pptx, response_pptx)
    pres = Presentation(TEMPLATE)
    build(pres, cfg)
    buf = io.BytesIO()
    pres.save(buf)
    return buf.getvalue(), plan(cfg)


def count_deck(week, songs_pptx=None, announcements_pptx=None,
               sermon_pptx=None, use_saved=True, offering_pptx=None,
               response_pptx=None):
    """Slide count the generator would produce (no build)."""
    os.chdir(HERE)
    cfg = _config(week, songs_pptx, announcements_pptx, sermon_pptx, {}, {},
                  use_saved, offering_pptx, response_pptx)
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


def build_hymn_pptx_bytes(cfg, hymn):
    """Compile a single worship hymn (標題卡 + that hymn's slides) into its own
    standalone .pptx for inline per-song preview.  `cfg` can be a full assembled
    week; only the hymn_* layout keys and date are used.  Returns
    (pptx_bytes, slide_count)."""
    from generate_deck import build_hymns_section, clear_slides
    os.chdir(HERE)
    pres = Presentation(TEMPLATE)
    clear_slides(pres)
    if hymn.get("verses") or hymn.get("title"):
        build_hymns_section(pres, cfg, [hymn])
    buf = io.BytesIO()
    pres.save(buf)
    return buf.getvalue(), len(pres.slides._sldIdLst)


def hymn_preview(data):
    """Read a compiled single-hymn pptx and summarise every slide for an
    inline text preview: [{lines: [(text, pt|None), ...], bg: bool}].  Text,
    font size and background come straight from the generated file, so it
    mirrors the real slides.  Page-number (N/M) flagged out."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    pres = Presentation(io.BytesIO(data))
    out = []
    for slide in pres.slides:
        lines, has_bg, seen = [], False, set()
        for sh in slide.shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                has_bg = True
                continue
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                txt = "".join(r.text for r in p.runs).strip()
                if not txt or re.match(r"^\d{1,3}/\d{1,3}$", txt):
                    continue
                size = p.runs[0].font.size
                pt = round(size.pt) if size else None
                if (txt, pt) not in seen:
                    seen.add((txt, pt))
                    lines.append((txt, pt))
        out.append({"lines": lines, "bg": has_bg})
    return out


def prepend_video_slide(pptx_data, mp4_bytes, poster_bytes):
    """Return a new compiled deck whose FIRST slide is the announcements MP4
    embedded full-bleed, set to play automatically when the slide appears and
    keep looping until the presenter clicks to the next slide.

    pptx_data : bytes of the compiled .pptx
    mp4_bytes : bytes of the MP4 to embed
    poster_bytes : bytes of a PNG poster (shown before playback / in editor)
    """
    with tempfile.NamedTemporaryFile(suffix=".mp4", delete=False) as tm:
        tm.write(mp4_bytes)
        mp4_path = tm.name
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tp:
        tp.write(poster_bytes)
        poster_path = tp.name
    try:
        pres = Presentation(io.BytesIO(pptx_data))
        slide = pres.slides.add_slide(_blank_layout(pres))
        for shp in list(slide.shapes):
            shp._element.getparent().remove(shp._element)
        # Poster frame laid beneath the video so the first slide always shows
        # the frame even in viewers that do not render embedded video.
        slide.shapes.add_picture(poster_path, 0, 0,
                                 pres.slide_width, pres.slide_height)
        gf = slide.shapes.add_movie(
            mp4_path, 0, 0, pres.slide_width, pres.slide_height,
            poster_frame_image=poster_path, mime_type="video/mp4")
        _set_media_autoplay_loop(slide._element, gf.shape_id)
        sldIdLst = pres.slides._sldIdLst
        last = sldIdLst[-1]
        sldIdLst.remove(last)
        sldIdLst.insert(0, last)
        out = io.BytesIO()
        pres.save(out)
        return out.getvalue()
    finally:
        os.unlink(mp4_path)
        os.unlink(poster_path)


def _set_media_autoplay_loop(slde, spid):
    """Flip the movie's playback options straight on the <p:timing> node that
    python-pptx already creates for add_movie(), mirroring exactly what
    PowerPoint writes for “Start: Automatically + Loop until stopped + Rewind
    after playing”.  Appending a second <p:timing> breaks playback, so we tune
    the existing one instead:
      - start <p:cond delay="0">            -> play on slide entry
      - <p:cTn repeatCount="indefinite">    -> loop until stopped
      - <p:cTn fill="remove">               -> rewind after playing
    """
    for video in slde.findall(".//p:timing//p:video",
                              {"p": _PPTX_NS}):
        tgt = video.find(".//p:tgtEl/p:spTgt", {"p": _PPTX_NS})
        if tgt is None or tgt.get("spid") != str(spid):
            continue
        ctn = video.find(".//p:cMediaNode/p:cTn", {"p": _PPTX_NS})
        if ctn is None:
            continue
        ctn.set("repeatCount", "indefinite")
        ctn.set("fill", "remove")
        cond = ctn.find(".//p:stCondLst/p:cond", {"p": _PPTX_NS})
        if cond is not None:
            cond.set("delay", "0")
        return