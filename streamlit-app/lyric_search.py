"""Lyric search over the extracted text library in ``lyrics/*/.lyrics_txt``.

Ported from the standalone ``lyric_search.py`` (Tk) tool.  Its extraction step
is not needed here because the ``.txt`` dumps it produced are already in the
repo, so this module only loads and queries them; ``app.py`` draws the UI.

``lyrics/`` holds one directory per collection (教會聖詩, 粵語其他詩歌, 美樂頌,
輔助本), each with a ``.lyrics_txt`` cache directory of dumps.  Subdirectories
inside a cache directory are searched too, so ``美樂頌`` works whether its dumps
sit at the top or inside a book subdirectory like ``Hymns Rejoice``.

Songs are keyed by their path rather than their filename, because the same
hymn appears in several books under the same name (122 敬拜主 is in both
粵語其他詩歌 and 輔助本).  Keying by filename made the later collection silently
overwrite the earlier one and lose the song.

Each ``.txt`` file is one song, split into ``=== Slide N ===`` blocks holding
that slide's text.  Matching ignores whitespace and punctuation, so a sentence
broken across slides or lines still matches, and several space-separated words
must all appear on the same slide (or in the song name).

Blocks are handed to the editor through ``paste_text()``, which joins a song's
slides with one blank line between them — the protocol
``app._parse_hymn_paste()`` turns into one 段 per slide.  ``highlight()`` renders
one slide as HTML with the query terms marked, for the reader pop-up.
"""
import collections
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
LYRICS_ROOT = os.path.join(HERE, "lyrics")
CACHE_DIRNAME = ".lyrics_txt"

SLIDE_RE = re.compile(r"^=== Slide (\d+) ===$", re.M)
# Page numbers ("3/12") are chrome, not lyrics.
PAGE_RE = re.compile(r"^\d{1,3}\s*/\s*\d{1,3}$")
# Whitespace + ASCII/CJK punctuation, dropped before comparing.  Python's \s
# also covers \x0b, the soft line break PowerPoint puts inside a paragraph.
STRIP_RE = re.compile(
    r"[\s\u3000.,;:!?'\"()\[\]{}<>、。，；：！？「」『』（）《》〈〉—…·\-_/\\|*&^%$#@~`+=]+")
# Leftover control characters once \x0b has become a newline.
CTRL_RE = re.compile(r"[\x00-\x08\x0c-\x1f]")

# search() cap; the reference tool used 500, which floods a Streamlit selectbox
MAX_HITS = 300

# key = unique path id (lookup), name = song name (填曲名), where = collection
Slide = collections.namedtuple("Slide", "key name where no text")
# One search result per song: `nos` is the slide numbers where every query term
# appears (empty when only the song name matched).
SongHit = collections.namedtuple("SongHit", "key name where nos")

# .txt stems carry a "[2]" style marker on repeat extractions; drop it so the
# song reads as 愛何等大 rather than 愛何等大[1].
_DUP_SUFFIX_RE = re.compile(r"\[\d+\]$")
# A title card carries the hymn number in its book ("聖詩 14", "輔助本 3",
# "Hymn 488-1", "美樂頌 62").
_BOOK_RE = re.compile(r"(聖詩|輔助本|美樂頌|Hymn|Sup)\s*\d", re.I)
# Leading book number / trailing edition letter on a display name.
_LEAD_NO_RE = re.compile(r"^\s*\d+\s*")
_TRAIL_ED_RE = re.compile(r"[（(][A-Za-z0-9][)）]\s*$")

# Title-card credits: "曲：X", "詞：Y", "詞曲：Z", "粵語譯詞：W", "Music By: V".
# Longest label first so 曲/詞 do not win over 詞曲/原曲詞.
_CREDIT_LABELS = (
    "粵語譯詞", "中文翻譯", "中文歌詞", "原曲詞", "詞曲", "曲詞", "曲/詞", "曲／詞",
    "詞/曲", "詞／曲", "作曲", "作詞", "粵詞", "譯詞", "中譯", "Music By",
    "Words by", "Words", "Composer", "Lyricist", "Lyrics", "Music", "曲", "詞", "譯",
)
_CREDIT_RE = re.compile(
    r"(?P<lab>%s)\s*[:：]"
    % "|".join(re.escape(x) for x in sorted(_CREDIT_LABELS, key=len, reverse=True)),
    re.I)
_CJK_RE = re.compile(r"[\u3400-\u9fff]")
_LATIN_RE = re.compile(r"[A-Za-z]")
# Lines that never name the album: bracketed markers, codes, copyright, junk.
_NOISE_RE = re.compile(r"^[（(].*[)）]$|ccli|copyright|版權|©|^\d+$|^[#\W]+$",
                       re.I)


def norm(s):
    """Casefolded text with whitespace and punctuation removed."""
    return STRIP_RE.sub("", s or "").lower()


def display_name(stem):
    """Song name for display: the filename without a ``[n]`` duplicate marker."""
    return _DUP_SUFFIX_RE.sub("", stem)


def _variant(s):
    """Fold 祢/祂 to 你/他 so a title card matches the name spelled either way."""
    return (s or "").replace("祢", "你").replace("祂", "他")


def _lines(text):
    return [x for x in (text or "").split("\n") if x.strip()]


def _line_set(text):
    return {norm(x) for x in _lines(text)}


def _song_title(name):
    """The hymn title without its book number or a trailing "(b)".

    Some collections name the dump after the book number on both sides
    (``88每一天88``); a matching pair is stripped so the title can be matched on
    its own.
    """
    s = (name or "").strip()
    m = re.match(r"^\s*(\d+)\s*(.*?)\s*(\d+)\s*$", s)
    if m and m.group(1) == m.group(3):
        s = m.group(2)
    return _TRAIL_ED_RE.sub("", _LEAD_NO_RE.sub("", s)).strip()


def _is_title_card(pages, name):
    """True when slide 1 is a title card (the title and/or a book marker alone).

    A title card is normally just a few lines (title, book marker, credits).
    Some decks (美樂頌, 輔助本) also bake a running header onto every page —
    slide 1 included — so the card can be long.  There the book marker is the
    tell, provided the next page does not carry it too: if it does, that is a
    running header printed on every page, not a card.
    """
    if not pages:
        return False
    text = pages[0][1]
    has_book = bool(_BOOK_RE.search(text))
    title = norm(_variant(_song_title(name)))
    has_title = bool(title) and title in norm(_variant(text))
    if len(_line_set(text)) <= 6:
        return has_book or has_title
    return (has_book and has_title
            and not (len(pages) > 1 and _BOOK_RE.search(pages[1][1])))


def _same_card(s, ref):
    """True when ``s`` repeats the title-card lines of ``ref`` and little else."""
    return len(s) <= len(ref) + 1 and len(s & ref) >= len(ref) - 1


def first_copy(pages, name):
    """``pages`` up to the second copy of the song, or all of them.

    These decks usually hold several editions ("copies") of one hymn, each
    introduced by the same title card, so importing the whole song used to drag
    every edition in.  Keep the first: cut at the first later page that repeats
    slide 1's title card.  A page that brings the title lines back *after a
    gap* starts a new copy; one that keeps repeating them without a gap is a
    running header (the title printed on every slide), which marks nothing, so
    those decks come back untouched.  When slide 1 is not a title card the deck
    is also left alone — guessing there would risk deleting real verses.
    """
    if len(pages) < 3 or not _is_title_card(pages, name):
        return pages
    ref = _line_set(pages[0][1])
    if len(ref) > 6:
        # Slide 1 is a title card sharing the page with a running header (see
        # _is_title_card), so most of ``ref`` sits on every page and the usual
        # overlap test is useless.  Only a page repeating slide 1 almost whole
        # — header plus title lines, no extra verse — marks the next copy.
        carries = [i for i, (_no, text) in enumerate(pages)
                   if _same_card(_line_set(text), ref)]
    else:
        carries = [i for i, (_no, text) in enumerate(pages)
                   if len(_line_set(text) & ref) >= 2]
    for i in carries:
        if i == 0 or (i - 1) in carries:        # still the same running header
            continue
        if _same_card(_line_set(pages[i][1]), ref):
            return pages[:i]
    return pages


def _credit_kind(label):
    """(composer?, lyricist?) named by one 曲/詞 style credit label."""
    music = "曲" in label or bool(re.search(r"music|composer|tune", label, re.I))
    lyric = ("詞" in label or "譯" in label
             or bool(re.search(r"words|lyric", label, re.I)))
    return music, lyric


def _title_match(cands, want, fallback=True):
    """The candidate line matching the song's display name, else the first.

    ``want`` is the normalised display name; the title card can put the album or
    publisher ahead of the title (``普天頌讚 / 高唱主名歌``), so the name picks
    the real title instead of blindly taking the first Chinese line.  With
    ``fallback=False`` nothing is returned when the name matches no line, so the
    caller can try the other script before settling for a guess.
    """
    if want:
        scored = []
        for cand in cands:
            n = norm(_variant(cand))
            if not n:
                continue
            if n == want:
                scored.append((2, len(n), cand))
            elif want in n or n in want:
                scored.append((1, min(len(n), len(want)), cand))
        if scored:
            return max(scored)[2]
        if not fallback:
            return ""
    return cands[0] if cands else ""


def parse_title_card(pages, name):
    """Fields read from a song's title card — ``{}`` when slide 1 is not one.

    Returns ``{"title", "subtitle", "source", "music", "lyricist"}`` (each a
    string, possibly empty) for the editor's 曲名/英文名/詩集/曲/詞 boxes.  ``name``
    is the display name, used to tell the title from the album when the card puts
    the album first.
    """
    if not pages or not _is_title_card(pages, name):
        return {}
    lines = _lines(pages[0][1])
    music, lyric, rest, i = [], [], [], 0
    while i < len(lines):
        line = lines[i]
        found = list(_CREDIT_RE.finditer(line))
        if not found:
            rest.append(line)
            i += 1
            continue
        for j, m in enumerate(found):
            end = found[j + 1].start() if j + 1 < len(found) else len(line)
            val = line[m.end():end].strip(" \t:：，,、/／.-")
            if not val:                 # "詞曲：" with names on the next line(s)
                grabbed, k = [], i + 1
                while (k < len(lines) and len(grabbed) < 2
                       and _LATIN_RE.search(lines[k])
                       and not _CREDIT_RE.search(lines[k])
                       and not _BOOK_RE.search(lines[k])
                       and not _CJK_RE.search(lines[k])):
                    grabbed.append(lines[k].strip())
                    lines[k] = ""
                    k += 1
                val = " ".join(grabbed)
            if val:
                is_music, is_lyric = _credit_kind(m.group("lab"))
                if is_music:
                    music.append(val)
                if is_lyric:
                    lyric.append(val)
        i += 1

    rest = [x for x in rest if x.strip()]
    want = norm(_variant(_song_title(name)))
    cjk = [x for x in rest if _CJK_RE.search(x) and not _BOOK_RE.search(x)]
    latin = [x for x in rest if _LATIN_RE.search(x) and not _CJK_RE.search(x)
             and not _BOOK_RE.search(x)]
    title = (_title_match(cjk, want, fallback=False)
             or _title_match(latin, want, fallback=False)
             or _title_match(cjk, want) or _title_match(latin, want))
    subtitle = next((x for x in latin if x != title), "")
    source = next((x for x in rest if _BOOK_RE.search(x)), "")
    if not source:
        for line in rest:
            m = re.search(r"(?:詩集|專輯|專輯名稱)\s*[:：]\s*(.+)", line)
            if m:
                source = m.group(1).strip()
                break
    if not source:
        for line in rest:
            if line in (title, subtitle) or _BOOK_RE.search(line):
                continue
            if _NOISE_RE.search(line):
                continue
            source = line
            break
    return {"title": title, "subtitle": subtitle, "source": source,
            "music": "；".join(music), "lyricist": "；".join(lyric)}


def tidy(text):
    """Clean one slide's extracted text.

    PowerPoint's soft line break becomes a real newline, control characters go,
    and blank lines / page numbers are dropped.  Dropping the blank lines
    matters: thousands of the library's slides contain one, and a blank line is
    the editor's段 separator, so leaving them in would split one slide into two.
    """
    out = []
    for line in (text or "").replace("\x0b", "\n").split("\n"):
        line = CTRL_RE.sub("", line).replace("\u3000", " ").strip()
        if line and not PAGE_RE.match(line):
            out.append(line)
    return "\n".join(out)


def snippet(text, query, limit=46):
    """One-line preview: the first line holding a query term, else the opening
    lines.  Unlike the reference tool this returns a whole line rather than a
    character window, because lyric lines are the meaningful unit."""
    terms = [norm(t) for t in (query or "").split() if norm(t)]
    lines = [x for x in (text or "").split("\n") if x.strip()]
    for line in lines:
        nline = norm(line)
        if any(t in nline for t in terms):
            line = line.strip()
            return line if len(line) <= limit else line[:limit] + "…"
    flat = " / ".join(x.strip() for x in lines)
    return flat if len(flat) <= limit * 2 else flat[:limit * 2] + "…"


def _html(s):
    """Escape `s` for inline HTML and keep its line breaks."""
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace("\n", "<br>"))


def _char_map(text):
    """(per-character normalised text, original index of each normalised char).

    ``highlight()`` needs to turn a match found in ``norm(text)`` back into a
    span of the *original* string, which is not a straight slice because
    matching drops whitespace and punctuation.  Mapping character by character
    gives that; the normed result is checked against ``norm(text)`` by the
    caller because a few characters casefold to more than one character.
    """
    chars, idx = [], []
    for i, ch in enumerate(text):
        n = STRIP_RE.sub("", ch).lower()
        if n:
            chars.append(n)
            idx.append(i)
    return "".join(chars), idx


def highlight(text, query):
    """`text` as HTML with every occurrence of each query term wrapped in
    ``<mark>``.  Falls back to plain escaped text when the match cannot be
    mapped back (or matched nothing), so it is always safe to render."""
    text = text or ""
    terms = sorted({norm(t) for t in (query or "").split() if norm(t)},
                   key=len, reverse=True)
    ntext = norm(text)
    if not terms or not ntext:
        return _html(text)
    mapped, idx = _char_map(text)
    if mapped != ntext:
        return _html(text)
    spans = []
    for t in terms:
        start = 0
        while True:
            j = ntext.find(t, start)
            if j < 0:
                break
            spans.append((idx[j], idx[j + len(t) - 1] + 1))
            start = j + 1
    if not spans:
        return _html(text)
    spans.sort()
    merged = [list(spans[0])]
    for a, b in spans[1:]:
        if a <= merged[-1][1]:          # overlapping or touching terms
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([a, b])
    out, prev = [], 0
    for a, b in merged:
        out.append(_html(text[prev:a]))
        out.append("<mark>%s</mark>" % _html(text[a:b]))
        prev = b
    out.append(_html(text[prev:]))
    return "".join(out)


def _parse(raw):
    """[(slide_no, text)] from one ``.txt`` dump, in slide order."""
    marks = list(SLIDE_RE.finditer(raw))
    out = []
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(raw)
        text = tidy(raw[m.end():end])
        if text:
            out.append((int(m.group(1)), text))
    return out


class LyricLibrary:
    """Every song in the ``.txt`` dumps, held in memory and searched by slide."""

    def __init__(self, root, songs, meta):
        self.root = root
        # key -> [(slide_no, text, normalised_text)], in slide order
        self.songs = songs
        # key -> (display_name, collection) — two keys can share a name
        self.meta = meta
        self.flat = []          # (key, slide_no, text, norm(text), norm(name))
        for key, segs in songs.items():
            nname = norm(meta[key][0])
            for no, text, ntext in segs:
                self.flat.append((key, no, text, ntext, nname))
        self.total_slides = sum(len(v) for v in songs.values())

    def __len__(self):
        return len(self.songs)

    def search(self, query, limit=MAX_HITS):
        """Slides containing every whitespace-separated term, or the term
        appearing in the song name.  Empty query -> no hits."""
        terms = [norm(t) for t in (query or "").split() if norm(t)]
        if not terms:
            return []
        hits = []
        for key, no, text, ntext, nname in self.flat:
            if all(t in ntext or t in nname for t in terms):
                name, where = self.meta[key]
                hits.append(Slide(key, name, where, no, text))
                if len(hits) >= limit:
                    break
        return hits

    def search_songs(self, query, limit=MAX_HITS):
        """One :class:`SongHit` per matching song, in library order.

        Same matching rule as ``search()`` — every whitespace-separated term —
        but grouped so a song whose chorus repeats on eight slides is listed
        once instead of eight times.  A song is included when all terms appear
        on some slide, or all terms appear in its name; ``nos`` holds the
        matching slide numbers (empty for a name-only match).
        """
        terms = [norm(t) for t in (query or "").split() if norm(t)]
        if not terms:
            return []
        out = []
        for key, (name, where) in self.meta.items():
            nname = norm(name)
            nos = [no for no, _text, ntext in self.songs[key]
                   if all(t in ntext for t in terms)]
            if nos or all(t in nname for t in terms):
                out.append(SongHit(key, name, where, nos))
                if len(out) >= limit:
                    break
        return out

    def segments(self, key):
        """[(slide_no, text)] of one song, ready to become 段 — every page.

        `key` is the unique id from ``search()``; a bare song name is accepted
        too and resolves to the first song with that name.

        One page becomes one 段, so the count the UI shows matches the page
        count and a source deck that re-slides its chorus after every verse
        keeps those pages as their own 段 rather than having them merged away.
        Extra editions of the hymn are already trimmed at load time by
        :func:`first_copy`, so "every page" here means every page of the first
        copy.
        """
        segs = list(self.songs.get(key)
                    or self.songs.get(self._find(key) or "")
                    or [])
        return [(no, text) for no, text, _ in segs]

    def _find(self, song):
        """Resolve a plain song name (handles the ``[n]`` marker) to a key."""
        for key, (name, _where) in self.meta.items():
            if name == song:
                return key
        return None

    def paste_text(self, key):
        """(text, n_segments) with one blank line between segments — the
        protocol ``app._parse_hymn_paste()`` splits into 段.  One page per 段."""
        segs = self.segments(key)
        return "\n\n".join(t for _, t in segs), len(segs)

    def import_song(self, key):
        """(info, text, n) prepared for a whole-song import.

        ``info`` is the title-card metadata read from slide 1 (empty when slide 1
        is not a title card).  A recognised title card is left out of the returned
        text, because the deck builds its own title slide from ``info``; importing
        it too would print two.  Every remaining page stays its own 段.
        """
        resolved = key if key in self.songs else (self._find(key) or "")
        segs = self.segments(resolved)
        name = self.meta.get(resolved, ("", ""))[0]
        info = parse_title_card(segs, name)
        if info:
            segs = segs[1:]
        return info, "\n\n".join(t for _no, t in segs), len(segs)


def _scan():
    """[(key, name, collection, path)] for every song dump, in a stable order.

    `key` is the path below ``lyrics/`` without the cache directory or the
    extension — unique even when two collections hold the same hymn under the
    same filename.
    """
    found = []
    if not os.path.isdir(LYRICS_ROOT):
        return found
    for coll in sorted(os.listdir(LYRICS_ROOT)):
        cache = os.path.join(LYRICS_ROOT, coll, CACHE_DIRNAME)
        if not os.path.isdir(cache):
            continue
        for dirpath, _dirs, names in os.walk(cache):
            # "" when the dumps sit directly in .lyrics_txt, else the book name
            book = os.path.relpath(dirpath, cache)
            book = "" if book == os.curdir else book.replace(os.sep, "/")
            for fname in names:
                if not fname.lower().endswith(".txt"):
                    continue
                stem = fname[:-4]
                key = "%s/%s" % (coll, stem) if not book else \
                      "%s/%s/%s" % (coll, book, stem)
                found.append((key, display_name(stem), coll,
                              os.path.join(dirpath, fname)))
    found.sort(key=lambda kv: (norm(kv[1]), kv[2], kv[0]))
    return found


_LIBRARY = None
_STAMP = None


def get_library():
    """The parsed library, cached until a dump file changes on disk.

    Parsing every dump costs a noticeable fraction of a second, which is too
    much to pay on each Streamlit rerun; the stamp makes edits show up without
    a restart while still costing only the directory walk.
    """
    global _LIBRARY, _STAMP
    files = _scan()
    stamp = tuple((key, os.path.getmtime(path)) for key, _n, _c, path in files)
    if _LIBRARY is not None and stamp == _STAMP:
        return _LIBRARY
    songs, meta = {}, {}
    for key, name, coll, path in files:
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as fh:
                segs = _parse(fh.read())
        except OSError:
            continue
        if not segs:
            continue        # a dump with no extractable text: nothing to search
        first = first_copy(segs, name)
        songs[key] = [(no, text, norm(text)) for no, text in first]
        meta[key] = (name, coll)
    _LIBRARY = LyricLibrary(LYRICS_ROOT, songs, meta)
    _STAMP = stamp
    return _LIBRARY