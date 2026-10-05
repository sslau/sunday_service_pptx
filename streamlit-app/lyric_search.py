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
``app._parse_hymn_paste()`` turns into one 段 per slide.
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

# .txt stems carry a "[2]" style marker on repeat extractions; drop it so the
# song reads as 愛何等大 rather than 愛何等大[1].
_DUP_SUFFIX_RE = re.compile(r"\[\d+\]$")


def norm(s):
    """Casefolded text with whitespace and punctuation removed."""
    return STRIP_RE.sub("", s or "").lower()


def display_name(stem):
    """Song name for display: the filename without a ``[n]`` duplicate marker."""
    return _DUP_SUFFIX_RE.sub("", stem)


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

    def segments(self, key, drop_repeats=False, skip_first=False):
        """[(slide_no, text)] of one song, ready to become 段.

        `key` is the unique id from ``search()``; a bare song name is accepted
        too and resolves to the first song with that name.

        `drop_repeats` removes slides whose text repeats one already seen, since
        the dumps come from song decks where the chorus is deliberately
        re-slid.  `skip_first` treats slide 1 as a title card; it is opt-in
        because for some songs slide 1 is a real verse.
        """
        segs = list(self.songs.get(key)
                    or self.songs.get(self._find(key) or "")
                    or [])
        if skip_first:
            segs = segs[1:]
        if drop_repeats:
            seen, uniq = set(), []
            for no, text, ntext in segs:
                if ntext and ntext in seen:
                    continue
                seen.add(ntext)
                uniq.append((no, text, ntext))
            segs = uniq
        return [(no, text) for no, text, _ in segs]

    def _find(self, song):
        """Resolve a plain song name (handles the ``[n]`` marker) to a key."""
        for key, (name, _where) in self.meta.items():
            if name == song:
                return key
        return None

    def paste_text(self, key, drop_repeats=False, skip_first=False):
        """(text, n_segments) with one blank line between segments — the
        protocol ``app._parse_hymn_paste()`` splits into 段."""
        segs = self.segments(key, drop_repeats, skip_first)
        return "\n\n".join(t for _, t in segs), len(segs)


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
        songs[key] = [(no, text, norm(text)) for no, text in segs]
        meta[key] = (name, coll)
    _LIBRARY = LyricLibrary(LYRICS_ROOT, songs, meta)
    _STAMP = stamp
    return _LIBRARY