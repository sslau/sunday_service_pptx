# bible.py — fetch scripture text from the FHL (信望愛) Bible site.
"""Endpoint (verified against bible.fhl.net):

    https://bible.fhl.net/json/qb.php
        ?chineses=<book-code>&chap=<n>&sec=<v>&version=<ver>

`chineses` must be FHL's Chinese abbreviation (full names like 詩篇 return no
records).  Omitting `sec` returns the whole chapter; there is no `sec2`, so a
verse range is filtered here.  `version` defaults to `unv` (和合本).

Example:
    >>> text, info = fetch_reference("詩篇 100:1-2")
    >>> print(text)
    1 （稱謝詩。）普天下當向耶和華歡呼！
    2 你們當樂意事奉耶和華，當來向他歌唱！
"""
import json
import os
import re
import urllib.parse
import urllib.request

API = "https://bible.fhl.net/json/qb.php"
DEFAULT_VERSION = "unv"
USER_AGENT = "Mozilla/5.0 (compatible; SundayServiceSlides/1.0)"

# Book name (full or abbreviation) -> FHL `chineses` code.  All 66 codes below
# were verified to return records from the API.
_BOOKS = {
    "創世記": "創", "創世紀": "創", "出埃及記": "出", "利未記": "利",
    "民數記": "民", "申命記": "申", "約書亞記": "書", "士師記": "士",
    "路得記": "得", "撒母耳記上": "撒上", "撒母耳記下": "撒下",
    "列王紀上": "王上", "列王紀下": "王下", "歷代志上": "代上",
    "歷代志下": "代下", "以斯拉記": "拉", "尼希米記": "尼", "以斯帖記": "斯",
    "約伯記": "伯", "詩篇": "詩", "箴言": "箴", "傳道書": "傳", "雅歌": "歌",
    "以賽亞書": "賽", "耶利米書": "耶", "耶利米哀歌": "哀", "以西結書": "結",
    "但以理書": "但", "何西阿書": "何", "約珥書": "珥", "阿摩司書": "摩",
    "俄巴底亞書": "俄", "約拿書": "拿", "彌迦書": "彌", "那鴻書": "鴻",
    "哈巴谷書": "哈", "西番雅書": "番", "哈該書": "該", "撒迦利亞書": "亞",
    "瑪拉基書": "瑪", "馬太福音": "太", "馬可福音": "可", "路加福音": "路",
    "約翰福音": "約", "使徒行傳": "徒", "羅馬書": "羅", "哥林多前書": "林前",
    "哥林多後書": "林後", "加拉太書": "加", "以弗所書": "弗", "腓立比書": "腓",
    "歌羅西書": "西", "帖撒羅尼迦前書": "帖前", "帖撒羅尼迦後書": "帖後",
    "提摩太前書": "提前", "提摩太後書": "提後", "提多書": "多",
    "腓利門書": "門", "希伯來書": "來", "雅各書": "雅", "彼得前書": "彼前",
    "彼得後書": "彼後", "約翰一書": "約一", "約翰二書": "約二",
    "約翰三書": "約三", "猶大書": "猶", "啟示錄": "啟",
}
# Accept the FHL abbreviations themselves as input too.
_BOOKS.update({v: v for v in set(_BOOKS.values())})

# Canonical full names in biblical order (for the book dropdown).
BOOK_NAMES = [
    "創世記", "出埃及記", "利未記", "民數記", "申命記", "約書亞記", "士師記",
    "路得記", "撒母耳記上", "撒母耳記下", "列王紀上", "列王紀下", "歷代志上",
    "歷代志下", "以斯拉記", "尼希米記", "以斯帖記", "約伯記", "詩篇", "箴言",
    "傳道書", "雅歌", "以賽亞書", "耶利米書", "耶利米哀歌", "以西結書",
    "但以理書", "何西阿書", "約珥書", "阿摩司書", "俄巴底亞書", "約拿書",
    "彌迦書", "那鴻書", "哈巴谷書", "西番雅書", "哈該書", "撒迦利亞書",
    "瑪拉基書", "馬太福音", "馬可福音", "路加福音", "約翰福音", "使徒行傳",
    "羅馬書", "哥林多前書", "哥林多後書", "加拉太書", "以弗所書", "腓立比書",
    "歌羅西書", "帖撒羅尼迦前書", "帖撒羅尼迦後書", "提摩太前書",
    "提摩太後書", "提多書", "腓利門書", "希伯來書", "雅各書", "彼得前書",
    "彼得後書", "約翰一書", "約翰二書", "約翰三書", "猶大書", "啟示錄",
]
# The book list lives in the static file `bible_books.json` (loaded at startup)
# so it can be maintained without touching code; built-ins are the fallback.
_BOOK_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "bible_books.json")
_BOOK_ALIASES = {"創世紀": "創"}


def _load_books():
    try:
        with open(_BOOK_FILE, encoding="utf-8") as fh:
            data = json.load(fh)
        names = [d.get("name") for d in data if d.get("name") and d.get("fhl")]
        if not names:
            raise ValueError("empty book list")
        return names, _BOOK_ALIASES | {d["name"]: d["fhl"] for d in data}
    except (OSError, ValueError, KeyError, TypeError):
        return BOOK_NAMES, _BOOKS


BOOK_NAMES, _loaded = _load_books()
_loaded.update({v: v for v in set(_loaded.values())})  # accept codes too
_BOOKS = _loaded
_NAMES_BY_LEN = sorted(_BOOKS, key=len, reverse=True)


def split_book(ref):
    """Split a reference into (leading_book_name|None, remainder)."""
    s = str(ref).strip()
    for name in _NAMES_BY_LEN:
        if s.startswith(name):
            return name, s[len(name):].strip()
    return None, s

# Translation label -> FHL `version` code.
_VERSIONS = {
    "和合本": "unv", "fhl和合本": "unv",
    "和合本2010": "rcuv", "和合本2010版": "rcuv",
    "現代中文譯本": "tcv95", "現代中文譯本1995版": "tcv95",
    "現代中文譯本2019版": "tcv2019",
    "環球譯本": "wcb", "新譯本": "ncv", "中文標準譯本": "csb",
    "恢復本": "recover", "思高譯本": "ofm", "呂振中譯本": "lcc",
    "net聖經中譯本": "cnet", "深文理和合本": "wlunv",
}
_VERSION_LABELS = {
    "unv": "和合本", "rcuv": "和合本2010", "tcv95": "現代中文譯本1995版",
    "tcv2019": "現代中文譯本2019版", "wcb": "環球譯本", "ncv": "新譯本",
    "csb": "中文標準譯本", "recover": "恢復本", "ofm": "思高譯本",
    "lcc": "呂振中譯本", "cnet": "NET聖經中譯本", "wlunv": "深文理和合本",
}


def _norm(text):
    return re.sub(r"\s+", "", str(text)).lower()


def version_code(text):
    return _VERSIONS.get(_norm(text))


def version_label(code):
    return _VERSION_LABELS.get(code, code)


def book_code(name):
    return _BOOKS.get(str(name).strip())


def parse_reference(ref):
    """Parse e.g. '羅馬書 12:1-8（和合本）' into (book_code, version, specs).

    `specs` is a list of (chapter, start_verse|None, end_verse|None); None means
    the start/end of that chapter.  Raises ValueError on anything unparseable.
    """
    s = str(ref).strip()
    version = DEFAULT_VERSION
    note = re.search(r"[（(]([^）)]+)[）)]", s)
    if note:
        version = version_code(note.group(1)) or version
        s = (s[:note.start()] + s[note.end():]).strip()

    m = re.match(r"^\s*([^\d]+?)\s*(\d.*)$", s)
    if not m:
        raise ValueError("格式應如「詩篇 34:1-3」")
    code = book_code(m.group(1))
    if not code:
        raise ValueError("找不到書卷「%s」" % m.group(1).strip())

    rest = m.group(2)
    rest = rest.replace("：", ":").replace("，", ",").replace(",", " ")
    rest = re.sub(r"[～~－–—]", "-", rest)
    specs = []
    for seg in rest.split():
        specs.extend(_parse_segment(seg))
    if not specs:
        raise ValueError("找不到經文範圍")
    return code, version, specs


def _parse_segment(seg):
    # chapter[:verse][-(chapter:verse | end_verse)]
    m = re.match(r"^(\d+)(?::(\d+))?(?:-(?:(\d+):(\d+)|(\d+)))?$", seg)
    if not m:
        raise ValueError("無法解析「%s」" % seg)
    c1, v1, c2, v2, end = (m.group(1), m.group(2), m.group(3),
                           m.group(4), m.group(5))
    if v1 is None:
        return [(int(c1), None, None)]                 # whole chapter
    if c2 is None:
        return [(int(c1), int(v1), int(end) if end else int(v1))]
    specs = [(int(c1), int(v1), None)]                 # c1:v1 -> end of c1
    for c in range(int(c1) + 1, int(c2)):
        specs.append((c, None, None))
    specs.append((int(c2), None, int(v2) if v2 else None))
    return specs


def _fetch_chapter(code, chap, version):
    url = "%s?%s" % (API, urllib.parse.urlencode(
        {"chineses": code, "chap": chap, "version": version}))
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=20) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    if data.get("status") != "success":
        raise ValueError("聖經網站回應失敗")
    return data.get("record", [])


def split_refs(ref):
    """Split an 出處 field into individual references (`;`, `；` or newline)."""
    return [p.strip() for p in re.split(r"[;；\n]+", str(ref)) if p.strip()]


def fetch_reference(ref):
    """Fetch one or more references (separated by `;`, `；` or newlines) and
    return (text, info).

    `text` is the whole passage as one continuous paragraph (整段), rendered in
    無空版 (SSS=2, no spacing before 神) with plain inline verse numbers — e.g.
    `1（稱謝詩。）普天下當向耶和華歡呼！ 2 …`.  With several books each gets
    its own paragraph headed by its `書 章:節` label (e.g. `詩篇 100:1-2`).
    Matches the site's `read.php?...&TABFLAG=2&SSS=2` view.
    """
    refs = split_refs(ref)
    if not refs:
        raise ValueError("請先輸入出處，例如「詩篇 34:1-3」")
    chunks = []
    for one in refs:
        code, version, specs = parse_reference(one)
        many_chaps = len({s[0] for s in specs}) > 1
        parts = []
        for chap, start, end in specs:
            records = _fetch_chapter(code, chap, version)
            if not records:
                raise ValueError("查無經文「%s」" % one)
            lo = start or 1
            hi = end or max(r.get("sec", 0) for r in records)
            for rec in records:
                sec = rec.get("sec")
                if sec is not None and lo <= sec <= hi:
                    chap_no = rec.get("chap") or chap
                    no = "%d:%d" % (chap_no, sec) if many_chaps else "%d" % sec
                    parts.append("%s%s" % (no, _clean(rec.get("bible_text", ""))))
        if parts:
            chunks.append((_strip_note(one), parts))
    if not chunks:
        raise ValueError("查無經文「%s」" % ref)
    labels = "；".join(h for h, _ in chunks)
    verses = sum(len(parts) for _, parts in chunks)
    if len(chunks) == 1:
        text = " ".join(chunks[0][1])
    else:
        text = "\n\n".join("%s\n%s" % (heading, " ".join(parts))
                           for heading, parts in chunks)
    info = "「%s」· %d 節" % (labels, verses)
    return text, info


def _clean(text):
    """Drop inline tags / line breaks; 無空版 = remove the \u3000 spacing."""
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"[\r\n\t ]+", " ", text)
    return text.replace("\u3000", "").strip()


def _strip_note(ref):
    return re.sub(r"[（(][^）)]*[）)]", "", str(ref)).strip()
