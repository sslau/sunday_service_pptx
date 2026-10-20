# scripture_sheet.py — read the 讀經經文 reference for a date from a Google Sheet.
"""A read-only companion to `store.py`.

The worship team keeps the week's plan in a Google Sheet: the Worship_Link tab
lists the sermon passage per service.  Column A is 日期 (the date) and column C
is the passage reference (e.g. `與#3 約 1：40-51`).  This module looks a date up
and returns the **reference**, already cleaned of any leading series marker, so
the 讀經經文 step can drop it into 出處 and let the fhl.net fetch fill the verses.

Auth is the same service account the week store already uses
(`gcp_service_account` in secrets.toml).  The reading sheet must be shared with
that service account's email as Viewer — otherwise Google answers with a
`PermissionError`, which `fetch_reference` surfaces as a short, readable message
instead of a traceback.

Secrets layout (secrets.toml):

    [scripture_sheet]
    # Spreadsheet title, id, or full URL (any works).
    spreadsheet = "https://docs.google.com/spreadsheets/d/<id>/edit"
    # The tab to read (a title or a gid number). Default: "Worship_Link".
    worksheet = "Worship_Link"
    # Optional: column letters — default A (日期) and C (經文).
    date_column = "A"
    text_column = "C"
"""
import re
import time

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# Same idea as store.py: keep the authorized client (and its HTTP session)
# alive across Streamlit reruns, and bound how stale the parsed rows can get.
_CLIENT_CACHE = {}
_ROWS_CACHE = {}
CACHE_TTL = 30.0

DEFAULT_SPREADSHEET = (
    "https://docs.google.com/spreadsheets/d/"
    "14RJ-lu2SmGGuHBphZuekQA6rCij8QT4ZBDgZ1AsfuUI/edit")
DEFAULT_WORKSHEET = "Worship_Link"


def _cfg(secrets):
    section = (secrets or {}).get("scripture_sheet") or {}
    if not isinstance(section, dict):
        section = {}
    return {
        "spreadsheet": str(section.get("spreadsheet") or DEFAULT_SPREADSHEET),
        "worksheet": str(section.get("worksheet") or DEFAULT_WORKSHEET),
        "date_column": str(section.get("date_column") or "A"),
        "text_column": str(section.get("text_column") or "C"),
    }


def _sa_dict(secrets):
    sa = (secrets or {}).get("gcp_service_account") or {}
    if not (sa.get("client_email") and sa.get("private_key")):
        return None
    return dict(sa)


def _client(sa):
    key = sa.get("client_email")
    if key in _CLIENT_CACHE:
        return _CLIENT_CACHE[key]
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(sa, scopes=SCOPES)
    client = gspread.authorize(creds)
    _CLIENT_CACHE[key] = client
    return client


def _col_index(letter):
    """`"A"` -> 0, `"C"` -> 2 (spreadsheet column letters to a 0-based index)."""
    letter = (letter or "").strip().upper()
    if not letter.isalpha():
        return 0
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _norm_date(value):
    """Any common date shape -> `"YYYY.MM.DD"` for comparison, else "".

    Accepts the sheet's `M/D/YY` (and `M/D/YYYY`), ISO/dotted dates, and real
    date cells (which arrive as `YYYY-MM-DD` or with a time riding along)."""
    s = str(value or "").strip()
    if not s:
        return ""
    m = re.match(r"(\d{4})[-/.]\s*(\d{1,2})[-/.]\s*(\d{1,2})", s)
    if m:
        return "%s.%02d.%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
    # M/D/YY or M/D/YYYY — the sheet's own format, e.g. 10/25/26.
    m = re.match(r"(\d{1,2})[-/.]\s*(\d{1,2})[-/.]\s*(\d{2,4})", s)
    if m:
        year = int(m.group(3))
        if year < 100:
            year += 2000
        return "%04d.%02d.%02d" % (year, int(m.group(1)), int(m.group(2)))
    return ""


def clean_reference(raw):
    """Drop a leading series marker from a passage reference.

    The sheet writes e.g. `與#3 約 1：40-51` or `將臨期#1 賽 9:1-6`; only the
    part from the first recognised book name on is an actual reference.  Series
    names change over time, so rather than matching them by name this finds the
    first book name in the string and keeps everything from there.  Anything
    with no book name is returned unchanged (trimmed)."""
    s = re.sub(r"\s+", " ", str(raw or "").replace("\u3000", " ")).strip()
    if not s:
        return ""
    import bible

    # Try each space-separated tail: "與#3 約 1：40-51" -> "約 1：40-51" matches.
    parts = s.split(" ")
    for i in range(len(parts)):
        tail = " ".join(parts[i:])
        name, rest = bible.split_book(tail)
        if name and rest:
            return tail
    # Also handle a marker with no space before the book ("與#3約 1:40-51").
    for i in range(1, len(s)):
        tail = s[i:].strip()
        name, rest = bible.split_book(tail)
        if name and rest:
            return tail
    return s


def _pick_worksheet(sheet, want):
    """The tab named/`gid`-numbered `want`, else the first tab."""
    want = (want or "").strip()
    if not want:
        return sheet.get_worksheet(0)
    if want.isdigit():
        for ws in sheet.worksheets():
            if str(ws.id) == want:
                return ws
    for ws in sheet.worksheets():
        if ws.title == want:
            return ws
    return sheet.get_worksheet(0)


def _load_rows(secrets, fresh=False):
    """Every row of the reading sheet as a list of value-lists (cached)."""
    cfg = _cfg(secrets)
    sa = _sa_dict(secrets)
    if not sa:
        raise RuntimeError(
            "尚未設定 Google 服務帳戶（secrets.toml 的 gcp_service_account）")
    cache_key = (sa.get("client_email"), cfg["spreadsheet"], cfg["worksheet"])
    hit = _ROWS_CACHE.get(cache_key)
    if hit and not fresh and time.monotonic() < hit[0]:
        return hit[1]
    try:
        client = _client(sa)
        sheet = client.open_by_url(cfg["spreadsheet"]) \
            if "docs.google.com" in cfg["spreadsheet"] \
            else client.open_by_key(cfg["spreadsheet"])
        ws = _pick_worksheet(sheet, cfg["worksheet"])
        rows = ws.get_all_values()
    except PermissionError:
        raise RuntimeError(
            "無法讀取經文表：請把試算表分享給服務帳戶 "
            f"{sa.get('client_email')}（至少「檢視者」權限）")
    except Exception as exc:      # network / bad id / gspread API changes
        raise RuntimeError(f"讀取經文表失敗：{exc}")
    _ROWS_CACHE[cache_key] = (time.monotonic() + CACHE_TTL, rows)
    return rows


def fetch_reference(date, secrets, fresh=False):
    """The passage reference for `date` (`"YYYY.MM.DD"`), or None when the row
    is absent or its 經文 cell is empty.

    The returned string is cleaned of any leading series marker, so it is ready
    for the 出處 field.  Raises RuntimeError with a readable message on a
    permission or network problem."""
    rows = _load_rows(secrets, fresh=fresh)
    cfg = _cfg(secrets)
    want = _norm_date(date)
    if not want:
        return None
    di, ti = _col_index(cfg["date_column"]), _col_index(cfg["text_column"])
    for row in rows:
        if di >= len(row):
            continue
        if _norm_date(row[di]) != want:
            continue
        if ti >= len(row):
            return None
        text = str(row[ti] or "").strip()
        return clean_reference(text) or None
    return None


def invalidate():
    """Drop the cached rows (keeps the authorized client)."""
    _ROWS_CACHE.clear()
