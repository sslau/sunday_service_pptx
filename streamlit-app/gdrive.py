# gdrive.py — Google Drive read access via the same service account used for
# Google Sheets, using the Drive read-only scope. Pure Python (google-auth +
# requests); no extra system dependencies, safe for Streamlit Cloud.
'''Fetch section decks / MP4s from a shared Google Drive folder instead of a
manual upload.  Requires:
  * the Drive API enabled for the service account's GCP project, and
  * the target folder shared (Viewer) with the service-account email.

    sa = {"client_email": ..., "private_key": ...}   # from st.secrets
    files = gdrive.list_files(sa, folder_id)
    ok, err = gdrive.download_to(sa, file_id, "data/decks/.../x.pptx")
'''
import os
import re
import tempfile

_DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.readonly"
_FILES_URL = "https://www.googleapis.com/drive/v3/files"
_FOLDER_RE = re.compile(r"/folders/([A-Za-z0-9_-]+)")
_ID_RE = re.compile(r"[?&]id=([A-Za-z0-9_-]+)")
_PLAIN_RE = re.compile(r"[A-Za-z0-9_-]{10,}")


def folder_id_from(value):
    """Extract a Drive folder id from a URL or a bare id."""
    value = (value or "").strip()
    if not value:
        return None
    m = _FOLDER_RE.search(value)
    if m:
        return m.group(1)
    m = _ID_RE.search(value)
    if m:
        return m.group(1)
    if _PLAIN_RE.fullmatch(value):
        return value
    return None


def _session(sa_dict):
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(
        dict(sa_dict), scopes=[_DRIVE_SCOPE])
    return AuthorizedSession(creds)


def list_files(sa_dict, folder_id):
    """List files directly inside the folder.
    Returns [{id, name, mimeType, size}] sorted by name."""
    out = []
    base = {
        "q": f"'{folder_id}' in parents and trashed=false",
        "fields": "nextPageToken, files(id,name,mimeType,size)",
        "pageSize": "1000",
        "supportsAllDrives": "true",
        "includeItemsFromAllDrives": "true",
    }
    page = None
    with _session(sa_dict) as session:
        while True:
            params = dict(base)
            if page:
                params["pageToken"] = page
            resp = session.get(_FILES_URL, params=params)
            if resp.status_code != 200:
                raise RuntimeError(
                    "Drive 列表失敗 (%s)：%s" % (resp.status_code,
                                                resp.text[:300]))
            data = resp.json()
            out.extend(data.get("files", []))
            page = data.get("nextPageToken")
            if not page:
                break
    return sorted(out, key=lambda f: (f.get("name") or "").lower())


_FOLDER_MIME = "application/vnd.google-apps.folder"


def _date_variants(date):
    """All plausible spellings of a date: YYYY.MM.DD, YYYY-MM-DD, YYYYMMDD
    and the YYYY.mmdd folder convention (e.g. 2026.0927)."""
    s = str(date)
    out = {s, s.replace(".", "-"), s.replace(".", "")}
    parts = s.split(".")
    if len(parts) == 3 and parts[0] and parts[1] and parts[2]:
        out.add("%s.%s%s" % (parts[0], parts[1], parts[2]))
    return out


def _date_subfolder(top, date):
    """Best guess of the per-date subfolder: exact name match first, then any
    folder whose name contains one of the date spellings."""
    variants = _date_variants(date)
    for f in top:
        name = (f.get("name") or "").strip()
        if f.get("mimeType") == _FOLDER_MIME and (
                name in variants or any(v in name for v in variants if len(v) >= 8)):
            return f
    return None


def files_in_date_folder(sa_dict, folder_id, date):
    """Return files under the per-date subfolder (name like 2026.09.20 /
    2026-09-20 / 20260920). Falls back to the top folder's own files when no
    matching subfolder exists."""
    top = list_files(sa_dict, folder_id)
    sub = _date_subfolder(top, date)
    if sub:
        return list_files(sa_dict, sub["id"])
    return top


def folder_contents(sa_dict, folder_id, date):
    """Return (top, sub, inner) so the UI can show what the service account
    actually sees: the top folder's items, the matched date subfolder (or
    None), and the subfolder's inner items (or [])."""
    top = list_files(sa_dict, folder_id)
    sub = _date_subfolder(top, date)
    inner = list_files(sa_dict, sub["id"]) if sub else []
    return top, sub, inner


def download_to(sa_dict, file_id, target_path):
    """Stream a Drive file's bytes into target_path (tmp file, then rename).
    Returns (ok, err); err is None on success."""
    url = "%s/%s?alt=media" % (_FILES_URL, file_id)
    target_path = os.path.abspath(target_path)
    parent = os.path.dirname(target_path)
    os.makedirs(parent, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=parent, suffix=".part")
    try:
        with _session(sa_dict) as session:
            with session.get(url, stream=True) as resp:
                if resp.status_code != 200:
                    return False, ("Drive 下載失敗 (%s)：%s"
                                   % (resp.status_code, resp.text[:300]))
                with os.fdopen(fd, "wb") as fh:
                    for chunk in resp.iter_content(1 << 16):
                        if chunk:
                            fh.write(chunk)
        os.replace(tmp, target_path)
        return True, None
    except Exception as exc:  # noqa: BLE001 — surfaced to the UI as-is
        try:
            os.remove(tmp)
        except OSError:
            pass
        return False, str(exc)