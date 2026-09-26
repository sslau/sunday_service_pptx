# onedrive.py — read section decks / MP4s from a OneDrive or SharePoint site
# folder via the Microsoft Graph API, app-only auth (client credentials).
# Pure Python (requests), no extra system dependencies.
'''Fetch section decks / MP4s from an OneDrive/SharePoint folder instead of a
manual upload.  Requires an Azure AD app registration with the application
permission `Sites.Selected` (granted on the target site) or `Sites.Read.All`,
and the following in st.secrets under `[onedrive]`:

    [onedrive]
    tenant_id     = "<Entra ID tenant GUID>"
    client_id     = "<app registration application (client) id>"
    client_secret = "<client secret>"
    site_host     = "my.sharepoint.com"          # host without scheme
    site_path     = "/sites/SEACantoneseMedia"   # or "/personal/alias"
    folder_path   = "Documents/…/SEA-Cantonese"  # drive-relative folder

    cfg = dict(st.secrets["onedrive"])
    files = onedrive.files_in_date_folder(cfg, folder_path, date)
    ok, err = onedrive.download_to(cfg, file_id, "data/decks/…/x.pptx")
'''
import os
import tempfile
import time
import urllib.parse

import requests

_AUTH_URL = "https://login.microsoftonline.com/%s/oauth2/v2.0/token"
_GRAPH = "https://graph.microsoft.com/v1.0"
_SCOPE = "https://graph.microsoft.com/.default"

# Hedge: token caches for the lifetime of the process. Refreshed automatically
# when near expiry. Shared across script reruns (module-level singleton).
_token = {"at": None, "exp": 0.0}


def _access_token(cfg):
    """App-only OAuth2 client-credentials token for Microsoft Graph."""
    if _token["at"] and time.time() < _token["exp"] - 60:
        return _token["at"]
    resp = requests.post(
        _AUTH_URL % cfg.get("tenant_id", ""),
        data={
            "grant_type": "client_credentials",
            "client_id": cfg.get("client_id", ""),
            "client_secret": cfg.get("client_secret", ""),
            "scope": _SCOPE,
        },
        timeout=30)
    if resp.status_code != 200:
        raise RuntimeError(
            "OneDrive 認證失敗 (%s)：%s" % (resp.status_code, resp.text[:300]))
    data = resp.json()
    _token["at"] = data.get("access_token", "")
    _token["exp"] = time.time() + float(data.get("expires_in", 3600))
    return _token["at"]


def _site_ref(cfg):
    """Graph site reference.  With the Sites.Selected permission, Graph only
    accepts direct site addressing (sites/{site-id}); hostname:+path needs
    Sites.Read.All.  site_id wins when present, otherwise falls back to
    hostname + server-relative path."""
    site_id = (cfg.get("site_id") or "").strip()
    if site_id:
        return "%s/sites/%s" % (_GRAPH, site_id)
    host = (cfg.get("site_host") or "").strip()
    site_path = (cfg.get("site_path") or "").strip()
    if not host or not site_path:
        raise RuntimeError(
            "請設定 site_id，或 site_host + site_path（secrets.toml [onedrive]）")
    return "%s/sites/%s:/%s" % (_GRAPH, host, site_path.lstrip("/"))


def _list_items(cfg, drive_path):
    """List drive items under a drive-relative path.
    Root folder: pass "" -> /drive/root/children.
    Returns [{id, name, size, folder}]."""
    site = _site_ref(cfg)
    # Path-based site addressing needs an extra ':' before /drive, the
    # site-id form does not.
    sep = ":" if not (cfg.get("site_id") or "").strip() else ""
    if drive_path:
        quoted = urllib.parse.quote(drive_path.lstrip("/"), safe="/")
        base = "%s%s/drive/root:/%s:/children" % (site, sep, quoted)
    else:
        base = "%s%s/drive/root/children" % (site, sep)
    headers = {"Authorization": "Bearer %s" % _access_token(cfg)}
    out = []
    while base:
        resp = requests.get(base, headers=headers, timeout=30)
        if resp.status_code != 200:
            raise RuntimeError(
                "OneDrive 列表失敗 (%s)：%s" % (resp.status_code, resp.text[:300]))
        data = resp.json()
        for it in data.get("value", []):
            out.append({
                "id": it.get("id"),
                "name": it.get("name", ""),
                "size": it.get("size"),
                "folder": bool(it.get("folder")),
            })
        base = data.get("@odata.nextLink")
    return out


def _date_variants(date):
    """All plausible spellings of a date: YYYY.MM.DD, YYYY-MM-DD, YYYYMMDD
    and the YYYY.mmdd folder convention (e.g. 2026.0927)."""
    s = str(date)
    out = {s, s.replace(".", "-"), s.replace(".", "")}
    parts = s.split(".")
    if len(parts) == 3 and parts[0] and parts[1] and parts[2]:
        out.add("%s.%s%s" % (parts[0], parts[1], parts[2]))
    return out


def list_files(cfg, folder_path):
    """List files directly inside a drive folder (by path)."""
    items = _list_items(cfg, folder_path or "")
    return [f for f in items if not f.get("folder")]


def folders(cfg, folder_path):
    """List drive folders directly inside a folder (by path)."""
    items = _list_items(cfg, folder_path or "")
    return [f for f in items if f.get("folder")]


def _date_subfolder(cfg, folder_path, date):
    """Best guess of the per-date subfolder: exact name match first, then any
    folder whose name contains one of the date spellings."""
    variants = _date_variants(date)
    for f in folders(cfg, folder_path):
        name = (f.get("name") or "").strip()
        if (name in variants or
                any(v in name for v in variants if len(v) >= 8)):
            return f
    return None


def files_in_date_folder(cfg, folder_path, date):
    """Return files under the per-date subfolder (name like 2026.09.20 /
    2026-09-20 / 20260920 / 2026.0920). Falls back to the top folder's own
    files when no matching subfolder exists."""
    sub = _date_subfolder(cfg, folder_path, date)
    if sub:
        items = _list_items(cfg,
                            (folder_path or "").rstrip("/") + "/" + sub["name"])
        return [f for f in items if not f.get("folder")]
    return list_files(cfg, folder_path)


def download_to(cfg, file_id, target_path):
    """Stream a drive file's bytes into target_path (tmp file, then rename).
    Returns (ok, err); err is None on success."""
    site = _site_ref(cfg)
    sep = ":" if not (cfg.get("site_id") or "").strip() else ""
    url = "%s%s/drive/items/%s/content" % (site, sep, file_id)
    headers = {"Authorization": "Bearer %s" % _access_token(cfg)}
    target_path = os.path.abspath(target_path)
    parent = os.path.dirname(target_path)
    os.makedirs(parent, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=parent, suffix=".part")
    try:
        with requests.get(url, headers=headers, stream=True, timeout=60) as resp:
            if resp.status_code != 200:
                return False, ("OneDrive 下載失敗 (%s)：%s"
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