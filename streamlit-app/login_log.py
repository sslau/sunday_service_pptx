# login_log.py — login events + per-session activity summaries in the Google
# Sheet: "logs" worksheet keeps one raw row per login attempt; "sessions"
# worksheet keeps one summary row per login session (login time, IP, UA,
# action count, and a running summary of what was updated).  Falls back to
# local JSON files on dev machines without Sheets credentials.
import json
import os
import threading
import time

SHEET_TAB = "logs"
SESSIONS_TAB = "sessions"
SESSIONS_HEADER = ["session_id", "login_ts", "updated_at", "ip", "ua",
                   "n_actions", "summary"]
SHEET_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
MAX_LOCAL_ENTRIES = 200
MAX_SUMMARY_ACTIONS = 8
_lock = threading.Lock()

LOCAL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "data", "login_log.json")
LOCAL_SESSIONS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   "data", "sessions.json")


# ---- spreadsheet plumbing ---------------------------------------------------
def _open(client, spec):
    spec = (spec or "").strip()
    if "docs.google.com" in spec or spec.startswith("http"):
        return client.open_by_url(spec)
    if " " not in spec and len(spec) >= 20:
        return client.open_by_key(spec)
    return client.open(spec)


def _sheet(secrets):
    """gspread Spreadsheet for the app's store, or None when credentials or
    the spreadsheet spec are missing (local dev without Sheets)."""
    sa = (secrets or {}).get("gcp_service_account") or {}
    spec = ((secrets or {}).get("store") or {}).get("spreadsheet")
    if not (sa.get("client_email") and sa.get("private_key") and spec):
        return None
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(
        dict(sa), scopes=SHEET_SCOPES)
    return _open(gspread.authorize(creds), spec)


def _get_or_create_ws(sheet, title, header, cols):
    for w in sheet.worksheets():
        if w.title == title:
            return w
    ws = sheet.add_worksheet(title, rows=100, cols=cols)
    ws.update(range_name="A1:%s1" % chr(ord("A") + cols - 1), values=[header])
    return ws


def _logs_ws(sheet):
    return _get_or_create_ws(sheet, SHEET_TAB, ["ts", "ip", "ua", "result"], 4)


def _sessions_ws(sheet):
    return _get_or_create_ws(sheet, SESSIONS_TAB, SESSIONS_HEADER,
                             len(SESSIONS_HEADER))


# ---- login events -----------------------------------------------------------
def append_login(result, ip="", ua="", secrets=None):
    """Record one login attempt (`result`: 'success' | 'failed') in the
    "logs" worksheet (local JSON fallback when Sheets unavailable)."""
    rec = {
        "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
        "ip": (ip or "").split(",")[0].strip(),
        "ua": (ua or "")[:200],
        "result": result,
    }
    with _lock:
        try:
            sheet = _sheet(secrets)
            if sheet is not None:
                _logs_ws(sheet).append_row(
                    [rec["ts"], rec["ip"], rec["ua"], rec["result"]])
                return rec
        except Exception as exc:
            print("login_log -> sheets failed (%s); using local fallback"
                  % exc)
        _append_local(rec)
    return rec


def _append_local(rec):
    entries = load()
    entries.append(rec)
    entries = entries[-MAX_LOCAL_ENTRIES:]
    os.makedirs(os.path.dirname(LOCAL_PATH), exist_ok=True)
    with open(LOCAL_PATH, "w", encoding="utf-8") as fh:
        json.dump(entries, fh, ensure_ascii=False, indent=2)


def load():
    try:
        with open(LOCAL_PATH, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return []


def clear():
    with _lock:
        try:
            os.remove(LOCAL_PATH)
        except OSError:
            pass


# ---- per-session summaries --------------------------------------------------
def _updated_summary(old_summary, n, label, detail):
    parts = [p for p in (old_summary or "").split("；") if p.strip()]
    parts = parts[-(MAX_SUMMARY_ACTIONS - 1):]
    parts.append(f"{n}. {label}"
                 + (f"（{detail}）" if detail else ""))
    return "；".join(parts)


def _local_sessions():
    try:
        with open(LOCAL_SESSIONS_PATH, encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _save_local_sessions(data):
    os.makedirs(os.path.dirname(LOCAL_SESSIONS_PATH), exist_ok=True)
    with open(LOCAL_SESSIONS_PATH, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)


def new_session(session_id, ip="", ua="", secrets=None):
    """Record the start of a login session: one row in the "sessions"
    worksheet (or the local sessions.json fallback)."""
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    rec = {
        "session_id": session_id, "login_ts": ts, "updated_at": ts,
        "ip": (ip or "").split(",")[0].strip(), "ua": (ua or "")[:200],
        "n_actions": 1, "summary": "1. 登入成功",
    }
    with _lock:
        try:
            sheet = _sheet(secrets)
            if sheet is not None:
                _sessions_ws(sheet).append_row(
                    [rec["session_id"], rec["login_ts"], rec["updated_at"],
                     rec["ip"], rec["ua"], rec["n_actions"], rec["summary"]])
                return rec
        except Exception as exc:
            print("sessions -> sheets failed (%s); using local fallback"
                  % exc)
        data = _local_sessions()
        data[session_id] = rec
        _save_local_sessions(data)
    return rec


def log_action(session_id, label, detail="", secrets=None, ip="", ua=""):
    """Append one important action and refresh that session's summary row
    (n_actions counter + running summary)."""
    now = time.strftime("%Y-%m-%d %H:%M:%S")
    with _lock:
        try:
            sheet = _sheet(secrets)
        except Exception as exc:
            sheet = None
            print("sessions -> sheets failed (%s); using local fallback"
                  % exc)
        if sheet is None:
            data = _local_sessions()
            rec = data.get(session_id)
            if rec is None:
                ts = time.strftime("%Y-%m-%d %H:%M:%S")
                rec = {
                    "session_id": session_id, "login_ts": ts,
                    "updated_at": ts,
                    "ip": (ip or "").split(",")[0].strip(),
                    "ua": (ua or "")[:200], "n_actions": 1, "summary": "",
                }
                data[session_id] = rec
            n = int(rec.get("n_actions") or 1) + 1
            rec["n_actions"] = n
            rec["updated_at"] = now
            rec["summary"] = _updated_summary(rec.get("summary", ""),
                                              n, label, detail)
            _save_local_sessions(data)
            return rec
        try:
            ws = _sessions_ws(sheet)
            values = ws.get_all_values()
            hit = None
            for i, row in enumerate(values):
                if row and row[0] == session_id:
                    hit = i + 1
                    break
            if hit is None:
                row = [session_id, now, now,
                       (ip or "").split(",")[0].strip() or "",
                       (ua or "")[:200], 2,
                       _updated_summary("", 2, label, detail)]
                ws.append_row(row)
                return dict(zip(SESSIONS_HEADER, row))
            row = values[hit - 1]
            ipv = row[3] if len(row) > 3 else (ip or "").split(",")[0].strip()
            uav = row[4] if len(row) > 4 else (ua or "")[:200]
            login_ts = row[1] if len(row) > 1 else now
            old_summary = row[6] if len(row) > 6 else ""
            old_n = int(row[5]) if len(row) > 5 and str(row[5]).isdigit() else 1
            n = old_n + 1
            row = [session_id, login_ts, now, ipv, uav, n,
                   _updated_summary(old_summary, n, label, detail)]
            ws.update("A%d:G%d" % (hit, hit), [row])
            return dict(zip(SESSIONS_HEADER, row))
        except Exception as exc:
            print("sessions -> sheets failed (%s); using local fallback"
                  % exc)
            data = _local_sessions()
            rec = data.get(session_id)
            if rec is None:
                rec = {"session_id": session_id,
                       "login_ts": now, "updated_at": now,
                       "ip": (ip or "").split(",")[0].strip(),
                       "ua": (ua or "")[:200], "n_actions": 1,
                       "summary": ""}
                data[session_id] = rec
            n = int(rec.get("n_actions") or 1) + 1
            rec["n_actions"] = n
            rec["updated_at"] = now
            rec["summary"] = _updated_summary(rec.get("summary", ""),
                                              n, label, detail)
            _save_local_sessions(data)
            return rec