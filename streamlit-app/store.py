# store.py — persistent storage for weekly service data.
'''Two backends exposing the same dict[date -> week] interface:

- SheetsStore  — Google Sheets via gspread (pure Python, NO Apps Script).
                 One worksheet, columns: date | json | updatedAt.
- LocalStore   — a local JSON file (local dev / testing only; Community Cloud
                 wipes its filesystem on every restart/redeploy, so never rely
                 on it there).

Backend choice: if valid Google credentials are configured (passed via the
`secrets` dict, itself straight from st.secrets) a SheetsStore is used,
otherwise a LocalStore.  gspread is imported lazily so the app still runs on a
machine without it (e.g. this Mac) using LocalStore.

Secrets layout for the Sheets backend (secrets.toml):

    [store]
    # Spreadsheet title, id, or full URL (any works). Must be shared with the
    # service-account email as Editor.
    spreadsheet = "https://docs.google.com/spreadsheets/d/<id>/edit"
    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = """-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"""
    client_email = "build@....iam.gserviceaccount.com"
    client_id = "..."
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.googleapis.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/build%40....iam.gserviceaccount.com"
'''
import json
import os

SHEET_NAME = "weeks"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def _sa_dict(secrets):
    sa = secrets.get("gcp_service_account") or {}
    if not (sa.get("client_email") and sa.get("private_key")):
        return None
    return dict(sa)


class LocalStore:
    """JSON file backend: {date: week, ...}."""

    def __init__(self, path):
        self.path = path

    def all(self):
        try:
            with open(self.path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    def get(self, date):
        return self.all().get(date)

    def save(self, week):
        data = self.all()
        data[week.get("date")] = week
        with open(self.path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)


def _open_spreadsheet(client, spec):
    """Open a spreadsheet by URL, id, or title (whichever `spec` looks like)."""
    spec = (spec or "").strip()
    if "docs.google.com" in spec or spec.startswith("http"):
        return client.open_by_url(spec)
    if " " not in spec and len(spec) >= 20:
        return client.open_by_key(spec)
    return client.open(spec)


class SheetsStore:
    """Google Sheets backend via gspread (no Apps Script involved)."""

    def __init__(self, sa_dict, spreadsheet):
        import gspread
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_info(sa_dict, scopes=SCOPES)
        self.sheet = _open_spreadsheet(gspread.authorize(creds), spreadsheet)
        self.ws = None
        for w in self.sheet.worksheets():
            if w.title == SHEET_NAME:
                self.ws = w
                break
        if self.ws is None:
            self.ws = self.sheet.add_worksheet(SHEET_NAME, rows=100, cols=3)
            self.ws.update("A1:C1", [["date", "json", "updatedAt"]])

    def _rows(self):
        try:
            return self.ws.get_all_values()
        except Exception:
            return []

    def all(self):
        out = {}
        rows = self._rows()
        for row in rows[1:]:
            if len(row) >= 2 and row[0] and row[1]:
                try:
                    week = json.loads(row[1])
                    if isinstance(week, dict):
                        week["date"] = row[0]
                        out[row[0]] = week
                except (TypeError, ValueError):
                    continue
        return out

    def get(self, date):
        return self.all().get(date)

    def save(self, week):
        from datetime import datetime, timezone

        date = week.get("date")
        stale = self.all()
        now = datetime.now(timezone.utc).isoformat(timespec="seconds")
        stale[date] = week
        rows = [["date", "json", "updatedAt"]]
        for d in sorted(stale.keys()):
            rows.append([d, json.dumps(stale[d], ensure_ascii=False), now])
        self.ws.update(range_name="A1:C%d" % len(rows), values=rows)


def get_store(secrets=None, local_path=None):
    """Return a store. `secrets` is the st.secrets dict (or a plain dict); when
    it carries gspread credentials a SheetsStore is returned, otherwise a
    LocalStore at `local_path` (default: data/weeks.json next to this file)."""
    if local_path is None:
        local_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                  "data", "weeks.json")
    sa = _sa_dict(secrets or {})
    spreadsheet = (secrets or {}).get("store", {}).get("spreadsheet")
    if sa and spreadsheet:
        try:
            return SheetsStore(sa, spreadsheet)
        except Exception as exc:
            print("SheetsStore init failed (%s); falling back to LocalStore" % exc)
    return LocalStore(local_path)