"""Read the weekly scripture passage from the worship-team Google Sheet."""
from datetime import date, datetime
import re

import store

_WORKSHEET_CACHE = {}


def _column_index(value):
    """Convert a Google Sheets column label such as A or AA to a zero-based index."""
    label = str(value or "").strip().upper()
    if not re.fullmatch(r"[A-Z]+", label):
        raise ValueError("scripture_sheet 欄位必須使用 A、B、C 等欄位代號。")
    index = 0
    for char in label:
        index = index * 26 + ord(char) - ord("A") + 1
    return index - 1


def _as_date(value):
    """Parse common formatted Google Sheets date values."""
    text = str(value or "").strip()
    match = re.search(r"(\d{4})[./-](\d{1,2})[./-](\d{1,2})", text)
    if match:
        try:
            return date(*(int(part) for part in match.groups()))
        except ValueError:
            return None
    if re.fullmatch(r"\d{8}", text):
        try:
            return datetime.strptime(text, "%Y%m%d").date()
        except ValueError:
            return None
    for fmt in ("%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def _worksheet(secrets):
    config = (secrets or {}).get("scripture_sheet") or {}
    spreadsheet = str(config.get("spreadsheet") or "").strip()
    worksheet_name = str(config.get("worksheet") or "").strip()
    service_account = store._sa_dict(secrets or {})
    if not (spreadsheet and worksheet_name and service_account):
        raise RuntimeError(
            "請設定 secrets.toml 的 [scripture_sheet] "
            "（spreadsheet、worksheet、date_column、text_column）及 "
            "[gcp_service_account]。")

    cache_key = (spreadsheet, worksheet_name,
                 service_account.get("client_email"))
    if cache_key not in _WORKSHEET_CACHE:
        import gspread
        from google.oauth2.service_account import Credentials

        credentials = Credentials.from_service_account_info(
            service_account, scopes=store.SCOPES)
        sheet = store._open_spreadsheet(
            gspread.authorize(credentials), spreadsheet)
        try:
            _WORKSHEET_CACHE[cache_key] = sheet.worksheet(worksheet_name)
        except Exception as exc:
            raise RuntimeError(
                f"找不到 Worship_Link 工作表「{worksheet_name}」：{exc}") from exc
    return config, _WORKSHEET_CACHE[cache_key]


def fetch_scripture(secrets, service_date):
    """Return column-C scripture text from the configured row matching a date."""
    config, worksheet = _worksheet(secrets)
    date_column = _column_index(config.get("date_column", "A"))
    text_column = _column_index(config.get("text_column", "C"))
    try:
        wanted_date = datetime.strptime(
            str(service_date).replace(".", "-"), "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"無效的主日日期：{service_date}") from exc

    for row_number, row in enumerate(worksheet.get_all_values(), start=1):
        if date_column >= len(row) or _as_date(row[date_column]) != wanted_date:
            continue
        text = row[text_column].strip() if text_column < len(row) else ""
        if not text:
            raise LookupError(
                f"Worship_Link 第 {row_number} 列的讀經經文是空白。")
        return text

    raise LookupError(f"Worship_Link 找不到 {service_date} 的讀經經文。")
