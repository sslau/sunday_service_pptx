# app.py — Sunday Service slide web editor (Streamlit, pure Python).
"""Edit a week's worship content (psalm, hymns, scripture, sermon,
announcements), persist it to Google Sheets (gspread, no Apps Script) or a
local JSON fallback, and generate the .pptx fully server-side.

Run locally (needs python-pptx; gspread optional):
    streamlit run app.py

Deploy: push this folder to GitHub → Streamlit Community Cloud (see README).
Data lives in Google Sheets so Community Cloud's ephemeral disk is never
relied on.
"""
import os
import uuid
from datetime import date as _date

import streamlit as st

st.set_page_config(page_title="崇拜投影片編輯器", page_icon="✝",
                   layout="wide")

import bible  # noqa: E402
from deck_builder import (IMAGE_SLOTS, build_deck,  # noqa: E402
                          build_section_pptx_bytes, resolve_deck_path,
                          save_section_deck, saved_section_path)
from model import DATE_RE, default_week, normalize_week, split_lines  # noqa: E402
from store import get_store  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)

try:
    secrets_obj = st.secrets
    st_secrets = {}
    for key in ("store", "gcp_service_account"):
        try:
            st_secrets[key] = secrets_obj[key]
        except Exception:
            pass
except Exception:
    st_secrets = {}

store = get_store(st_secrets)

if "active_date" not in st.session_state:
    dates = sorted(store.all().keys(), reverse=True)
    st.session_state["active_date"] = dates[0] if dates else "2026.09.20"

DATE = st.session_state["active_date"]
P = f"{DATE}::"


def K(name):
    return f"{P}{name}"


def sv(name, default=""):
    return st.session_state.get(K(name), default)


def section_source(upload, name, section, date, use_saved):
    """Where a section's slides will come from, for the readiness indicator."""
    if upload is not None:
        return "已上載檔案（合併）", "ready"
    path = resolve_deck_path(name) if name else None
    if path:
        return f"使用檔案：{os.path.basename(path)}", "ready"
    if use_saved:
        saved = saved_section_path(section, date)
        if saved:
            return f"使用已儲存：{os.path.basename(saved)}", "saved"
    return "由網頁內容編譯", "compile"


def persist_uploads(date, uploads):
    """Save freshly uploaded section decks to data/decks/<date>/ so the
    section status survives a page refresh (Streamlit clears file_uploader
    state on reload). `uploads` maps section -> UploadedFile|None."""
    for sec, up in uploads.items():
        if up is None:
            continue
        sig = (getattr(up, "name", ""), getattr(up, "size", None))
        if st.session_state.get(f"{date}::uploadsig_{sec}") != sig:
            save_section_deck(sec, date, up.getvalue())
            st.session_state[f"{date}::uploadsig_{sec}"] = sig


def is_first_sunday(date_str):
    """True when date_str (YYYY.MM.DD or YYYY-MM-DD) is the month's first
    Sunday — the weeks that add 聖餐 + 使徒信經."""
    parts = str(date_str).replace("-", ".").split(".")
    if len(parts) != 3:
        return False
    try:
        y, m, d = (int(p) for p in parts)
        first = _date(y, m, 1)
    except ValueError:
        return False
    first_sunday = 1 + (6 - first.weekday()) % 7  # Monday=0 .. Sunday=6
    return d == first_sunday


def clear_date_state(date):
    prefix = f"{date}::"
    for key in [k for k in st.session_state.keys() if k.startswith(prefix)]:
        del st.session_state[key]


def set_active(date, seed_week=None):
    clear_date_state(date)
    week = seed_week or store.get(date) or default_week(date)
    st.session_state["active_date"] = date
    st.session_state[f"{date}::ui_hymns"] = _seed_hymns(week)
    st.session_state[f"{date}::ui_slides"] = _seed_slides(week)
    st.session_state[f"{date}::ui_ann"] = _seed_ann(week)
    _seed_fields(date, week)
    st.rerun()


def _seed_hymns(week):
    out = []
    for h in week.get("hymns", []):
        out.append({
            "_id": uuid.uuid4().hex[:8],
            "title": h.get("title", ""),
            "subtitle": h.get("subtitle", ""),
            "source": h.get("source", ""),
            "refrain": "\n".join(h.get("refrain") or []),
            "repeat": bool(h.get("refrain_after_every_verse", True)),
            "stanzas": ["\n".join(s) for s in h.get("verses", [])] or [""],
        })
    return out


def _seed_slides(week):
    return [{"_id": uuid.uuid4().hex[:8], "title": s.get("title", ""),
             "body": "\n".join(s.get("body", []))}
            for s in week.get("sermon", {}).get("slides", [])]


def _seed_ann(week):
    out = []
    for a in week.get("announcements", []):
        out.append({
            "_id": uuid.uuid4().hex[:8],
            "mode": "image" if a.get("image_name") else "text",
            "title": a.get("title", ""),
            "body": "\n".join(a.get("body", [])),
            "image_name": a.get("image_name", ""),
            "_upload": None,
        })
    return out


def _seed_fields(date, week):
    """Seed the scalar editor widgets (refs, font sizes, verse text areas,
    sermon meta, communion) from a stored/default week.  Without this, saved
    宣召經文／經文（讀經）等欄位開檔時顯示空白，再儲存就會被清空。"""
    psalm = week.get("psalm") or {}
    scripture = week.get("scripture") or {}
    sermon = week.get("sermon") or {}
    values = {
        "psalm_ref": psalm.get("ref", ""),
        "psalm_ref_size": psalm.get("ref_size", 48),
        "psalm_font_size": psalm.get("font_size", 44),
        "psalm_verses": "\n".join(psalm.get("verses") or []),
        "hymn_font_size": week.get("hymn_font_size", 44),
        "hymn_font_min": week.get("hymn_font_min", 36),
        "hymn_margin_in": week.get("hymn_margin_in", 0.83),
        "scripture_ref": scripture.get("ref", ""),
        "scripture_ref_size": scripture.get("ref_size", 48),
        "scripture_font_size": scripture.get("font_size", 44),
        "scripture_verses": "\n".join(scripture.get("verses") or []),
        "sermon_title": sermon.get("title", ""),
        "sermon_title_2": sermon.get("title_2", ""),
        "sermon_speaker": sermon.get("speaker", ""),
        "sermon_date": sermon.get("date", date),
        "sermon_ref": sermon.get("scripture", ""),
        "sermon_ref_size": sermon.get("ref_size", 28),
        "communion": bool(week.get("communion", False)),
    }
    for name, val in values.items():
        st.session_state[f"{date}::{name}"] = val


if f"{DATE}::ui_hymns" not in st.session_state:
    week = store.get(DATE) or default_week(DATE)
    st.session_state[f"{DATE}::ui_hymns"] = _seed_hymns(week)
    st.session_state[f"{DATE}::ui_slides"] = _seed_slides(week)
    st.session_state[f"{DATE}::ui_ann"] = _seed_ann(week)
    _seed_fields(DATE, week)


def _get_ui(kind):
    key = f"{DATE}::ui_{kind}"
    st.session_state.setdefault(key, [])
    return st.session_state[key]


def _add_item(kind, item):
    _get_ui(kind).append(item)


def _del_item(kind, index):
    ui = _get_ui(kind)
    if 0 <= index < len(ui):
        del ui[index]


# ---------------------------------------------------------------- assemble --
def assemble_week(date):
    Kp = lambda n: f"{date}::{n}"  # noqa: E731

    def gv(n, d):
        return st.session_state.get(Kp(n), d)

    ui_hymns = st.session_state.get(f"{date}::ui_hymns", [])
    ui_slides = st.session_state.get(f"{date}::ui_slides", [])
    ui_ann = st.session_state.get(f"{date}::ui_ann", [])

    week = {
        "date": date,
        "psalm": {
            "ref": gv("psalm_ref", ""),
            "ref_size": int(gv("psalm_ref_size", 48)),
            "font_size": int(gv("psalm_font_size", 44)),
            "verses": split_lines(gv("psalm_verses", "")),
        },
        "hymn_font_size": int(gv("hymn_font_size", 44)),
        "hymn_font_min": int(gv("hymn_font_min", 36)),
        "hymn_margin_in": float(gv("hymn_margin_in", 0.83)),
        "hymns": [assemble_hymn(date, h) for h in ui_hymns],
        "scripture": {
            "ref": gv("scripture_ref", ""),
            "ref_size": int(gv("scripture_ref_size", 48)),
            "font_size": int(gv("scripture_font_size", 44)),
            "max_lines": int(gv("scripture_max_lines", 4)),
            "verses": split_lines(gv("scripture_verses", "")),
        },
        "sermon": {
            "title": gv("sermon_title", ""),
            "title_2": gv("sermon_title_2", ""),
            "scripture": gv("sermon_ref", ""),
            "ref_size": int(gv("sermon_ref_size", 28)),
            "speaker": gv("sermon_speaker", ""),
            "date": gv("sermon_date", date),
            "slides": [
                {"title": st.session_state.get(f"{date}::slide_{s['_id']}_title", ""),
                 "body": split_lines(st.session_state.get(
                     f"{date}::slide_{s['_id']}_body", ""))}
                for s in ui_slides],
        },
        "announcements": [assemble_ann(date, a) for a in ui_ann],
        "communion": bool(gv("communion", False)),
    }
    return normalize_week(week)


def assemble_hymn(date, ui):
    base = f"{date}::hymn_{ui['_id']}_"
    refrain = st.session_state.get(base + "refrain", ui.get("refrain", ""))
    stanzas = []
    for j in range(max(len(ui.get("stanzas", ["", ])), 1)):
        txt = st.session_state.get(f"{base}verses_{j}", ui["stanzas"][j])
        stanzas.append(split_lines(txt))
    return {
        "title": st.session_state.get(base + "title", ui.get("title", "")),
        "subtitle": st.session_state.get(base + "subtitle", ui.get("subtitle", "")),
        "source": st.session_state.get(base + "source", ui.get("source", "")),
        "refrain": split_lines(refrain) or None,
        "refrain_after_every_verse": bool(
            st.session_state.get(base + "repeat", ui.get("repeat", True))),
        "verses": stanzas,
    }


def assemble_ann(date, ui):
    base = f"{date}::ann_{ui['_id']}_"
    upload = st.session_state.get(base + "img", None)
    ui["_upload"] = upload or ui.get("_upload")
    if upload is not None:
        ui["image_name"] = upload.name or ""
    return {
        "title": st.session_state.get(base + "title", ui.get("title", "")),
        "body": split_lines(st.session_state.get(base + "body", ui.get("body", ""))),
        "image": None,
        "image_name": ui.get("image_name", ""),
    }


# ------------------------------------------------------------------- style --
_CSS = """
<style>
:root {
  --accent:#3F5B8B; --accent-2:#5E76A8; --accent-soft:#EEF2F9;
  --ink:#1E2430; --muted:#6B7480; --line:#E3E7EE;
}
.block-container { padding-top:2.4rem; padding-bottom:3rem; max-width:1180px; }
#MainMenu, footer { visibility:hidden; }

[data-testid="stSidebar"] {
  background:linear-gradient(180deg,#F8FAFD,#EEF2F9);
  border-right:1px solid var(--line);
}
.brand { font-size:1.08rem; font-weight:700; color:var(--ink);
  margin:.1rem 0 1rem; }
.brand .brand-tail { color:var(--muted); font-weight:500; }

.hero { background:linear-gradient(135deg,var(--accent),var(--accent-2));
  color:#fff; border-radius:16px; padding:1.5rem 1.7rem; margin-bottom:1.1rem;
  box-shadow:0 8px 22px rgba(63,91,139,.18); }
.hero-date { font-size:1.75rem; font-weight:700; letter-spacing:.02em; }
.hero-sub { opacity:.88; font-size:.9rem; margin-top:.25rem; }
.hero-sm { padding:1.05rem 1.35rem; margin-bottom:1rem; }
.hero-sm .hero-date { font-size:1.35rem; }

.sec-head { display:flex; align-items:center; gap:.65rem; margin:.1rem 0 .5rem; }
.sec-num { display:inline-flex; align-items:center; justify-content:center;
  width:1.65rem; height:1.65rem; border-radius:50%; background:var(--accent);
  color:#fff; font-size:.85rem; font-weight:600; flex:0 0 auto; }
.sec-title { font-size:1.15rem; font-weight:650; color:var(--ink); }
.sec-sub { color:var(--muted); font-size:.85rem; margin:-.3rem 0 .5rem 2.3rem; }

.stButton > button, .stDownloadButton > button {
  border-radius:10px; font-weight:600; }
.stButton > button[kind="primary"] {
  box-shadow:0 4px 14px rgba(63,91,139,.28); }
[data-baseweb="input"], [data-baseweb="textarea"], [data-baseweb="select"] > div {
  border-radius:10px !important; }
[data-testid="stAlert"] { border-radius:10px; }
[data-testid="stExpander"] { border-radius:12px; border:1px solid var(--line); }
.stTabs [data-baseweb="tab-list"] { gap:.35rem; }
.stTabs [data-baseweb="tab"] {
  border-radius:8px 8px 0 0; padding:.5rem 1rem; font-weight:600; }
[data-testid="stVerticalBlockBorderWrapper"] {
  border-radius:14px; border:1px solid var(--line);
  box-shadow:0 1px 3px rgba(16,24,40,.04); }
</style>
"""


def _inject_css():
    st.markdown(_CSS, unsafe_allow_html=True)


def section_header(num, title, subtitle=""):
    sub = f'<div class="sec-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div class="sec-head"><span class="sec-num">{num}</span>'
        f'<span class="sec-title">{title}</span></div>{sub}',
        unsafe_allow_html=True)


def _fetch_into(ref_key, target_key):
    """Button callback: fetch 出處 from bible.fhl.net into the verse text box."""
    info_key, err_key = target_key + "__info", target_key + "__err"
    ref = (st.session_state.get(ref_key) or "").strip()
    if not ref:
        st.session_state[err_key] = "請先輸入出處，例如「詩篇 34:1-3」"
        st.session_state[info_key] = ""
        return
    try:
        text, info = bible.fetch_reference(ref)
    except Exception as exc:
        st.session_state[err_key] = f"載入失敗：{exc}"
        st.session_state[info_key] = ""
        return
    st.session_state[target_key] = text
    st.session_state[err_key] = ""
    st.session_state[info_key] = f"已從聖經網載入：{info}"


def _pick_book(select_key, ref_key):
    """Book dropdown callback: append the picked book's FHL abbreviation to the
    出處 field (multi-book entry), separated by `；` from any existing text."""
    name = st.session_state.get(select_key)
    if not name:
        return
    code = bible.book_code(name) or name
    prev = (st.session_state.get(ref_key) or "").strip().rstrip("；;")
    st.session_state[ref_key] = f"{prev}；{code}" if prev else code
    st.session_state[select_key] = None


def _fetch_status(name):
    err = sv(name + "__err", "")
    if err:
        st.error(err)
        return
    info = sv(name + "__info", "")
    if info:
        st.caption("✅ " + info)


_inject_css()

# --------------------------------------------------------------- sidebar --
with st.sidebar:
    st.markdown('<div class="brand">崇拜投影片'
                '<span class="brand-tail"> 編輯器</span></div>',
                unsafe_allow_html=True)

    dates = sorted(store.all().keys(), reverse=True)
    idx = dates.index(DATE) if DATE in dates else 0
    picked = st.selectbox("主日日期", dates or ["2026.09.20"],
                          index=idx if dates else 0)
    if dates and picked != DATE:
        set_active(picked)

    if st.button("💾 儲存目前內容", use_container_width=True,
                 key=K("save_sidebar")):
        week_to_save = assemble_week(DATE)
        if not DATE_RE.match(week_to_save["date"]):
            st.error("日期格式應為 2026.09.20")
        else:
            store.save(week_to_save)
            st.cache_data.clear()
            st.toast(f"已儲存 {week_to_save['date']}")

    gen_clicked = st.button("📽 製成整場投影片（合併＋編譯）", type="primary",
                            use_container_width=True, key=K("generate"))
    gen_slot = st.container()

    st.divider()
    st.markdown('<div class="sec-title" style="font-size:1rem">📜 經文投影片'
                '<span class="brand-tail">（獨立流程）</span></div>',
                unsafe_allow_html=True)
    st.caption("只把「經文（讀經）」內容製成 pptx；\n"
               "不屬於整場崇拜投影片流程，也不會被合併使用。")
    scr_clicked = st.button("📜 製成經文（讀經）投影片",
                            use_container_width=True,
                            key=K("generate_scripture"))
    scr_slot = st.container()

    with st.expander("＋ 新增週次"):
        new_date = st.text_input("日期（YYYY.MM.DD）",
                                 placeholder="2026.10.04", key=K("new_week_date"))
        if st.button("新增並開啟", use_container_width=True,
                     key=K("new_week_add")):
            if DATE_RE.match(new_date):
                if new_date not in (dates or []):
                    st.cache_data.clear()
                set_active(new_date)
            else:
                st.error("日期格式應為 2026.09.20")

# ------------------------------------------------------------------ tabs --
tab_edit, tab_build = st.tabs(["✏️ 編輯內容", "📽 製作成投影片"])

with tab_edit:
    st.markdown(
        f'<div class="hero"><div class="hero-date">主日：{DATE}</div>'
        f'<div class="hero-sub">崇拜投影片編輯器 · 內容自動儲存至 Google 試算表'
        f'</div></div>', unsafe_allow_html=True)

    with st.container(border=True):
        col1, col2 = st.columns(2)
        with col1:
            section_header("1", "宣召經文", "貼上整段經文，自動分頁")
            sb, rb = st.columns([2, 3], vertical_alignment="bottom")
            sb.selectbox("書卷", bible.BOOK_NAMES, index=None,
                         placeholder="選擇書卷", key=K("psalm_book"),
                         on_change=_pick_book,
                         args=(K("psalm_book"), K("psalm_ref")))
            rb.button("從 fhl.net 聖經網輸入", type="primary",
                      key=K("psalm_load"), use_container_width=True,
                      on_click=_fetch_into,
                      args=(K("psalm_ref"), K("psalm_verses")))
            st.text_input("出處", value=sv("psalm_ref"), key=K("psalm_ref"),
                          placeholder="詩篇 34:1-3")
            st.caption("支援多卷書：以 `；` 分隔（如 詩篇 34:1-3；馬太福音 6:9-13）")
            c1, c2 = st.columns(2)
            c1.number_input("出處字型(pt)", value=48, min_value=20,
                            max_value=90, key=K("psalm_ref_size"))
            c2.number_input("內文字型(pt)", value=44, min_value=44,
                            max_value=90, key=K("psalm_font_size"))
            st.text_area("內容（貼上全部經文，自動分頁）",
                         value=sv("psalm_verses"),
                         key=K("psalm_verses"), height=160)
            _fetch_status("psalm_verses")

        with col2:
            section_header("3", "經文（讀經）", "貼上整段經文，自動分頁")
            sb, rb = st.columns([2, 3], vertical_alignment="bottom")
            sb.selectbox("書卷", bible.BOOK_NAMES, index=None,
                         placeholder="選擇書卷", key=K("scripture_book"),
                         on_change=_pick_book,
                         args=(K("scripture_book"), K("scripture_ref")))
            rb.button("從 fhl.net 聖經網輸入", type="primary",
                      key=K("scripture_load"), use_container_width=True,
                      on_click=_fetch_into,
                      args=(K("scripture_ref"), K("scripture_verses")))
            st.text_input("出處", value=sv("scripture_ref"),
                          key=K("scripture_ref"),
                          placeholder="羅馬書 12:1-8（和合本）")
            st.caption("支援多卷書：以 `；` 分隔（如 詩篇 34:1-3；馬太福音 6:9-13）")
            c1, c2 = st.columns(2)
            c1.number_input("出處字型(pt)", value=48, min_value=20,
                            max_value=90, key=K("scripture_ref_size"))
            c2.number_input("內文字型(pt)", value=44, min_value=44,
                            max_value=90, key=K("scripture_font_size"))
            st.text_area("內容（貼上全部經文，自動分頁）",
                         value=sv("scripture_verses"),
                         key=K("scripture_verses"), height=220)
            _fetch_status("scripture_verses")

    with st.container(border=True):
        section_header("2", "詩歌敬拜（Hymns）")
        st.caption("如已有現成的詩歌 .pptx，也可直接在「📽 製作成投影片」頁上載"
                   "合併（該節以檔案為準，字型／項目符號沿用來源檔）。")
        c1, c2, c3 = st.columns(3)
        c1.number_input("歌詞字體(pt)", value=44, min_value=20, max_value=90,
                        key=K("hymn_font_size"))
        c2.number_input("最小字體(pt)", value=36, min_value=20, max_value=80,
                        key=K("hymn_font_min"))
        c3.number_input("左右邊距(吋)", value=0.83, min_value=0.2, max_value=2.0,
                        step=0.05, key=K("hymn_margin_in"))
        st.caption("每首詩歌的背景圖依序取自 media/song_bg1.jpg…（可於製成時替換）")

        for i, h in enumerate(_get_ui("hymns")):
            uid = h["_id"]
            with st.expander(f"詩歌 {i + 1} · {h.get('title') or '（未命名）'}",
                             expanded=(i == 0)):
                c1, c2 = st.columns(2)
                c1.text_input("曲名",
                              value=sv(f"hymn_{uid}_title", h.get("title")),
                              key=K(f"hymn_{uid}_title"))
                c2.text_input("英文名",
                              value=sv(f"hymn_{uid}_subtitle",
                                       h.get("subtitle")),
                              key=K(f"hymn_{uid}_subtitle"))
                st.text_input("來源",
                              value=sv(f"hymn_{uid}_source", h.get("source")),
                              key=K(f"hymn_{uid}_source"))
                st.text_area("副歌（每行一句，可留空）",
                             value=sv(f"hymn_{uid}_refrain", h.get("refrain")),
                             key=K(f"hymn_{uid}_refrain"), height=60)
                st.checkbox("每段後重覆副歌",
                            value=sv(f"hymn_{uid}_repeat",
                                     h.get("repeat", True)),
                            key=K(f"hymn_{uid}_repeat"))
                for j in range(max(len(h.get("stanzas", ["", ""])), 1)):
                    st.text_area(f"第 {j + 1} 段（每行一句）",
                                 value=sv(f"hymn_{uid}_verses_{j}",
                                          h["stanzas"][j]
                                          if j < len(h["stanzas"]) else ""),
                                 key=K(f"hymn_{uid}_verses_{j}"), height=90)
                c1, c2 = st.columns(2)
                if c1.button("＋ 新增一段", key=K(f"hymn_{uid}_addstanza")):
                    h["stanzas"].append("")
                if c2.button("移除這首詩歌", key=K(f"hymn_{uid}_del")):
                    _del_item("hymns", i)
                    st.rerun()
        if st.button("＋ 新增詩歌", key=K("add_hymn")):
            _add_item("hymns", {"_id": uuid.uuid4().hex[:8], "title": "",
                                "subtitle": "", "source": "", "refrain": "",
                                "repeat": True, "stanzas": [""]})
            st.rerun()

    with st.container(border=True):
        section_header("4", "講道信息（Sermon）")
        st.info("此節只接受現成 .pptx 檔案：請於「📽 製作成投影片」頁上載，"
                "製成時以該檔案為準（此頁無需輸入）。")

    with st.container(border=True):
        section_header("5", "公佈事項（Announcements）")
        st.info("此節只接受現成 .pptx 檔案：請於「📽 製作成投影片」頁上載，"
                "製成時以該檔案為準（此頁無需輸入）。")

    with st.container(border=True):
        section_header("6", "聖餐（Communion）")
        st.checkbox("聖餐（第一主日：加聖餐＋使徒信經頁）",
                    value=sv("communion", is_first_sunday(DATE)),
                    key=K("communion"))
        if is_first_sunday(DATE):
            st.caption("此日期為當月第一主日，已預設勾選聖餐＋使徒信經。")

with tab_build:
    st.markdown(
        f'<div class="hero hero-sm"><div class="hero-date">製作成投影片</div>'
        f'<div class="hero-sub">{DATE} · 上載要合併的檔案，或以網頁內容現場編譯'
        f'</div></div>', unsafe_allow_html=True)

    with st.container(border=True):
        section_header("A", "上載要合併的檔案", "留空的節以網頁內容編譯，可混用")
        c1, c2, c3 = st.columns(3)
        songs_pptx = c1.file_uploader("詩歌敬拜 .pptx（上載合併）",
                                      type=["pptx"], key=K("build_songs"))
        sermon_pptx = c2.file_uploader("講道信息 .pptx（上載合併）",
                                       type=["pptx"], key=K("build_sermon"))
        ann_pptx = c3.file_uploader("公佈事項 .pptx（上載合併）",
                                    type=["pptx"], key=K("build_ann"))
        st.caption("上載後該節以檔案為準（合併）；留空則用編輯頁內容編譯（可混用）。")
        with st.expander("🔎 或輸入檔名（本機執行時，從 app 資料夾 / "
                         "data/downloads / ~/Downloads 找到即合併）"):
            fc1, fc2, fc3 = st.columns(3)
            songs_name = fc1.text_input(
                "詩歌檔名", placeholder="songs_slides_2026.09.20.pptx",
                key=K("build_songs_name"))
            sermon_name = fc2.text_input(
                "講道檔名", placeholder="sermon_2026.09.20.pptx",
                key=K("build_sermon_name"))
            ann_name = fc3.text_input(
                "公佈事項檔名", placeholder="announcements_2026.09.20.pptx",
                key=K("build_ann_name"))
            st.caption("上載優先；兩者皆無則該節用編輯頁內容編譯；"
                       "檔名找不到會自動回退並提示。")

    # Persist uploaded section decks so the section status survives a page
    # refresh (Streamlit clears file_uploader state).
    persist_uploads(DATE, {"songs": songs_pptx, "sermon": sermon_pptx,
                           "announcements": ann_pptx})

    with st.container(border=True):
        section_header("B", "個別產生獨立 pptx", "可下載，也會自動儲存供整場合併")
        sec_cols = st.columns(3)
        sec_jobs = [
            ("🎵 產生詩歌投影片", "hymns", "songs",
             f"songs_slides_{DATE}.pptx"),
            ("🗣 產生講道投影片", "sermon", "sermon",
             f"sermon_{DATE}.pptx"),
            ("📋 產生公佈事項投影片", "announcements", "announcements",
             f"announcements_{DATE}.pptx"),
        ]
        for col, (label, build_sec, save_sec, fname) in zip(sec_cols, sec_jobs):
            if col.button(label, use_container_width=True):
                data, _count = build_section_pptx_bytes(build_sec,
                                                        assemble_week(DATE))
                st.session_state["sec_bytes"] = data
                st.session_state["sec_name"] = fname
                st.session_state["sec_date"] = DATE
                path = save_section_deck(save_sec, DATE, data)
                st.toast(f"已產生並儲存 {os.path.basename(path)}"
                         if path else "已產生")

        sec_bytes = st.session_state.get("sec_bytes")
        if sec_bytes and st.session_state.get("sec_date") == DATE:
            st.download_button(
                "⬇️ 下載此節 pptx",
                data=sec_bytes,
                file_name=st.session_state.get("sec_name"),
                mime="application/vnd.openxmlformats-officedocument"
                     ".presentationml.presentation",
            )
            st.caption("已儲存的節會自動用於「製成整場投影片」，除非在下方另行"
                       "上載／輸入檔名；可用下方開關停用。")

    with st.expander("替換節目的背景圖片（可選）"):
        overrides = {}
        for slot in IMAGE_SLOTS:
            up = st.file_uploader(f"{slot}", type=["jpg", "jpeg", "png", "webp"],
                                  key=K(f"build_img_{slot}"),
                                  label_visibility="collapsed",
                                  accept_multiple_files=False)
            if up is not None:
                overrides[slot] = up

    ui_ann = _get_ui("ann")
    ann_images = {}
    if ui_ann:
        with st.expander("公佈事項圖片（圖片型公佈，可選）"):
            for i, a in enumerate(ui_ann):
                if a.get("mode") == "image":
                    up = st.file_uploader(
                        f"公佈 {i + 1} 圖片：{a.get('image_name') or '(未上載)'}",
                        type=["jpg", "jpeg", "png", "webp"],
                        key=K(f"build_annimg_{i}"), accept_multiple_files=False)
                    if up is not None:
                        ann_images[i] = up

    with st.container(border=True):
        section_header("C", "製成選項與來源狀態")
        use_saved = st.checkbox(
            "自動使用先前儲存／已產生的各節投影片", value=True,
            key=K("use_saved"),
            help="同一日期先「個別產生」過的節，製成整場時會自動合併；"
                 "取消則全部用網頁編輯內容現場編譯（除非上載或輸入檔名）。")
        for label, up, nm, sec in (
                ("🎵 詩歌敬拜", songs_pptx, songs_name, "songs"),
                ("🗣 講道信息", sermon_pptx, sermon_name, "sermon"),
                ("📋 公佈事項", ann_pptx, ann_name, "announcements")):
            text, kind = section_source(up, nm, sec, DATE, use_saved)
            icon = {"ready": "✅", "saved": "💾", "compile": "🧩"}[kind]
            st.markdown(f"{icon} **{label}** — {text}")
        if sv("communion", is_first_sunday(DATE)):
            st.markdown("✅ **聖餐＋使徒信經** — 已加入")
        else:
            st.markdown("⬜ **聖餐＋使徒信經** — 未加入"
                        + ("（第一主日建議勾選）"
                           if is_first_sunday(DATE) else ""))

# --- generate/download are processed last (uploads/options exist by now) but
# rendered back into the sidebar slot defined next to the button, so the
# download button appears right under it without scrolling --
with gen_slot:
    if gen_clicked:
        missing = []

        def pick(upload, name, label):
            if upload is not None:
                return upload
            if name and resolve_deck_path(name):
                return name
            if name and str(name).strip():
                missing.append(f"{label}：{str(name).strip()} 找不到")
            return None

        try:
            week = assemble_week(DATE)
            data, slide_count = build_deck(
                week,
                songs_pptx=pick(songs_pptx, songs_name, "詩歌檔名"),
                sermon_pptx=pick(sermon_pptx, sermon_name, "講道檔名"),
                announcements_pptx=pick(ann_pptx, ann_name, "公佈事項檔名"),
                image_overrides=overrides,
                announcement_images=ann_images,
                use_saved=use_saved,
            )
            st.session_state["built_bytes"] = data
            st.session_state["built_name"] = f"Sunday_Service_{DATE}.pptx"
            st.session_state["built_count"] = slide_count
            st.session_state["built_date"] = DATE
            st.success(f"完成：{slide_count} 頁")
            if missing:
                st.warning("以下檔名找不到，已改用已儲存檔案或網頁內容編譯："
                           + "；".join(missing))
        except Exception as exc:
            st.error(f"製成失敗：{exc}")

    built = st.session_state.get("built_bytes")
    if built and st.session_state.get("built_date") == DATE:
        st.download_button(
            "⬇️ 下載 PPTX",
            data=built,
            file_name=st.session_state.get("built_name"),
            mime="application/vnd.openxmlformats-officedocument"
                 ".presentationml.presentation",
        )
        st.caption(f"預估 {st.session_state.get('built_count', '?')} 頁；"
                   "檔案在瀏覽器下載，不會儲存在伺服器。")

with scr_slot:
    if scr_clicked:
        try:
            data, scr_count = build_section_pptx_bytes(
                "scripture", assemble_week(DATE))
            st.session_state["scr_bytes"] = data
            st.session_state["scr_name"] = f"scripture_{DATE}.pptx"
            st.session_state["scr_date"] = DATE
            st.session_state["scr_count"] = scr_count
            st.success(f"完成：{scr_count} 頁（獨立流程，只含讀經經文）")
        except Exception as exc:
            st.error(f"製成失敗：{exc}")

    scr = st.session_state.get("scr_bytes")
    if scr and st.session_state.get("scr_date") == DATE:
        st.download_button(
            "⬇️ 下載經文投影片 pptx",
            data=scr,
            file_name=st.session_state.get("scr_name"),
            mime="application/vnd.openxmlformats-officedocument"
                 ".presentationml.presentation",
        )
        st.caption("此獨立檔案不會被「製成整場投影片」讀取或合併。")

with tab_build:
    st.caption("提示：投影片即時製成，無需本機架設伺服器；"
               "內容儲存在 Google 試算表，不靠伺服器磁碟。"
               "儲存與製成按鈕在左側面板。")