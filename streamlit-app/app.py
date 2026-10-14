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
import difflib
import hashlib
import io
import os
import re
import tempfile
import uuid
import html

import streamlit as st

st.set_page_config(page_title="崇拜投影片編輯器", page_icon="✝",
                   layout="wide")

import bible  # noqa: E402
import login_log  # noqa: E402
import lyric_search  # noqa: E402
import onedrive  # noqa: E402
from generate_deck import (IMAGE_SLIDES,  # noqa: E402
                           READING_BG, SERMON_GOLD, _DIVIDER_TITLES)
from deck_builder import (build_deck,  # noqa: E402
                          build_section_pptx_bytes, clear_section_deck,
                          build_hymn_pptx_bytes, hymn_preview,
                          clear_upload_name, clear_video, load_upload_names,
                          prepend_video_slide, resolve_deck_path,
                          save_section_deck, save_upload_name,
                          save_video_meta, save_video_mp4,
                          saved_section_path, saved_video_meta,
                          saved_video_mp4, section_path)
from video_deck import (announcements_from_pptx,  # noqa: E402
                        build_announcements_mp4, build_images_mp4,
                        build_pptx_slides_mp4)
from model import (DATE_RE, READING_LAYOUT_DEFAULT,  # noqa: E402
                    READING_LAYOUT_OPTIONS, default_week, is_first_sunday,
                    normalize_week, split_lines)
from store import get_store, invalidate  # noqa: E402

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


# --- 登入門（簡單密碼鎖）----------------------------------------------------
# 密碼放在 `.streamlit/secrets.toml` 的 `app_password`（被 .gitignore 排除，
# 部署時請在 Streamlit Cloud 的 Settings → Secrets 貼上同一欄位）。
# 沒有設定 app_password 時會放行，但顯示黃色提醒。
# 在本機（loopback：localhost / 127.0.0.1 / ::1）開啟時直接放行，見 `_is_local`。

try:
    APP_PASSWORD = st.secrets.get("app_password") or None
except Exception:
    APP_PASSWORD = None


def _auth_ok():
    return bool(st.session_state.get("auth_ok", False))


_LOOPBACK_HOSTS = {"localhost", "127.0.0.1", "::1"}
# Headers a proxy / load balancer adds.  Their presence means the request did
# not come straight from a browser on this machine.
_FORWARD_HEADERS = ("x-forwarded-for", "x-forwarded-host", "x-forwarded-proto",
                    "x-real-ip", "forwarded")


def _host_only(value):
    """``example.com:443`` / ``user@host`` -> lowercase ``example.com``."""
    value = (value or "").rsplit("@", 1)[-1]
    return value.rsplit(":", 1)[0].strip("[]").lower()


def _is_local():
    """True when the browser is on this machine, so the password gate is skipped.

    This deliberately does NOT use ``st.context.ip_address``: behind a proxy or
    load balancer — which is exactly what Streamlit Community Cloud is — that
    value is always ``None`` (Streamlit never enables Tornado's ``xheaders``),
    so ``None`` does not mean "loopback" there.  Treating it as local is what
    silently disabled the password gate on the deployed app.  Instead we trust
    only the browser-set ``Host`` / ``Origin`` headers, and refuse to call a
    request local if any forwarding header shows a proxy is in the way.  This is
    a convenience, not a security boundary — the headers can be spoofed.
    """
    try:
        headers = {str(k).lower(): str(v)
                   for k, v in dict(st.context.headers or {}).items()}
    except Exception:
        return False
    if any(headers.get(h) for h in _FORWARD_HEADERS):
        return False
    if _host_only(headers.get("host")) not in _LOOPBACK_HOSTS:
        return False
    origin = headers.get("origin") or ""
    if origin and _host_only(origin.split("://", 1)[-1].split("/", 1)[0]) \
            not in _LOOPBACK_HOSTS:
        return False
    return True


def sv(name, default=""):
    return st.session_state.get(K(name), default)


def _song_title(uid, h):
    """A hymn's name for display: what the user has typed, else the stored one.

    The 曲名 box is a widget, so the sheet's copy goes stale the moment someone
    types — anything naming the song (expander header, delete prompt) has to read
    the live value or it keeps saying （未命名）.
    """
    return sv(f"hymn_{uid}_title", h.get("title", ""))


def _request_meta():
    """Best-effort remote IP + user agent for the login log."""
    try:
        h = dict(st.context.headers or {})
    except Exception:
        return {"ip": "", "ua": ""}
    ip = (h.get("X-Forwarded-For") or h.get("X-Real-IP")
          or h.get("Remote-Addr") or "")
    return {"ip": ip, "ua": h.get("User-Agent", "")}


def _log_action(label, detail=""):
    """Record one important action in the current login session's summary
    row of the "sessions" worksheet (Google Sheet)."""
    sess = st.session_state.get("auth_session")
    if not sess:
        return
    try:
        login_log.log_action(sess, label, detail, secrets=st_secrets)
    except Exception:
        pass


if not _auth_ok():
    if _is_local():
        # On the presenter's own machine (loopback): no password needed.
        st.session_state["auth_ok"] = True
    elif APP_PASSWORD is None:
        # A password is required off-loopback, so an unconfigured deployment is
        # locked rather than left open.  Set `app_password` in the deployment's
        # secrets (Streamlit Cloud: Settings → Secrets).
        with st.container(border=True):
            st.subheader("🔒 崇拜投影片編輯器")
            st.error("此 App 未設定 `app_password`，為安全起見已停用。請在 "
                     "Streamlit Cloud 的 Settings → Secrets 加入 "
                     "`app_password = \"你的密碼\"`，儲存後重新啟動 App。")
        st.stop()
    else:
        with st.container(border=True):
            st.subheader("🔒 崇拜投影片編輯器 — 請登入")
            with st.form(K("gate_form"), clear_on_submit=False):
                st.text_input("密碼", type="password", key=K("gate_password"),
                              placeholder="請輸入密碼後按 Enter 或「登入」")
                submitted = st.form_submit_button("登入", type="primary")
        if submitted:
            meta = _request_meta()
            ok = sv("gate_password") == APP_PASSWORD
            login_log.append_login("success" if ok else "failed", **meta,
                                   secrets=st_secrets)
            if ok:
                sess = uuid.uuid4().hex[:8]
                st.session_state["auth_session"] = sess
                login_log.new_session(sess, ip=meta.get("ip", ""),
                                      ua=meta.get("ua", ""), secrets=st_secrets)
                st.session_state["auth_ok"] = True
                st.rerun()
            else:
                st.error("密碼錯誤，請重試。")
        st.stop()


def song_bg_files():
    """media/song_bg*.jpg relative paths, sorted numerically (bg1<bg2<…)."""
    import glob
    files = sorted(glob.glob(os.path.join(os.getcwd(), "media", "song_bg*.jpg")))
    return [os.path.relpath(f, os.getcwd()) for f in files]


def song_bg_options(current):
    """Selectbox choices: default '依序' plus each media background; the
    current value is always present so saved choices never go missing."""
    opts = [""] + song_bg_files()
    if current not in opts:
        opts.append(current)
    return opts


def _bg_picker(wkey, default):
    """背景圖縮圖點選 + 上傳存檔 + 預覽（詩歌敬拜／詩歌回應共用）。

    Each background is picked with the checkbox under its thumbnail; the picked
    file (or the 依序 default, ``""``) is stored under ``K(wkey)`` — the same
    slot the old selectbox used, so nothing else reads it differently.
    """
    current = sv(wkey, default)
    opts = song_bg_options(current)

    def _choose(opt):
        st.session_state[K(wkey)] = opt
        # Radio behaviour: un-tick every other background so only the picked
        # one stays checked.  Runs as an `on_change` callback — that is, at the
        # clicked checkbox's position in this run — so writing the *other* (or
        # already-created) checkbox keys here is legal; they re-render on the
        # next pass with their updated stored values.
        for j, o in enumerate(opts):
            if o != opt:
                st.session_state[K(f"{wkey}_bgck_{j}")] = False

    thumb_cols = st.columns(len(opts) if opts else 1)
    for i, (col, opt) in enumerate(zip(thumb_cols, opts)):
        with col:
            chosen = opt == current
            cand = ""
            if opt:
                cand = opt if os.path.isfile(opt) else \
                    os.path.join(os.getcwd(), opt)
            name = ("依序 song_bg1…" if not opt
                    else os.path.splitext(os.path.basename(opt))[0])
            if os.path.isfile(cand):
                st.image(cand, width=84, caption=("✓ " + name if chosen else name))
            else:
                st.markdown(
                    '<div style="display:flex;align-items:center;'
                    'justify-content:center;height:63px;border:'
                    + ("2px solid var(--accent);" if chosen else "1px dashed var(--line);")
                    + 'border-radius:8px;color:'
                    + ("var(--accent);font-weight:700;" if chosen else "var(--muted);")
                    + 'font-size:.75rem;text-align:center">依序<br>song_bg1…'
                    '</div>', unsafe_allow_html=True)
            # Keyed checkboxes: the auto-generated id of two identical unkeyed
            # ones collides, so each gets a key; the callback keeps the row
            # acting like a radio (see `_choose`).
            st.checkbox("✓ 使用中" if chosen else "選用",
                        value=chosen,
                        key=K(f"{wkey}_bgck_{i}"),
                        on_change=_choose, args=(opt,))
    st.caption("點縮圖下方的勾選框即套用該背景（第一格「依序」＝照既定順序，"
               "同一個背景只會勾一個）。")
    up = st.file_uploader("上傳新背景（jpg，存到 media/song_bg*.jpg，"
                          "可供所有詩歌選擇）",
                          type=["jpg", "jpeg", "png"],
                          key=K(wkey + "_bgupload"))
    if up is not None:
        ext = ".png" if (up.name or "").lower().endswith(".png") else ".jpg"
        path = os.path.join(os.getcwd(), "media",
                            "song_bg_%s%s" % (uuid.uuid4().hex[:8], ext))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as fh:
            fh.write(up.getbuffer())
        st.session_state[K(wkey)] = os.path.relpath(path, os.getcwd())
        st.rerun()
    st.caption("亦可直接放圖片進 media/（檔名 song_bg*.jpg 就會出現在縮圖列）；"
               "上傳的圖片只存本機，要放上 Cloud 請把 media/ 一起提交到 Git。")
    if current:
        cand = current if os.path.isfile(current) else \
            os.path.join(os.getcwd(), current)
        if os.path.isfile(cand):
            st.image(cand, caption="背景預覽：" + os.path.basename(cand),
                     width=260)
    return current


def _uploaded_name(upload):
    """Display name of an UploadedFile, or None."""
    if upload is None:
        return None
    return (getattr(upload, "name", "") or "").strip() or None


def _deck_present(section, upload_key=None):
    """(text, present) describing the deck a section already has, or (None, False)
    when it has none.  Names the file it actually came from: an upload's own
    name, else the name recorded when it was fetched from OneDrive/Google Drive,
    else the local save name.  The fetched name is preferred because that is the
    file the user picked in the cloud folder; the local name is a canonical
    re-write of it."""
    wkey = "ann" if section == "announcements" else section
    up = st.session_state.get(K(upload_key)) if upload_key else None
    if up is not None:
        fname = _uploaded_name(up)
        return ("已上載檔案（合併）" + (f"：**{fname}**" if fname else ""), True)
    name = st.session_state.get(K(f"build_{wkey}_name"))
    path = resolve_deck_path(name) if name else None
    if not path:
        path = saved_section_path(section, DATE)
    if not path:
        return None, False
    orig = (load_upload_names(DATE) or {}).get(section)
    shown = orig or os.path.basename(path)
    where = "已從雲端讀取" if orig else "已儲存"
    return f"{where}：**{shown}**（製成整場時以該檔案為準）", True


def _cloud_deck_status(section, label, upload_key=None):
    """Status row for a section that may already have a deck, with a ✖ button to
    clear it.  Returns True when a deck is present, so the caller can skip its
    plain "nothing here yet" caption."""
    text, present = _deck_present(section, upload_key)
    if not present:
        return False
    line = f"✅ **{label}** — {text}"
    if not sv("use_saved", True):
        line += "　⚠️「合併先前各步驟投影片檔」已取消，實際不會用到此檔案。"
    r1, r2 = st.columns([0.76, 0.24], vertical_alignment="center")
    r1.markdown(line)
    r2.button("✖ 清除", key=K(f"clr_{section}"), use_container_width=True,
              on_click=_signal_clear, args=(DATE, section),
              help=f"刪除{label}已讀取／上載的 pptx 與記錄，"
                   "讓你可以重新從雲端讀取或改用網頁內容編譯")
    return True


def section_deck_panel(sec_key, build_sec, save_sec, label, dl_name,
                       btn=None, widget=None):
    """Section A uploader + section B '產生獨立 pptx' button, moved into the
    wizard steps. Uses the same session keys the merge flow reads
    (`build_<widget>`, `sec_bytes_<sec_key>`), so 來源狀態 and 製成整場投影片
    keep working regardless of which tab/step is active."""
    st.file_uploader(
        f"上載 {label} .pptx（上載合併；留空則用網頁編輯內容編譯）",
        type=["pptx"], key=K("build_" + (widget or sec_key)))
    if not btn:
        return
    bc1, bc2 = st.columns(2)
    if bc1.button(btn, use_container_width=True, key=K(f"genppt_{sec_key}")):
        try:
            data, _count = build_section_pptx_bytes(
                build_sec, assemble_week(DATE))
            path = save_section_deck(save_sec, DATE, data)
            st.session_state[K(f"sec_bytes_{sec_key}")] = data
            st.session_state[K(f"sec_date_{sec_key}")] = DATE
            st.toast("已產生並儲存 "
                     + (os.path.basename(path) if path else "已產生"))
            _log_action("產生並儲存 " + label, DATE)
        except Exception as exc:
            st.error(f"產生失敗：{exc}")
    b = st.session_state.get(K(f"sec_bytes_{sec_key}"))
    if b and st.session_state.get(K(f"sec_date_{sec_key}")) == DATE:
        bc2.download_button(
            f"⬇️ 下載 {label} .pptx", data=b, file_name=dl_name,
            mime="application/vnd.openxmlformats-officedocument"
                 ".presentationml.presentation",
            use_container_width=True)


def _od_config():
    """[onedrive] block from secrets, or None when not configured."""
    try:
        cfg = dict(st.secrets.get("onedrive", {}) or {})
    except Exception:
        cfg = {}
    if not (cfg.get("tenant_id") and cfg.get("client_id")
            and cfg.get("client_secret")):
        return None
    return cfg


def _od_fetch_section(section, date, cfg, picked_file=None):
    """Save a section deck from the OneDrive date-subfolder to the saved-deck
    path.  When `picked_file` ({id,name}) is given it is downloaded directly;
    otherwise fall back to the canonical name, then any .pptx containing a
    date spelling.  Records the original filename like the Drive path.
    Returns the file name, or None."""
    want = os.path.basename(section_path(section, date) or "")
    if not want:
        raise RuntimeError("不支援此節（無對應檔名）")
    picked = picked_file if (
        picked_file and picked_file.get("id") and picked_file.get("name")) else None
    if picked is None:
        files = onedrive.files_in_date_folder(
            cfg, cfg.get("folder_path", ""), date)
        picked = next((f for f in files if (f.get("name") or "") == want), None)
        if picked is None:
            variants = onedrive._date_variants(date)
            picked = next(
                (f for f in files
                 if (f.get("name") or "").endswith(".pptx")
                 and any(v in (f.get("name") or "") for v in variants)), None)
    if picked is None:
        return None
    path = section_path(section, date)
    ok, err = onedrive.download_to(cfg, picked["id"], path)
    if not ok:
        raise RuntimeError(err)
    fname = picked.get("name") or ""
    st.session_state[f"{date}::build_{'ann' if section == 'announcements' else section}_name"] = \
        fname
    save_upload_name(section, date, fname)
    st.cache_data.clear()
    return fname


def _od_fetch_video(cfg, picked_file=None):
    """Download the date's 家事MP4 into the persisted video store.  When
    `picked_file` ({id,name}) is given it is downloaded directly; otherwise
    fall back to announcements_<date>.mp4, then any .mp4 containing a date
    spelling.  Returns the file name, or None."""
    picked = picked_file if (
        picked_file and picked_file.get("id") and picked_file.get("name")) else None
    if picked is None:
        want = f"announcements_{DATE}.mp4"
        files = onedrive.files_in_date_folder(
            cfg, cfg.get("folder_path", ""), DATE)
        picked = next((f for f in files if (f.get("name") or "") == want), None)
        if picked is None:
            variants = onedrive._date_variants(DATE)
            picked = next(
                (f for f in files
                 if (f.get("name") or "").lower().endswith(".mp4")
                 and any(v in (f.get("name") or "") for v in variants)), None)
    if picked is None:
        return None
    tmp = tempfile.NamedTemporaryFile(
        suffix=".mp4", prefix="odrive_", delete=False).name
    try:
        ok, err = onedrive.download_to(cfg, picked["id"], tmp)
        if not ok:
            raise RuntimeError(err)
        with open(tmp, "rb") as fh:
            data = fh.read()
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass
    save_video_mp4(DATE, data)
    st.session_state[K("video_data")] = data
    st.cache_data.clear()
    return picked.get("name")


def _odrive_panel(section, label, is_video=False):
    """Expander that lists the section decks (or the 家事MP4 when is_video) in
    the configured OneDrive/SharePoint folder and lets the user pick which file
    to download.  Reusable across wizard steps."""
    with st.expander("🌐 從 OneDrive 讀取%s投影片" % label):
        cfg = _od_config()
        if not cfg:
            st.info("未設定 OneDrive（secrets.toml 的 [onedrive] 區塊："
                    "tenant_id / client_id / client_secret / site_host / "
                    "site_path / folder_path）。設定後即可從雲端讀取。")
            return
        _p = (DATE or "").split(".")
        _mmdd = ("%s.%s%s" % (_p[0], _p[1], _p[2])
                 if len(_p) == 3 else DATE)
        want = (f"announcements_{DATE}.mp4" if is_video
                else os.path.basename(section_path(section, DATE) or ""))
        ext = ".mp4" if is_video else ".pptx"
        if is_video:
            st.caption("📽 家事MP4 會從 OneDrive **`%s`** 資料夾（含所有子資料夾）"
                       "搜尋 **.mp4**；請從下方清單選擇檔案後按「讀取並儲存」。"
                       "預期名稱：**`%s`**。"
                       % (cfg.get("folder_path", ""), want))
        else:
            st.caption("📽 %s投影片會直接從 OneDrive **`%s`** 資料夾"
                       "（依日期子資料夾，如 `%s`）讀取；請從下方清單選擇檔案後"
                       "按「讀取並儲存」。預期名稱：**`%s`**。"
                       % (label, cfg.get("folder_path", ""), _mmdd, want))
        try:
            if is_video:
                # 家事MP4 可能放在任何子資料夾：搜尋整個 tree，日期夾優先。
                cand = onedrive.find_mp4s(cfg, cfg.get("folder_path", ""))
                _sub = onedrive._date_subfolder(
                    cfg, cfg.get("folder_path", ""), DATE)
                if _sub:
                    _datep = ((cfg.get("folder_path", "") or "").rstrip("/")
                              + "/" + _sub["name"]).strip("/")
                    cand.sort(key=lambda f: 0 if (f.get("path") or "").strip("/")
                              .startswith(_datep + "/") else 1)
            else:
                cand = onedrive.files_in_date_folder(
                    cfg, cfg.get("folder_path", ""), DATE)
        except Exception as exc:
            st.error(f"OneDrive 讀取失敗：{exc}")
            return
        # Display label: 名稱；影片在日期夾外時附上位置。
        labels = []
        for f in cand:
            n = f.get("name", "")
            p = (f.get("path") or "").strip("/")
            if is_video and _sub:
                _datep = ((cfg.get("folder_path", "") or "").rstrip("/")
                          + "/" + _sub["name"]).strip("/")
                if p.startswith(_datep + "/") or p == _datep:
                    labels.append(n)
                else:
                    labels.append("%s（%s）" % (n, p or "根目錄"))
            else:
                labels.append(n)
        # Preselect: exact canonical name, then any name containing the date.
        default_i = None
        if want:
            for i, l in enumerate(labels):
                if l == want or l.startswith(want + "（"):
                    default_i = i
                    break
        if default_i is None:
            variants = onedrive._date_variants(DATE)
            for i, l in enumerate(labels):
                base = l.split("（", 1)[0]
                if base.endswith(ext) and any(v in base for v in variants):
                    default_i = i
                    break
        if default_i is None and labels:
            default_i = 0
        if labels:
            picked_i = st.radio(
                "選擇要下載的檔案：",
                range(len(labels)),
                index=default_i if default_i is not None else 0,
                format_func=lambda i: labels[i],
                key=K(f"odrive_pick_{section}"))
            if is_video and cand[picked_i].get("path"):
                st.caption("位置：`%s`"
                           % (cand[picked_i].get("path") or "根目錄"))
            picked = cand[picked_i]
        else:
            picked = None
            if is_video:
                st.info("OneDrive **`%s`** 資料夾中目前找不到任何 "
                        "**MP4** 檔（含子資料夾）。上載家事影片後即可在此選擇。"
                        % cfg.get("folder_path", ""))
            else:
                st.info("該日期資料夾中沒有可下載的簡報檔。")
        if st.button("⚡ 從 OneDrive 讀取並儲存", type="primary",
                     key=K(f"odrive_fetch_{section}")):
            if not picked:
                st.warning("沒有可下載的檔案。")
                return
            try:
                if is_video:
                    name = _od_fetch_video(cfg, picked)
                else:
                    name = _od_fetch_section(section, DATE, cfg, picked)
            except Exception as exc:
                st.error(f"OneDrive 讀取失敗：{exc}")
            else:
                if name:
                    st.success(f"已從 OneDrive 讀取：{name}")
                    _log_action("OneDrive 讀取 " + label, name)
                else:
                    st.warning("下載未完成，請再試一次。")


def _finalize_stamp(key, label):
    """Checkbox that 'stamps' a section as finalized (persists with the week)."""
    st.markdown(
        "<style>"
        '[data-testid="stCheckbox"] [data-testid="stWidgetLabel"] p,'
        '[data-testid="stAlert"] p {'
        "font-size:1.25rem;"
        "}</style>",
        unsafe_allow_html=True)
    st.write("")
    marked = st.checkbox(f"✔️ 標記「{label}」此節為定稿（蓋章）",
                         value=sv(key, False), key=K(key))
    if marked:
        st.success(f"📌「{label}」已標記為定稿 —— 內容以目前編輯為準。")


def parse_slide_numbers(text):
    """Parse comma-separated (and/or dash-ranged) 1-based slide numbers into
    a list, preserving input order and dropping duplicates/invalid parts."""
    out, seen = [], set()
    for part in str(text or "").split(","):
        part = part.strip().replace("–", "-")
        if not part or "-" not in part:
            try:
                i = int(part)
            except (TypeError, ValueError):
                continue
            if i not in seen:
                seen.add(i)
                out.append(i)
            continue
        a, _, b = part.partition("-")
        try:
            lo, hi = int(a), int(b)
        except (TypeError, ValueError):
            continue
        if hi < lo:
            lo, hi = hi, lo
        for i in range(lo, hi + 1):
            if i not in seen:
                seen.add(i)
                out.append(i)
    return out


class ZipSlideImage:
    """A slide image inside an uploaded ZIP: keeps the slide number parsed
    from its filename plus the raw bytes for MP4 encoding, so the user's
    投影片編號 input matches the numbers in the filenames."""

    __slots__ = ("num", "name", "_blob")

    def __init__(self, num, name, blob):
        self.num, self.name, self._blob = num, name, blob

    def getvalue(self):
        return self._blob


def video_poster_png(mp4_bytes):
    """Extract the MP4's first frame as PNG poster bytes; falls back to a
    navy placeholder so embedding never fails on a codec hiccup."""
    try:
        import imageio.v2 as iio
        import numpy as np
        from PIL import Image
        reader = iio.get_reader(io.BytesIO(mp4_bytes))
        frame = reader.get_data(0)
        reader.close()
        img = Image.fromarray(np.asarray(frame)).convert("RGB")
    except Exception:
        img = None
    if img is None:
        from PIL import Image as _Image
        img = _Image.new("RGB", (1920, 1080), (17, 36, 74))
    bio = io.BytesIO()
    img.save(bio, "PNG")
    return bio.getvalue()


def build_video_bytes(src_kind, src_items, zip_map, src_pptx_bytes,
                      nums, secs, fpsv, ann_images=None, progress=None):
    """Rebuild the announcements MP4 from the section-D source state.
    Returns (mp4_bytes, None) on success or (None, error_message)."""
    if src_kind == "zip":
        sel = [n for n in parse_slide_numbers(nums) if n in zip_map]
        bad = [n for n in parse_slide_numbers(nums) if n not in zip_map]
    else:
        sel = [n for n in parse_slide_numbers(nums)
               if 1 <= n <= len(src_items)]
        bad = [n for n in parse_slide_numbers(nums)
               if not 1 <= n <= len(src_items)]
    if not parse_slide_numbers(nums):
        return None, "請輸入至少一個編號（例如 1,3,5）。"
    if bad:
        return None, "以下編號沒有對應的投影片：%s" % ", ".join(map(str, bad))
    try:
        if src_kind == "images":
            data = build_images_mp4(
                [src_items[i - 1].getvalue() for i in sel],
                secs_per_slide=secs, fps=fpsv, progress=progress)
        elif src_kind == "zip":
            data = build_images_mp4(
                [zip_map[n] for n in sel],
                secs_per_slide=secs, fps=fpsv, progress=progress)
        elif src_kind == "pptx":
            data = build_pptx_slides_mp4(
                src_pptx_bytes, [n - 1 for n in sel],
                secs_per_slide=secs, fps=fpsv, progress=progress)
        else:
            zero = [n - 1 for n in sel]
            imgs = {
                z: (ann_images[z].getvalue()
                    if ann_images and z in ann_images else None)
                for z in zero}
            data = build_announcements_mp4(
                src_items, zero, secs_per_slide=secs, fps=fpsv, images=imgs,
                progress=progress)
        return data, None
    except Exception as exc:
        return None, str(exc)


def date_from_filename(name):
    """Find a YYYY.MM.DD / YYYY-MM-DD token anywhere in a filename (prefix or
    suffix, e.g. `announcements_2026.09.20.pptx`), returning a normalized
    `YYYY.MM.DD` string, or None when the name carries no date."""
    m = re.search(r"(\d{4})[.\-](\d{1,2})[.\-](\d{1,2})", str(name or ""))
    if not m:
        return None
    y, mo, d = (int(x) for x in m.groups())
    return f"{y:04d}.{mo:02d}.{d:02d}"


def persist_uploads(date, uploads):
    """Save freshly uploaded section decks to data/decks/<date>/ so the
    section status survives a page refresh (Streamlit clears file_uploader
    state on reload). `uploads` maps section -> UploadedFile|None.
    The original uploaded filename is remembered in the `build_<widget>_name`
    key so the status page keeps showing it (instead of the canonical saved
    name) even after the file has been persisted."""
    for sec, up in uploads.items():
        if up is None:
            continue
        sig = (getattr(up, "name", ""), getattr(up, "size", None))
        if st.session_state.get(f"{date}::uploadsig_{sec}") != sig:
            save_section_deck(sec, date, up.getvalue())
            st.session_state[f"{date}::uploadsig_{sec}"] = sig
            wkey = "ann" if sec == "announcements" else sec
            st.session_state[f"{date}::build_{wkey}_name"] = \
                getattr(up, "name", "") or ""
            save_upload_name(sec, date, getattr(up, "name", "") or "")
            _log_action("上載 " + sec, getattr(up, "name", "") or "")


def _signal_clear(date, section):
    """Callback for the ✖ buttons in 製成選項與來源狀態: schedules clearing the
    section's uploaded/specified/saved deck on the next run."""
    st.session_state[f"{date}::clear_{section}"] = True


def clear_date_state(date):
    prefix = f"{date}::"
    for key in [k for k in st.session_state.keys() if k.startswith(prefix)]:
        del st.session_state[key]


def set_active(date, seed_week=None):
    clear_date_state(date)
    week = normalize_week(seed_week or store.get(date) or default_week(date))
    st.session_state["active_date"] = date
    st.session_state[f"{date}::ui_hymns"] = _seed_hymns(week)
    st.session_state[f"{date}::ui_slides"] = _seed_slides(week)
    st.session_state[f"{date}::ui_ann"] = _seed_ann(week)
    _seed_fields(date, week)
    _seed_response_fields(date, week)
    st.rerun()


def _seed_hymns(week):
    hymns = week.get("hymns") or []
    if not hymns:
        # 預設：詩歌敬拜 3 首、歌詞空白。
        hymns = [{} for _ in range(3)]
    out = []
    for h in hymns:
        out.append({
            "_id": uuid.uuid4().hex[:8],
            "title": h.get("title", ""),
            "subtitle": h.get("subtitle", ""),
            "source": h.get("source", ""),
            "music": h.get("music", ""),
            "lyricist": h.get("lyricist", ""),
            "bg": h.get("bg", ""),
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


def _seed_values(week, date):
    psalm = week.get("psalm") or {}
    scripture = week.get("scripture") or {}
    sermon = week.get("sermon") or {}
    response = week.get("response") or {}
    return {
        "psalm_ref": psalm.get("ref", ""),
        "psalm_ref_size": psalm.get("ref_size", 48),
        "psalm_font_size": psalm.get("font_size", 44),
        "psalm_layout": psalm.get("layout", READING_LAYOUT_DEFAULT),
        "psalm_verses": "\n".join(psalm.get("verses") or []),
        "hymn_font_max": week.get("hymn_font_max", 52),
        "hymn_font_min": week.get("hymn_font_min", 48),
        "hymn_margin_in": week.get("hymn_margin_in", 0.83),
        "hymn_text_shadow": bool(week.get("hymn_text_shadow", True)),
        "call_font": week.get("call_font", "DFKai-SB") or "DFKai-SB",
        "hymns_finalized": bool(week.get("hymns_finalized", False)),
        "scripture_ref": scripture.get("ref", ""),
        "scripture_ref_size": scripture.get("ref_size", 48),
        "scripture_font_size": scripture.get("font_size", 44),
        "scripture_layout": scripture.get("layout", READING_LAYOUT_DEFAULT),
        "scripture_verses": "\n".join(scripture.get("verses") or []),
        "sermon_title": sermon.get("title", ""),
        "sermon_title_2": sermon.get("title_2", ""),
        "sermon_speaker": sermon.get("speaker", ""),
        "sermon_date": sermon.get("date", date),
        "sermon_ref": sermon.get("scripture", ""),
        "sermon_ref_size": sermon.get("ref_size", 28),
        "communion": bool(week.get("communion", False)),
        "resp_title": response.get("title", ""),
        "resp_subtitle": response.get("subtitle", ""),
        "resp_source": response.get("source", ""),
        "resp_music": response.get("music", ""),
        "resp_lyricist": response.get("lyricist", ""),
        "resp_bg": response.get("bg", ""),
        "resp_refrain": "\n".join(response.get("refrain") or []),
        "resp_repeat": bool(response.get("refrain_after_every_verse", True)),
        "resp_finalized": bool(response.get("finalized", False)),
    }


def _seed_fields(date, week):
    """Seed the scalar editor widgets (refs, font sizes, verse text areas,
    sermon meta, communion) from a stored/default week.  Without this, saved
    宣召經文／經文（讀經）等欄位開檔時顯示空白，再儲存就會被清空。"""
    for name, val in _seed_values(week, date).items():
        st.session_state[f"{date}::{name}"] = val


def _draft_for(date):
    """Plain (non-widget) dict holding the authoritative field values for `date`.

    Editor widgets are re-created on every rerun, and a widget declared with both
    `value=` and `key=` has its session key reset to that `value` each time. So
    the widget key cannot be trusted to remember an edit once its step has been
    left — this dict is what does. It is not a widget key, so Streamlit never
    touches it, and it lives under the `f"{date}::"` prefix so clear_date_state()
    drops it when switching dates."""
    return st.session_state.setdefault(f"{date}::_draft", {})


def _lapse(key, draft, stored):
    """Keep one editor field's value alive across the step it is rendered on.

    The draft is authoritative. Each run the widget key is forced back to the
    draft, which makes the widget's own `value=` re-seed a no-op, so leaving a
    step and returning shows the edit the user made rather than the stored
    default. A value that *differs* from the draft can only have come from the
    user interacting with the live widget, so that is what gets recorded.

    `stored` — the last saved value — seeds fields the user has not touched."""
    if key in draft:
        current = st.session_state.get(key, draft[key])
        if current != draft[key]:
            draft[key] = current
        st.session_state[key] = draft[key]
    else:
        draft[key] = st.session_state.setdefault(key, stored)


def _clear_drafts(date):
    """Drop the unsaved-edit draft for `date` (after a save, or a forced
    re-read of the store) so later restores fall back to stored data."""
    st.session_state.pop(f"{date}::_draft", None)


def _restore_lapsed_fields(date, week):
    """Keep every editor field's value consistent as steps come and go.

    Editor widgets are rebuilt on each rerun and any widget declared with both
    `value=` and `key=` has its session key reset to that `value`, so a field
    whose step is off-screen falls back to the last *saved* value. That silently
    discarded unsaved edits — ticking 聖餐, stepping to 宣召經文 and back left it
    unticked — and a field whose step is not active would otherwise read as
    empty and wipe the stored week on save. _lapse() runs each field through the
    per-date draft, which holds the authoritative value: live edit if the user
    just touched it, otherwise the draft, otherwise stored data."""
    draft = _draft_for(date)
    for name, val in _seed_values(week, date).items():
        _lapse(f"{date}::{name}", draft, val)
    hymns = week.get("hymns") or []
    for i, h in enumerate(st.session_state.get(f"{date}::ui_hymns") or []):
        sw = hymns[i] if i < len(hymns) else {}
        base = f"{date}::hymn_{h['_id']}_"
        _lapse(base + "title", draft, sw.get("title", ""))
        _lapse(base + "subtitle", draft, sw.get("subtitle", ""))
        _lapse(base + "source", draft, sw.get("source", ""))
        _lapse(base + "music", draft, sw.get("music", ""))
        _lapse(base + "lyricist", draft, sw.get("lyricist", ""))
        _lapse(base + "bg", draft, sw.get("bg", ""))
        _lapse(base + "refrain", draft, "\n".join(sw.get("refrain") or []))
        _lapse(base + "repeat", draft,
               bool(sw.get("refrain_after_every_verse", True)))
        _lapse(base + "paste", draft, _stanzas_to_text(sw.get("verses", [])))
    slides = (week.get("sermon") or {}).get("slides") or []
    for i, s in enumerate(st.session_state.get(f"{date}::ui_slides") or []):
        sw = slides[i] if i < len(slides) else {}
        base = f"{date}::slide_{s['_id']}_"
        _lapse(base + "title", draft, sw.get("title", ""))
        _lapse(base + "body", draft, "\n".join(sw.get("body") or []))
    rverses = (week.get("response") or {}).get("verses") or []
    _lapse(f"{date}::resp_paste", draft, _stanzas_to_text(rverses))


def _slide_blocks(text):
    """Editor text -> 段 (slide) blocks: one blank line separates slides.

    The box is the single lyric editor, so this split is applied live — the
    same protocol the old 匯入為各段 enforced, minus the chorus detection which
    is no longer part of the UI.  Blank blocks are dropped; an empty or
    whitespace-only box yields one empty block.
    """
    blocks, cur = [], []
    for ln in (text or "").split("\n"):
        if ln.strip():
            cur.append(ln.rstrip())
        elif cur:
            blocks.append("\n".join(cur))
            cur = []
    if cur:
        blocks.append("\n".join(cur))
    return blocks or [""]


def _stanzas_to_text(stanzas):
    """Stored 段 list (list of line-lists) -> the editor's blank-line separated
    text, so the box and the stored week round-trip without loss."""
    filled = ["\n".join(s) for s in (stanzas or [])
              if s and any((ln or "").strip() for ln in s)]
    return "\n\n".join(filled)


def _seed_response_fields(date, week):
    """Seed the 詩歌回應 lyric editor from the stored week."""
    response = week.get("response") or {}
    verses = ["\n".join(v) for v in (response.get("verses") or [])] or [""]
    st.session_state[f"{date}::ui_resp"] = verses
    st.session_state.setdefault(
        f"{date}::resp_paste",
        _stanzas_to_text(response.get("verses") or []))


@st.cache_data(show_spinner=False, max_entries=64)
def _cached_week(date):
    """Current stored/default week for a date (cached; invalidated on save)."""
    return normalize_week(store.get(date) or default_week(date))


week = _cached_week(DATE)
if f"{DATE}::ui_hymns" not in st.session_state:
    st.session_state[f"{DATE}::ui_hymns"] = _seed_hymns(week)
    st.session_state[f"{DATE}::ui_slides"] = _seed_slides(week)
    st.session_state[f"{DATE}::ui_ann"] = _seed_ann(week)
    _seed_fields(DATE, week)
    _seed_response_fields(DATE, week)
    # Pre-seed the per-hymn scalar widget keys so their values survive the
    # wizard stepping away from the hymns step and back (widgets re-render
    # with these as their session values).
    for h in st.session_state[f"{DATE}::ui_hymns"]:
        uid = h["_id"]
        base = f"{DATE}::hymn_{uid}_"
        st.session_state[base + "title"] = h.get("title", "")
        st.session_state[base + "subtitle"] = h.get("subtitle", "")
        st.session_state[base + "source"] = h.get("source", "")
        st.session_state[base + "music"] = h.get("music", "")
        st.session_state[base + "lyricist"] = h.get("lyricist", "")
        st.session_state[base + "bg"] = h.get("bg", "")
        st.session_state[base + "refrain"] = h.get("refrain", "")
        st.session_state[base + "repeat"] = bool(h.get("repeat", True))
        st.session_state[base + "paste"] = _stanzas_to_text(
            h.get("stanzas", []))

_restore_lapsed_fields(DATE, week)


# ── clear-upload signals ────────────────────────────────────────────────────
# The ✖ buttons in 製成選項與來源狀態 set these flags; they are processed here
# at the top of the run — before the file_uploader widgets are instantiated —
# so their session keys can be safely popped and the page rerun.
_SECTION_KEYS = ("songs", "offering", "response", "sermon", "announcements",
                 "psalm", "scripture", "communion")
_cleared = False
for _sec in _SECTION_KEYS:
    _wkey = "ann" if _sec == "announcements" else _sec
    if st.session_state.get(f"{DATE}::clear_{_sec}"):
        st.session_state.pop(K("build_" + _wkey), None)
        st.session_state.pop(K("build_" + _wkey + "_name"), None)
        st.session_state.pop(f"{DATE}::uploadsig_{_sec}", None)
        st.session_state.pop(f"{DATE}::clear_{_sec}", None)
        clear_section_deck(_sec, DATE)
        clear_upload_name(_sec, DATE)
        _cleared = True
if st.session_state.get(f"{DATE}::clear_video"):
    st.session_state.pop(f"{DATE}::clear_video", None)
    st.session_state.pop(K("video_data"), None)
    st.session_state.pop(K("video_first_deck_s9"), None)
    st.session_state.pop(K("video_first_deck"), None)
    clear_video(DATE)
    _cleared = True
if _cleared:
    st.rerun()

# After a fresh login, session state is empty but the saved decks and their
# original filenames survive on disk. Repopulate the `build_<widget>_name`
# keys from the durable record so the status page keeps showing the uploaded
# filename instead of the canonical default.
_src_names = load_upload_names(DATE)
for _sec in _SECTION_KEYS:
    _wkey = "ann" if _sec == "announcements" else _sec
    if (_src_names.get(_sec) and
            K("build_" + _wkey + "_name") not in st.session_state and
            saved_section_path(_sec, DATE)):
        st.session_state[K("build_" + _wkey + "_name")] = _src_names[_sec]


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


@st.dialog("移除詩歌")
def _confirm_del_hymn():
    """Confirmation modal before removing a song (removal is permanent).
    A button click inside the dialog auto-closes it on rerun."""
    idx = st.session_state.get(K("del_hymn_pending"), -1)
    ui = _get_ui("hymns")
    if not 0 <= idx < len(ui):
        st.session_state.pop(K("del_hymn_pending"), None)
        return
    h = ui[idx]
    st.markdown(f"確定要移除 **{_song_title(h['_id'], h) or '（未命名）'}** 嗎？"
                "此操作不會被還原。")
    d1, d2 = st.columns(2)
    if d1.button("確認移除", type="primary", use_container_width=True):
        _del_item("hymns", idx)
        st.session_state.pop(K("del_hymn_pending"), None)
    if d2.button("取消", use_container_width=True):
        st.session_state.pop(K("del_hymn_pending"), None)


# ---------------------------------------------------------------- assemble --
def assemble_week(date):
    Kp = lambda n: f"{date}::{n}"  # noqa: E731

    def gv(n, d):
        return st.session_state.get(Kp(n), d)

    ui_hymns = st.session_state.get(f"{date}::ui_hymns", [])
    ui_slides = st.session_state.get(f"{date}::ui_slides", [])
    ui_ann = st.session_state.get(f"{date}::ui_ann", [])
    ui_resp = st.session_state.get(f"{date}::ui_resp", [""])

    resp_paste = st.session_state.get(f"{date}::resp_paste")
    if resp_paste is None:
        resp_verses = [split_lines(v) for v in ui_resp if v and v.strip()]
    else:
        resp_verses = [split_lines(b) for b in _slide_blocks(resp_paste)]

    week = {
        "date": date,
        "psalm": {
            "ref": gv("psalm_ref", ""),
            "ref_size": int(gv("psalm_ref_size", 48)),
            "font_size": int(gv("psalm_font_size", 44)),
            "layout": str(gv("psalm_layout", READING_LAYOUT_DEFAULT)
                          or READING_LAYOUT_DEFAULT),
            "verses": split_lines(gv("psalm_verses", "")),
        },
        "hymn_font_max": int(gv("hymn_font_max", 52)),
        "hymn_font_min": int(gv("hymn_font_min", 48)),
        "hymn_margin_in": float(gv("hymn_margin_in", 0.83)),
        "hymn_text_shadow": bool(gv("hymn_text_shadow", True)),
        "call_font": str(gv("call_font", "DFKai-SB") or "DFKai-SB"),
        "hymns_finalized": bool(gv("hymns_finalized", False)),
        "hymns": [assemble_hymn(date, h) for h in ui_hymns],
        "scripture": {
            "ref": gv("scripture_ref", ""),
            "ref_size": int(gv("scripture_ref_size", 48)),
            "font_size": int(gv("scripture_font_size", 44)),
            "max_lines": int(gv("scripture_max_lines", 4)),
            "layout": str(gv("scripture_layout", READING_LAYOUT_DEFAULT)
                          or READING_LAYOUT_DEFAULT),
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
        "response": {
            "title": gv("resp_title", ""),
            "subtitle": gv("resp_subtitle", ""),
            "source": gv("resp_source", ""),
            "music": gv("resp_music", ""),
            "lyricist": gv("resp_lyricist", ""),
            "bg": gv("resp_bg", ""),
            "refrain": None,
            "refrain_after_every_verse": False,
            "verses": resp_verses,
            "finalized": bool(gv("resp_finalized", False)),
        },
        "communion": bool(gv("communion", False)),
    }
    return normalize_week(week)


def assemble_hymn(date, ui):
    base = f"{date}::hymn_{ui['_id']}_"
    paste = st.session_state.get(base + "paste")
    if paste is None:
        paste = _stanzas_to_text(ui.get("stanzas", []))
    return {
        "title": st.session_state.get(base + "title", ui.get("title", "")),
        "subtitle": st.session_state.get(base + "subtitle", ui.get("subtitle", "")),
        "source": st.session_state.get(base + "source", ui.get("source", "")),
        "music": st.session_state.get(base + "music", ui.get("music", "")),
        "lyricist": st.session_state.get(base + "lyricist",
                                        ui.get("lyricist", "")),
        "bg": st.session_state.get(base + "bg", ui.get("bg", "")),
        "refrain": None,
        "refrain_after_every_verse": False,
        "verses": [split_lines(b) for b in _slide_blocks(paste)],
    }


def _parse_hymn_paste(text, with_chorus=True):
    """Split a whole-song paste into (stanzas, refrain, repeat).

    Protocol: blocks are lyric groups separated by blank lines —
      * 1 blank line → each block becomes a separate 段 (slide)
      * 2 blank lines before a block → that block is the 副歌 (chorus)

    Lines with no blank gap belong to the same block, so a 導歌＋副歌 kept
    together reads as one chorus unit.  If the paste marks nothing as 副歌,
    falls back to treating the most repeated block (or a shared trailing
    tail) as the 副歌.  A leading verse number on a block (e.g. `1 至聖的
    主…`, `2. 你縱尊貴…`) is stripped from its first line."""
    blocks, before = [], []
    cur, blanks, started = None, 0, False
    for ln in (text or "").split("\n"):
        if ln.strip():
            if cur is not None:
                cur.append(ln.rstrip())
            else:
                before.append(2 if blanks >= 2 else 1 if blanks == 1 else 0)
                cur = [ln.rstrip()]
                blanks = 0
                started = True
        else:
            if cur is not None:
                blocks.append(cur)
                cur = None
            if started:
                blanks += 1
    if cur is not None:
        blocks.append(cur)

    for b in blocks:
        m = re.match(r"^\s*(\d{1,3})[.．、:）)\s]\s*", b[0])
        if m:
            b[0] = b[0][m.end():]

    if not with_chorus:
        return ["\n".join(b) for b in blocks] or [""], "", False

    def std(b):
        return tuple(" ".join(x.split()) for x in b)

    if not blocks:
        return [""], "", True
    if len(blocks) == 1:
        return ["\n".join(blocks[0])], "", True

    marked = [i for i in range(len(blocks)) if before[i] >= 2]
    if marked:
        chorus_i = marked[0]
        stanzas = ["\n".join(b) for i, b in enumerate(blocks)
                   if i not in marked]
        refrain = "\n".join(blocks[chorus_i])
        repeat = True
        return (stanzas or [""]), refrain, repeat

    counts = {}
    for b in blocks:
        counts[std(b)] = counts.get(std(b), 0) + 1
    chorus = [b for b in set(counts) if counts[b] > 1]
    if chorus:
        key = max(chorus, key=lambda k: (counts[k], len(k)))
        keep = [b for b in blocks if std(b) != key]
        refrain = "\n".join([b for b in blocks if std(b) == key][0])
        repeat = True
        return (["\n".join(b) for b in keep] or [""]), refrain, repeat

    tail_len = max((i for i in range(min(len(b) for b in blocks), 0, -1)
                    if len({std(b[-i:]) for b in blocks}) == 1), default=0)
    if tail_len >= 2:
        refrain = "\n".join(blocks[0][-tail_len:])
        keep = [b[:-tail_len] for b in blocks]
        return (["\n".join(b) for b in keep if b] or [""]), refrain, True

    return ["\n".join(b) for b in blocks], "", True


def _ly_base(uid_kind, uid):
    """Widget-key prefix for a song editor: per-song uid, or the shared 詩歌回應."""
    return f"hymn_{uid}_" if uid_kind == "hymn" else "resp_"


def _focus_hymn(uid, uids):
    """Summary-card 「前往編輯」: open one song's expander and close the rest.

    Runs as an `on_click` callback, i.e. before the next pass builds anything, so
    writing the expanders' keys here is legal.  The expanders are keyed
    (`hymn_exp_<uid>`) and track state (`on_change="rerun"`), so the stored value
    wins over the default and the panel really opens.
    """
    for u in uids:
        st.session_state[K(f"hymn_exp_{u}")] = (u == uid)
    h = next((x for x in _get_ui("hymns") if x.get("_id") == uid), None)
    name = (_song_title(uid, h) if h is not None else "") or "（未命名）"
    st.toast(f"已展開：{name}")


def _song_badge(uid, h):
    """One-glance status for a song header / summary card: how many 段, or the
    first thing still missing.  Reads the live editor text so it tracks typing."""
    title = sv(f"hymn_{uid}_title", h.get("title", "")) or ""
    stanzas = _slide_blocks(sv(f"hymn_{uid}_paste", ""))
    if not title.strip():
        return "⚠ 未填曲名"
    if not any(b.strip() for b in stanzas):
        return "⚠ 沒有歌詞"
    return f"{len(stanzas)} 段"


_TITLE_FIELDS = (("title", "曲名"), ("subtitle", "英文名"), ("source", "詩集"),
                 ("music", "曲"), ("lyricist", "詞"))


def _ly_fill_paste(uid_kind, uid, text, n_seg, info=None):
    """Drop a library song into the 歌詞 editor box — its single editing surface.

    Library pages often lead with a verse number (`1. …`) and tab indents; the
    import strips the number from each page's first line (the same cleaning the
    old 匯入為各段 did) so the box shows exactly what lands on the slides.
    ``info`` (a title card read from the library song) fills the
    曲名/英文名/詩集/曲/詞 boxes below; empty fields are left as they were.

    No `st.rerun()` on purpose: the search panel sits above the editor box and
    the metadata boxes, so they render later in this same pass and pick the
    values up on the way down.
    """
    base = _ly_base(uid_kind, uid)
    stanzas, _refrain, _repeat = _parse_hymn_paste(text, with_chorus=False)
    st.session_state[K(base + "paste")] = "\n\n".join(stanzas)
    filled = []
    for field, label in _TITLE_FIELDS:
        val = (info or {}).get(field, "")
        if val:
            st.session_state[K(base + field)] = val
            filled.append(label)
    usable = [s for s in stanzas if s.strip()]
    msg = (f"已匯入 {n_seg} 頁 → **{len(usable)} 段**到歌詞編輯框。"
           "可直接在框內繼續編輯：每段之間 1 個空行＝分 slide。")
    if info:
        msg += "第一頁是標題卡，已用來填寫下方欄位，不計入段數。"
    if filled:
        msg += "已填寫：" + "、".join(filled) + "。"
    st.success(msg)


def _lyric_editor(uid_kind, uid):
    """Single lyric editor, shown directly in the card.

    One blank line separates slides; each line inside a slide is one slide line.
    The box's session key is preserved across wizard steps by
    `_restore_lapsed_fields` (like every other editor field), so edits survive
    switching between sections.  The count line reflects the live blank-line
    split as you type.
    """
    base = _ly_base(uid_kind, uid)
    key = K(base + "paste")
    st.text_area(
        "歌詞編輯（每段之間 1 個空行＝分 slide · 每行＝一句）",
        value=sv(base + "paste", ""),
        key=key, height=360,
        placeholder="貼上整首歌詞（或按上方「歌詞庫搜尋」匯入）…",
        help="每個空行分隔的一段＝1 張 slide；段內每行＝畫面上的一句。"
             "直接在框內增刪、調整行序即可，不需按任何匯入按鈕。")
    stanzas = _slide_blocks(sv(base + "paste", ""))
    usable = [b for b in stanzas if b.strip()]
    if usable:
        st.caption(f"共 **{len(usable)} 段**（每段＝1 張 slide）")
    else:
        st.caption("尚未輸入歌詞。每段之間留 1 個空行＝分 slide。")


# Whole-song import keeps every page but the title card: slide 1 fills the
# metadata boxes (and is skipped, since the deck builds its own title slide),
# and every other page becomes one 段.  So a source deck that re-slides its
# chorus after every verse keeps those pages.


def _lyric_open(flag_key):
    """selectbox on_change: ask for the lyrics pop-up on the next render."""
    st.session_state[flag_key] = True


# width="small" is Streamlit's narrowest dialog (a 500px max). Lyrics are short lines, so the
# 1280px of "large" left the text stranded in a mostly empty box. The only other sizes are
# "medium" (750px) and "large"; this version takes no pixel value.
@st.dialog("全首歌詞", width="small")
def _lyrics_dialog(name, where, slides, query):
    """Read-only pop-up: the whole song, one 頁 per block, query terms marked."""
    st.markdown("**%s**　<span style='color:var(--muted)'>(%s)</span>"
                % (name, where), unsafe_allow_html=True)
    parts = []
    for no, text in slides:
        parts.append(
            "<div style='color:var(--muted);font-size:.75rem;"
            "margin:.6rem 0 .15rem'>第 %d 頁</div>" % no)
        parts.append("<div style='line-height:1.75'>%s</div>"
                     % lyric_search.highlight(text, query))
    st.markdown("".join(parts), unsafe_allow_html=True)


def _lyric_search_panel(uid_kind, uid):
    """🔍 歌詞庫搜尋 — look a song up in lyrics/*/.lyrics_txt and get its lyrics.

    Results are grouped by song: a song whose chorus repeats on eight slides is
    listed once, not eight times.  Selecting a song opens a read-only pop-up
    with its whole lyrics and the query terms highlighted; 「匯入整首詩歌」 hands
    the whole song to the 貼上整首詩歌 box in the same "one blank line per 段"
    protocol that box already uses.
    """
    base = _ly_base(uid_kind, uid)
    open_key = K(f"{base}lyopen")
    try:
        lib = lyric_search.get_library()
    except Exception as exc:                       # unreadable library
        st.warning(f"歌詞庫無法載入：{exc}")
        return
    with st.expander("🔍 歌詞庫搜尋"):
        if not len(lib):
            st.info(f"歌詞庫是空的（預期在 `lyrics/*/{lyric_search.CACHE_DIRNAME}/`）。")
            return
        books = sorted({w for _n, w in lib.meta.values()})
        st.caption(f"{len(lib)} 首詩歌／{lib.total_slides} 頁，來自 "
                   f"{'、'.join(books)}。"
                   "空格分開多個詞＝同一頁要全部找到；忽略標點與空格，"
                   "所以跨行、跨頁的句子也找得到。")
        q = st.text_input("搜尋歌詞", key=K(f"{base}lyq"),
                          placeholder="例如 中國啊　或　讓我們 愛")
        q = (q or "").strip()
        if not q:
            return
        songs = lib.search_songs(q)
        if not songs:
            st.info("找不到符合的歌詞。試試少一點字，或只用其中一個詞。")
            return
        n_pages = sum(len(s.nos) for s in songs)
        cap = (f"{len(songs)} 首詩歌（找到 {n_pages} 頁）。"
               if len(songs) < lyric_search.MAX_HITS
               else f"超過 {lyric_search.MAX_HITS} 首詩歌，只顯示前面 "
                    f"{lyric_search.MAX_HITS} 首——請加詞收窄。")
        st.caption(cap)

        def _snippet_of(s):
            """First matching line of the song, for the result label."""
            for no, text, _n in lib.songs[s.key]:
                if not s.nos or no in s.nos:
                    return lyric_search.snippet(text, q)
            return ""

        # 曲名 · 詩集 · 找到頁數 — the book disambiguates two same-named hymns,
        # and the snippet shows why the row matched.
        labels = []
        for s in songs:
            why = ("找到 %d 頁" % len(s.nos)) if s.nos else "找到曲名"
            labels.append("%s · %s · %s — %s"
                          % (s.name, s.where, why, _snippet_of(s)))
        pick = st.selectbox("結果", options=list(range(len(songs))),
                            format_func=lambda i: labels[i],
                            key=K(f"{base}lysel"),
                            on_change=_lyric_open, args=(open_key,),
                            help="選一首歌即彈出全首歌詞（搜尋字詞會標示）。")
        hit = songs[pick]
        all_slides = [(no, text) for no, text, _n in lib.songs[hit.key]]
        # 匯入整首詩歌: the title card (slide 1, when there is one) fills the
        # metadata boxes and is left out of the 段, because the deck builds its
        # own title slide from those fields.  Every other page becomes one 段, so
        # a source deck that re-slides its chorus after every verse keeps them.
        info, text, n_seg = lib.import_song(hit.key)
        b1, b3 = st.columns([1.5, 1])
        if b1.button(f"⬇️ 匯入整首詩歌（{n_seg} 段）",
                     key=K(f"{base}lyall"), use_container_width=True,
                     type="primary",
                     help="把各頁歌詞貼進下方歌詞編輯框；第一頁標題卡用來填寫"
                          "曲名／英文名／詩集／曲／詞，不計入段數。"):
            _ly_fill_paste(uid_kind, uid, text, n_seg, info)
        # Opens the same pop-up as selecting a row, without a rerun: the button
        # click and the check below are in this same run, and the flag is
        # cleared as the dialog opens so a later rerun does not re-open it.
        see = b3.button("👁 看全首歌詞", key=K(f"{base}lysee"),
                        use_container_width=True,
                        help="彈出視窗顯示全首歌詞，並標示搜尋字詞。")
        st.caption(f"曲名可填：{hit.name}　（{hit.where}）")
        if see or st.session_state.get(open_key):
            st.session_state[open_key] = False
            _lyrics_dialog(hit.name, hit.where, all_slides, q)


def _hymn_preview_html(preview):
    """Inline HTML cards summarising the generated slides of one hymn."""
    cards = []
    for i, sl in enumerate(preview, 1):
        body = []
        for line, pt in sl["lines"]:
            sz = (f" <span style='color:var(--muted);font-size:.72em'>· "
                  f"{pt}pt</span>" if pt else "")
            body.append(f"<div>{html.escape(line)}{sz}</div>")
        if sl["bg"]:
            body.append("<div style='color:var(--muted);font-size:.72em'>"
                        "▲ 自訂背景圖</div>")
        if not body:
            body.append("<div style='color:var(--muted);font-size:.8rem'>"
                        "（空白）</div>")
        cards.append(
            '<div style="border:1px solid var(--line);border-radius:10px;'
            'padding:.55rem .7rem;min-width:170px;max-width:240px;'
            'flex:1 1 170px;background:#fff">'
            f'<div style="font-weight:700;font-size:.8rem;color:var(--accent);'
            f'border-bottom:1px solid var(--line);padding-bottom:.3rem;'
            f'margin-bottom:.4rem">Slide {i}</div>'
            '<div style="text-align:center">'
            + "".join(body) + "</div></div>")
    return ('<div style="display:flex;flex-wrap:wrap;gap:10px">'
            + "".join(cards) + "</div>")


def _render_hymn_preview(uid, title):
    """Show the generated single-hymn pptx preview + download (if any)."""
    pbytes = st.session_state.get(K(f"hymn_{uid}_preview_data"))
    if not pbytes:
        return
    count = st.session_state.get(K(f"hymn_{uid}_preview_count"), 0)
    st.caption(f"已產生 {count} 頁（含標題卡）。文字示意為實際投影片內容，"
               "背景／字型以開啟檔為準。")
    st.markdown(st.session_state.get(K(f"hymn_{uid}_preview_html"), ""),
                unsafe_allow_html=True)
    st.download_button(
        "⬇️ 下載這首詩歌的 pptx",
        data=pbytes,
        file_name="hymn_{}_{}.pptx".format(
            re.sub(r'[\\/:*?"<>|]+', "_", title) or uid, DATE),
        mime="application/vnd.openxmlformats-officedocument"
             ".presentationml.presentation",
        key=K(f"hymn_{uid}_preview_dl"))


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
.block-container { padding-top:2.4rem; padding-bottom:3rem; max-width:1500px; }
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
  border-radius:10px; font-weight:600; white-space:nowrap; }
[data-testid="stSidebar"] .stButton > button,
[data-testid="stSidebar"] .stDownloadButton > button {
  white-space:normal; }
.stButton > button[kind="primary"] {
  box-shadow:0 4px 14px rgba(63,91,139,.28); }
[data-baseweb="input"], [data-baseweb="textarea"], [data-baseweb="select"] > div {
  border-radius:10px !important; }
[data-testid="stAlert"] { border-radius:10px; }
[data-testid="stExpander"] { border-radius:12px; border:1px solid var(--line); }

/* Collapser headers (詩歌 N, 背景圖, 預覽這首詩歌的 slides, 歌詞庫搜尋,
   OneDrive 讀取面板, …): a thick accent outline so the row reads as a
   button; open sections fill solid accent with white text.  DOM:
   div[data-testid="stExpander"] > details > summary (stExpanderDetails
   marks the *content* box, not the header). */
div[data-testid="stExpander"] > details > summary {
  background: var(--accent-soft);
  border: 2px solid var(--accent) !important;
  border-radius: 10px;
  margin: .3rem;
  padding: .5rem .9rem !important;
  font-weight: 700;
}
div[data-testid="stExpander"] > details[open] > summary {
  background: var(--accent);
  color: #fff;
  border-color: var(--accent-2) !important;
}
div[data-testid="stExpander"] > details[open] > summary
  [data-testid="stExpanderIcon"] { color: #fff; }
.stTabs [data-baseweb="tab-list"] { gap:.35rem; }
.stTabs [data-baseweb="tab"] {
  border-radius:8px 8px 0 0; padding:.5rem 1rem; font-weight:600; }
[data-testid="stVerticalBlockBorderWrapper"] {
  border-radius:14px; border:1px solid var(--line);
  box-shadow:0 1px 3px rgba(16,24,40,.04); }
.steps { display:flex; gap:.45rem; margin:.4rem 0 .1rem; }
.steps .step { flex:1; padding:.55rem .4rem; text-align:center;
  border-radius:10px; font-size:.8rem; font-weight:700; line-height:1.25;
  background:var(--line); color:var(--muted); border:1px solid transparent; }
.steps .step b { display:block; font-size:1rem; }
.steps .step.done { background:var(--accent-soft); color:var(--accent); }
.steps .step.active { background:var(--accent); color:#fff;
  box-shadow:0 5px 14px rgba(63,91,139,.3); }
.steps .step.wait { opacity:.75; }
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


def _fetch_into(ref_key, target_key, section=None):
    """Button callback: fetch 出處 from bible.fhl.net into the verse text box.
    With `section` ("psalm"|"scripture"), also saves the fetched content to
    the store immediately so it survives an app restart without an extra
    儲存目前內容 click."""
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
    if section:
        try:
            week = normalize_week(store.get(DATE) or default_week(DATE))
            cur = dict(week.get(section) or {})
            cur["ref"] = ref
            cur["verses"] = split_lines(text)
            week[section] = cur
            store.save(week)
            _clear_drafts(DATE)
            st.cache_data.clear()
        except Exception as exc:
            st.session_state[err_key] = (
                f"已載入，但儲存到 Google Sheet 失敗：{exc}；請另按「儲存目前內容」")
            st.session_state[info_key] = ""


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


VERSION_OPTIONS = ["和合本", "和修版", "新譯本"]
VERSION_NOTE = {"和合本": "和合本",
                "和修版": "和合本2010", "新譯本": "新譯本"}

# 經文 size floors shared by the widgets and the stored-week clamp.
REF_SIZE_MIN, REF_SIZE_MAX = 20, 90
BODY_SIZE_MIN, BODY_SIZE_MAX = 44, 90

FONT_OPTIONS = [
    ("DFKai-SB", "DFKai-SB（標楷體）"),
    ("PMingLiU", "PMingLiU（新細明體）"),
    ("JhengHei", "JhengHei（正黑體）"),
    ("DengXian", "DengXian（等線體）"),
    ("LiSu", "LiSu（隸書）"),
]
FONT_LABEL = dict(FONT_OPTIONS)


def _font_radio(label, key):
    """統一的全場字型選擇（寫入 `call_font`）。"""
    st.radio(
        label,
        options=[f for f, _ in FONT_OPTIONS],
        format_func=lambda f: FONT_LABEL.get(f, f),
        key=key,
        horizontal=True)
    return sv(key.split("::")[-1], "DFKai-SB")


def _clamp_widget(key, lo, hi):
    """Pull a numeric widget's session value back into [lo, hi].

    Streamlit validates an already-set widget value against min_value/max_value
    and raises when it falls outside, so weeks saved under older (looser) bounds
    would crash the editor instead of rendering.  Rounding to int keeps
    number_input from handing back a float."""
    val = st.session_state.get(key)
    if val is None:
        return
    try:
        n = int(round(float(val)))
    except (TypeError, ValueError):
        return
    if n < lo or n > hi:
        st.session_state[key] = max(lo, min(hi, n))


def _layout_radio(key):
    """經文排版方式：整段（段落連續、節碼上標）／列表（一節一行）。"""
    value = sv(key.split("::")[-1], READING_LAYOUT_DEFAULT)
    opts = [v for v, _ in READING_LAYOUT_OPTIONS]
    try:
        idx = opts.index(value) if value in opts else 0
    except Exception:
        idx = 0
    st.radio(
        "經文排版方式",
        options=opts,
        format_func=lambda v: dict(READING_LAYOUT_OPTIONS).get(v, v),
        index=idx,
        key=key,
        horizontal=True)
    return sv(key.split("::")[-1], READING_LAYOUT_DEFAULT)


def _song_font_options(cur):
    """歌詞字級選項：52pt（預設，置頂）往下遞減至 48pt。舊週次存有範圍外的
    數值時一併補進，避免 selectbox 因值不在清單而報錯。"""
    opts = list(range(52, 47, -1))          # 52, 51, 50, 49, 48
    if cur not in opts:
        opts.append(int(cur))
    return opts


def _song_font_selector(label, hint=None):
    """歌詞字級下拉（見 _song_font_options 的選項），label 與下拉同一行並緊貼。
    若舊週次存有範圍外的數值，自動補進選項，避免 selectbox 因值不在清單而報錯。
    """
    cur = int(sv("hymn_font_max", 52) or 52)
    options = _song_font_options(cur)
    c1, c2, _ = st.columns([1.3, 0.9, 5], vertical_alignment="center")
    c1.markdown(f"**{label}**")
    c2.selectbox(
        "歌詞字級",
        options=options,
        index=options.index(cur),
        format_func=lambda v: f"{v} pt",
        key=K("hymn_font_max"),
        label_visibility="collapsed")
    if hint:
        st.caption(hint)


def _strip_ref_note(ref):
    """Remove a trailing `（版本）` / `(版本)` note from a reference."""
    s = str(ref).strip()
    for o, c in (("（", "）"), ("(", ")")):
        if s.endswith(c) and o in s:
            return s[:s.rfind(o)].strip()
    return s


def _ref_version_index(ref):
    """Index into VERSION_OPTIONS matching the note currently on the ref (0)."""
    s = str(ref).strip()
    for o, c in (("（", "）"), ("(", ")")):
        if s.endswith(c) and o in s:
            note = s[s.rfind(o) + 1:-1].strip()
            for i, opt in enumerate(VERSION_OPTIONS):
                if note == VERSION_NOTE[opt]:
                    return i
            break
    return 0


def _pick_version(select_key, ref_key, target_key=None, section=None):
    """Version radio callback: set/replace the `（版本）` note on the 出處 field
    then auto-refetch the 內容 (like the fhl.net button) in that translation."""
    label = st.session_state.get(select_key)
    if not label:
        return
    note = f"（{VERSION_NOTE[label]}）"
    ref = _strip_ref_note(st.session_state.get(ref_key) or "")
    st.session_state[ref_key] = f"{ref}{note}" if ref else ""
    st.session_state[select_key] = None
    if ref and target_key:
        _fetch_into(ref_key, target_key, section)


def _fetch_status(name):
    err = sv(name + "__err", "")
    if err:
        st.error(err)
        return
    info = sv(name + "__info", "")
    if info:
        st.caption("✅ " + info)


# Wizard steps for the 編輯內容 tab — filled top to bottom, one at a time.
def _saved_mp4_bytes(date):
    """Bytes of a previously generated 家事MP4 for `date`, or None."""
    path = saved_video_mp4(date)
    if not path:
        return None
    try:
        with open(path, "rb") as fh:
            return fh.read()
    except OSError:
        return None


def _video_result_tail(ns, vid, fname, note):
    """Download button + 併入第一頁 for a ready MP4 (uploaded or generated)."""
    st.download_button(
        "⬇️ 下載 MP4", data=vid, file_name=fname,
        mime="video/mp4", key=K(f"video_dl{ns}"))
    st.caption(note)
    first_toggle = st.checkbox(
        "把影片併入整場投影片的第一頁\n"
        "（自動循環播放，直到點擊下一頁）",
        value=True, key=K(f"video_first_deck{ns}"))
    save_video_meta(DATE, bool(first_toggle))
    if first_toggle:
        st.caption("記得按左上「製成完整投影片」的按鈕，"
                   "影片才會被放入第一頁。")


def render_video_panel(ns="", num="D",
                       title="🎞 家事分享投影片 → MP4"):
    """The section-D MP4 panel, reusable in the 「製作成投影片」 tab (D) and as
    wizard step 9 (s9).  `ns` namespaces widget keys so both panels can render
    in the same run; the produced MP4 is stored under the shared video_data key
    so the sidebar 製成 flow always sees it."""
    section_header(num, title,
                   "直接上載 PowerPoint 匯出的 MP4，或選用週期家事分享內容產製")
    st.caption("**從 PowerPoint 匯出 MP4**：開啟投影片 → 檔案 → 匯出 → "
               "建立影片（或「匯出影片」），選 1080p 並等待匯出，"
               "再把 .mp4 上載到下方「① 直接上載」。"
               "若尚未備好影片，也可選用週期家事分享內容，由 **Server 端**以 "
               "Pillow 製成 H.264 影片（純 Python，雲端可執行）。")
    ann_vsnap = st.session_state.get(K("video_snap_s9")) or {}
    ann_pptx = ((ann_vsnap.get("pptx") if ann_vsnap.get("kind") == "pptx"
                 else None) or st.session_state.get(K("build_ann")))
    ui_ann = _get_ui("ann")
    ann_images = {}
    if ui_ann:
        for i, a in enumerate(ui_ann):
            if a.get("mode") == "image":
                up = st.session_state.get(K(f"build_annimg_{i}"))
                if up is not None:
                    ann_images[i] = up
    video_mp4 = st.file_uploader(
        "① 直接上載 PowerPoint 匯出的 MP4（不需再產製）",
        type=["mp4"], key=K(f"video_mp4{ns}"))
    vdate = date_from_filename(
        getattr(ann_pptx, "name", "") if ann_pptx is not None else "")
    src_kind, src_items, src_pptx_bytes = "none", [], None
    src_fname = (getattr(ann_pptx, "name", "")
                 if ann_pptx is not None else "")
    saved_ann_path = saved_section_path("announcements", DATE)
    cur_ann = (assemble_week(DATE).get("announcements") or [])
    if video_mp4 is None:
        if ann_pptx is not None:
            try:
                src_pptx_bytes = ann_pptx.getvalue()
                src_items = announcements_from_pptx(src_pptx_bytes)
                src_kind = "pptx"
                st.caption(f"已上載 **{src_fname}**：{len(src_items)} 張，"
                           "以檔案頁面原貌產製（不受日期限制）。")
            except Exception as exc:
                st.caption(f"無法讀取已上載檔 {src_fname}：{exc}")
                src_items = []
        elif saved_ann_path:
            try:
                with open(saved_ann_path, "rb") as fh:
                    src_pptx_bytes = fh.read()
                src_items = announcements_from_pptx(src_pptx_bytes)
                src_kind = "pptx"
                st.caption(f"使用已儲存 announcements_{DATE}.pptx："
                           f"{len(src_items)} 張。")
            except Exception as exc:
                st.caption(f"無法讀取已儲存檔（announcements_{DATE}.pptx）：{exc}")
                src_items = []
        elif vdate:
            wk_video = normalize_week(store.get(vdate) or default_week(vdate))
            src_items = wk_video.get("announcements") or []
            src_kind = "content"
            if not src_items and cur_ann:
                src_items = cur_ann
                st.caption(f"檔名含日期 **{vdate}**，但該日沒有家事分享內容，"
                           f"退回目前日期 **{DATE}** 的內容。")
            else:
                st.caption(f"已上載檔名含日期 **{vdate}**（不需等於目前日期），"
                           "使用該日的家事分享內容。")
        else:
            src_items = cur_ann
            if not src_items:
                src_kind = "none"
            else:
                src_kind = "content"
                if src_fname:
                    st.caption(f"檔名未含日期（{src_fname}），使用目前日期 "
                               f"{DATE} 的內容。")
    st.session_state[K(f"video_snap{ns}")] = {
        "kind": src_kind, "items": src_items, "zip_map": {},
        "pptx": src_pptx_bytes, "vdate": vdate, "ns": ns,
        "nums_auto": ",".join(str(i + 1) for i in range(len(src_items)))}
    if video_mp4 is not None:
        vid = video_mp4.getvalue()
        st.session_state[K("video_data")] = vid
        save_video_mp4(DATE, vid)
        _video_result_tail(
            ns, vid, f"announcements_{DATE}.mp4",
            f"✅ 已採用你上載的 PowerPoint MP4（約 {len(vid) / 1048576:.1f} MB）。")
    elif src_kind == "none":
        st.caption("沒有可用的家事分享來源：請上載 PowerPoint 匯出的 MP4，"
                   "或在「✏️ 編輯內容 → 家事分享」輸入至少一項。")
        st.text_input("投影片編號", value="", disabled=True,
                      key=K(f"video_nums_off{ns}"),
                      help="上載 MP4 或輸入家事分享內容後即可產製")
    else:
        v1, v2, v3 = st.columns(3)
        nums = v1.text_input(
            "投影片編號", value=",".join(str(i + 1)
                                      for i in range(len(src_items))),
            key=K(f"video_nums{ns}"), placeholder="例如：1,3,5 或 1-3,5")
        secs = v2.number_input("每張停留秒數", 1.0, 30.0, 10.0, step=1.0,
                               key=K(f"video_secs{ns}"))
        fpsv = v3.select_slider("fps", options=[15, 20, 24, 30], value=15,
                                key=K(f"video_fps{ns}"))
        if st.button("🎞 產生影片（MP4）", key=K(f"video_gen{ns}"),
                     type="primary"):
            bar = st.progress(0.0, text="準備產製 MP4…")
            data, err = build_video_bytes(
                src_kind, src_items, {}, src_pptx_bytes,
                nums, secs, fpsv, ann_images,
                progress=lambda f: bar.progress(
                    f, text=f"產製 MP4… {f * 100:.0f}%"))
            if err:
                bar.empty()
                st.error(err)
            elif data is None:
                bar.empty()
                st.error("沒有可用的影片來源。")
            else:
                bar.progress(1.0, text="MP4 已產製完成")
                st.session_state[K("video_data")] = data
                save_video_mp4(DATE, data)
        vid = st.session_state.get(K("video_data")) or _saved_mp4_bytes(DATE)
        if vid:
            _video_result_tail(
                ns, vid, f"announcements_{vdate or DATE}.mp4",
                f"已產生影片（約 {len(vid) / 1048576:.1f} MB）。")


EDIT_STEPS = [
    ("宣召經文", "宣召經文"),
    ("詩歌敬拜（Hymns）", "詩歌"),
    ("獻詩（Offering）", "獻詩"),
    ("讀經經文", "讀經"),
    ("講道信息（Sermon）", "講道"),
    ("詩歌回應（Response Hymn）", "詩歌回應"),
    ("聖餐（Communion）", "聖餐"),
    ("家事分享（Announcements）", "家事分享"),
    ("家事分享 → MP4（影片）", "家事MP4"),
]


def _edit_step():
    try:
        return max(1, min(len(EDIT_STEPS), int(sv("edit_step", 1))))
    except (TypeError, ValueError):
        return 1


def _step_marker(step):
    st.progress(step / len(EDIT_STEPS),
                text=f"第 {step} / {len(EDIT_STEPS)} 步驟 · "
                     f"{EDIT_STEPS[step - 1][0]}")
    cols = st.columns(len(EDIT_STEPS))
    for i, (_title, short) in enumerate(EDIT_STEPS):
        idx = i + 1
        if cols[i].button(f"{idx} {short}", key=K(f"stepj_{idx}"),
                          type="primary" if idx == step else "secondary",
                          use_container_width=True):
            if idx != step:
                st.session_state[K("edit_step")] = idx
                st.rerun()
    st.caption("頂端按鈕可按**任意順序**填寫，不需要依序完成。")


def _step_nav(step):
    ncol, pcol = st.columns([1, 1])
    prev = pcol.button("⬅ 上一步", use_container_width=True,
                       disabled=(step <= 1), key=K("step_prev"))
    if step < len(EDIT_STEPS):
        nxt = ncol.button(f"下一步：{EDIT_STEPS[step][1]}", type="primary",
                          use_container_width=True, key=K("step_next"))
    else:
        nxt = ncol.button("完成 ✔（已到最後一節）", type="primary",
                          use_container_width=True, disabled=True,
                          key=K("step_next"))
    if prev:
        st.session_state[K("edit_step")] = step - 1
        st.rerun()
    if nxt:
        st.session_state[K("edit_step")] = step + 1
        st.rerun()


def _sanity_ref_norm(text):
    """Normalize a 出處 string for comparison: expand book names (詩→詩篇),
    drop any version label in parens (（和合本）), then strip whitespace and
    punctuation and lowercase.  「詩 34:1-3」→「詩篇3413」；
    「羅馬書 12:1-8（和合本）」→「羅馬書1218」."""
    s = bible.expand_book_ref(text or "")
    s = re.sub(r"[（(][^）)]*[）)]", "", s)
    s = re.sub(r"[^\w\s]", "", s, flags=re.UNICODE)
    return re.sub(r"\s+", "", s).lower()


def _sanity_find_ref(line):
    """Return the first bible-reference substring (e.g. 羅馬書 12:1-8) found
    anywhere in a slide line, or None.  Labels like 「經文：」 or quotes around
    the reference are tolerated; the longest book name wins."""
    if not line:
        return None
    for name in sorted(bible.BOOK_NAMES, key=len, reverse=True):
        idx = line.find(name)
        if idx < 0:
            continue
        rest = line[idx + len(name):]
        m = re.match(r"\s*(\d+)\s*[:：]\s*(\d+(?:[\-–—～,，]\d+)*)", rest)
        if m:
            return f"{name} {m.group(1)}:{m.group(2)}"
    return None


def _deck_slide_infos(data):
    """Read a compiled pptx back into [{bg, lines, imgs, maxsize}] for the sanity
    check: `bg` is the slide's solid background hex (or None when it is a picture
    or inherits the template), `lines` the non-empty text lines of its shapes,
    `imgs` the SHA1 hashes of its picture shapes (template pages such as
    聖餐／使徒信經 are recognised by these), and `maxsize` the largest explicit
    font size on the slide (used to spot song title cards in imported decks)."""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    pres = Presentation(io.BytesIO(data))
    out = []
    for slide in pres.slides:
        lines = []
        imgs = []
        maxsize = None
        for shp in slide.shapes:
            try:
                if shp.shape_type == MSO_SHAPE_TYPE.PICTURE \
                        and shp.image is not None:
                    imgs.append(hashlib.sha1(shp.image.blob).hexdigest())
            except Exception:
                pass
            if not shp.has_text_frame:
                continue
            for para in shp.text_frame.paragraphs:
                for r in para.runs:
                    if r.font.size is not None:
                        pt = float(r.font.size.pt)
                        if maxsize is None or pt > maxsize:
                            maxsize = pt
                t = "".join(r.text for r in para.runs).strip()
                if t:
                    lines.append(t)
        bg = None
        try:
            fill = slide.background.fill
            if fill.type is not None:
                bg = str(fill.fore_color.rgb)
        except Exception:
            bg = None
        out.append({"bg": bg, "lines": lines, "imgs": imgs, "maxsize": maxsize})
    return out


def _sanity_check_scripture(week, deck=None):
    """Sanity check: does 讀經經文 match the 講道信息 經文?

    Each side is taken from the week's content first (scripture.ref vs
    sermon.scripture).  When the generated full deck (`deck` = pptx bytes) is
    given, the refs actually rendered on the slides are read back too — the
    講道 經文 sits on one of the first four slides of the 講道信息 section —
    the slides right after the gold divider when the deck was generated here,
    otherwise the slides right after the navy 讀經 pages, so imported .pptx
    sections are covered too.  Labels like 「經文：羅馬書 12:1-8」 are
    tolerated.  Returns a report dict:
    {status: ok|contained|mismatch|missing, scr, ser, scr_src, ser_src}."""
    scr = week.get("scripture") or {}
    ser = week.get("sermon") or {}
    src = str(scr.get("ref") or "").strip()
    s_sr = str(ser.get("scripture") or "").strip()

    deck_line = {"scr": None, "ser": None}
    if deck:
        infos = _deck_slide_infos(deck)
        navy = [i for i, s in enumerate(infos) if s["bg"] == READING_BG]
        # 讀經 ref：最後一張 navy（讀經）slide 上的經文出處。
        if navy:
            for ln in infos[navy[-1]]["lines"]:
                r = _sanity_find_ref(ln)
                if r:
                    deck_line["scr"] = r
                    break
        # 講道 經文：講道章節開頭（本 App 產生的 gold divider 之後；匯入的
        # .pptx 則為讀經之後）的前四張 slide，取最先出現的經文出處。
        start = 0
        for i in range(len(infos) - 1):
            if infos[i]["bg"] == SERMON_GOLD:
                start = i + 1
                break
        if start == 0 and navy:
            start = navy[-1] + 1
        if start:
            for idx in range(start, min(start + 4, len(infos))):
                for ln in infos[idx]["lines"]:
                    r = _sanity_find_ref(ln)
                    if r:
                        deck_line["ser"] = r
                        break
                if deck_line["ser"]:
                    break

    # Source first; the deck read-back only steps in when the field is blank
    # (covers sections pulled in from .pptx).
    scr_ref, scr_src = (src, "week") if src else (deck_line["scr"], "deck")
    ser_ref, ser_src = (s_sr, "week") if s_sr else (deck_line["ser"], "deck")
    scr_ref = scr_ref or ""
    ser_ref = ser_ref or ""

    n_scr, n_ser = _sanity_ref_norm(scr_ref), _sanity_ref_norm(ser_ref)
    if not n_scr and not n_ser:
        status = "missing"
    elif not n_scr or not n_ser:
        status = "missing"
    elif n_scr == n_ser:
        status = "ok"
    elif n_scr in n_ser or n_ser in n_scr:
        status = "contained"
    else:
        status = "mismatch"
    return {"status": status, "scr": scr_ref, "ser": ser_ref,
            "scr_src": scr_src, "ser_src": ser_src}


_PAGE_LABEL_RE = re.compile(r"^(\d+)/(\d+)$")

# Section headers that may sit atop an imported song deck (獻詩／回應) — treated
# as headers, not songs, by the slide-count check.  generate_deck only
# auto-drops the dividers it adds itself, so 獻詩 is listed here.
_SANITY_DIVIDERS = set(_DIVIDER_TITLES) | {"獻詩", "奉獻詩", "奉獻"}


def _sanity_page_label(lines):
    """The `N/M` page label of a song lyric slide, or None (a title card)."""
    for ln in lines:
        m = _PAGE_LABEL_RE.match(ln.strip())
        if m:
            return int(m.group(1)), int(m.group(2))
    return None


def _sanity_card_title(lines):
    """The title printed on a song title card — its first text line that is not
    an `N/M` page label — or None for a blank card."""
    for ln in lines:
        t = ln.strip()
        if t and not _PAGE_LABEL_RE.match(t):
            return t
    return None


def _sanity_is_divider(lines):
    """True when a slide is nothing but a section title card such as 詩歌敬拜 —
    its text (page labels aside) equals one of the divider titles, so a hymn
    title card ("真光普照 / 美樂頌125") never matches."""
    joined = "".join(ln.strip() for ln in lines
                     if not _PAGE_LABEL_RE.match(ln.strip()))
    joined = joined.replace(" ", "").replace("　", "")
    return joined in _SANITY_DIVIDERS


def _sanity_song_groups(infos):
    """[(slides, title)] splitting a song section deck into songs, where
    `slides` is that song's slide infos (its title card first).

    Decks produced here tag every lyric slide with an `N/M` page label, so a
    label-less slide starts a new song — counting even a title-only song as one
    slide.  Imported decks carry no labels; there a slide whose largest font
    stands clearly above the lyric body size is taken as a title card, and
    known section-title cards (詩歌敬拜／獻詩／詩歌回應…) are ignored.  Returns
    None when the songs cannot be told apart (fall back to content estimates)."""
    if not infos:
        return None
    if any(_sanity_page_label(s["lines"]) for s in infos):
        groups, cur = [], None
        for s in infos:
            if cur is None or _sanity_page_label(s["lines"]) is None:
                if cur is not None:            # a new title card begins a song
                    groups.append(cur)
                cur = ([[s], _sanity_card_title(s["lines"])])
            else:
                cur[0].append(s)
        if cur is not None:
            groups.append(cur)
        return [(sl, t) for sl, t in groups]

    # No page labels → imported deck: split on font-size title cards.
    sizes = [s.get("maxsize") for s in infos if s.get("maxsize")]
    if not sizes:
        return None
    body = max(set(sizes), key=sizes.count)      # most common lyric size
    cards = [i for i, s in enumerate(infos)
             if s.get("maxsize") and s["maxsize"] > body + 5
             and not _sanity_is_divider(s["lines"])]
    if not cards:
        return None
    groups = []
    for k, i in enumerate(cards):
        end = cards[k + 1] if k + 1 < len(cards) else len(infos)
        groups.append((infos[i:end], _sanity_card_title(infos[i]["lines"])))
    return groups


def _sanity_group_song_slides(infos):
    """[(slides, title)] → [(count, title)] (each count includes the title card)."""
    groups = _sanity_song_groups(infos)
    return None if groups is None else [(len(sl), t) for sl, t in groups]


def _sanity_norm_text(s):
    """Lyric text reduced to its letters/characters — whitespace and punctuation
    dropped — so full-width commas/periods and line wrapping do not matter."""
    return re.sub(r"\W+", "", s or "", flags=re.UNICODE)


def _sanity_expected_lyrics(cfg):
    """The lyric lines `cfg` would put on its slides, in the generator's order
    (each verse, the refrain repeated after it when so configured, else once)."""
    verses = cfg.get("verses") or []
    refrain = cfg.get("refrain") or None
    repeat = bool(refrain) and cfg.get("refrain_after_every_verse", True)
    out = []
    for vs in verses:
        out.extend(vs)
        if repeat:
            out.extend(refrain)
    if refrain and not repeat:
        out.extend(refrain)
    return out


def _sanity_song_sources(week):
    """[(label, song_cfg, slides)] for every content song in the service — the
    詩歌敬拜 hymns masked by index onto the songs deck, then 詩歌回應.  `slides`
    is that song's slide infos, or None when it has no generated slides."""
    hymns = week.get("hymns") or []
    groups = _sanity_song_groups(_sanity_section_deck_infos("songs", week))
    resp = week.get("response") or {}
    rgroups = _sanity_song_groups(_sanity_section_deck_infos("response", week))
    out = []
    for i, h in enumerate(hymns):
        sl = groups[i][0] if groups and i < len(groups) else None
        out.append((h.get("title") or f"詩歌敬拜 第 {i + 1} 首", h, sl))
    if resp.get("title") or resp.get("verses"):
        sl = rgroups[0][0] if rgroups else None
        out.append((resp.get("title") or "詩歌回應", resp, sl))
    return out


def _sanity_lyric_check(week):
    """Compare each song's entered/edited lyrics with the words actually on its
    generated slides.  Returns (diffs, compared): `diffs` is [(label, missing,
    extra)] — content lyric lines not found on the slides, and slide lines not
    found in the content — for songs that differ (an uploaded deck, or a saved
    songs deck left stale after an edit); `compared` is how many songs had both
    content lyrics and slides to check."""
    diffs, compared = [], 0
    for label, cfg, slides in _sanity_song_sources(week):
        if not slides or len(slides) < 2:
            continue
        expected = _sanity_expected_lyrics(cfg)
        if not any(_sanity_norm_text(ln) for ln in expected):
            continue
        compared += 1
        slide_lines = [ln for s in slides[1:] for ln in s["lines"]
                       if not _PAGE_LABEL_RE.match(ln.strip())]
        exp_norm = [(ln, _sanity_norm_text(ln)) for ln in expected]
        sl_norm = [(ln, _sanity_norm_text(ln)) for ln in slide_lines]
        exp_set = {x for _ln, x in exp_norm if x}
        sl_set = {x for _ln, x in sl_norm if x}
        missing = [ln for ln, x in exp_norm if x and x not in sl_set]
        extra = [ln for ln, x in sl_norm if x and x not in exp_set]
        if missing or extra:
            diffs.append((label, missing, extra))
    return diffs, compared


def _sanity_char_diff(a, b):
    """Mark, character by character, how normalised line `a` becomes `b`:
    changed runs as 〔a→b〕, drops as 〔-a〕 and additions as 〔+b〕."""
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    out = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            out.append(a[i1:i2])
        elif tag == "replace":
            out.append("〔%s→%s〕" % (a[i1:i2], b[j1:j2]))
        elif tag == "delete":
            out.append("〔-%s〕" % a[i1:i2])
        else:
            out.append("〔+%s〕" % b[j1:j2])
    return "".join(out)


def _sanity_lib_norm(s):
    """Library-style fold of a lyric line: 祢/祂 read as 你/他 (a spelling
    variant, not a typo) and punctuation/space dropped (see lyric_search.norm)."""
    return lyric_search.norm((s or "").replace("祢", "你").replace("祂", "他"))


def _sanity_lib_song(cfg, lib):
    """The built-in library hymn matching song `cfg` — (name, collection,
    [lines]) or None.

    Matches the 曲名 against the library names and confirms it with a lyric
    overlap, so a same-named but different hymn is never reported against;
    returns None (skip the check) when nothing overlaps."""
    title = _sanity_lib_norm(cfg.get("title"))
    if len(title) < 2:
        return None
    clines = {x for x in (_sanity_lib_norm(ln)
                          for ln in _sanity_expected_lyrics(cfg))
              if len(x) >= 3}
    if not clines:
        return None
    need = max(2, len(clines) // 3)
    best = None
    for key, (name, where) in lib.meta.items():
        nname = _sanity_lib_norm(name)
        if not nname:
            continue
        if nname != title and title not in nname \
                and not (len(nname) >= 3 and nname in title):
            continue
        ltext = "\n".join(text for _no, text, _n in lib.songs[key])
        lnorms = [_sanity_lib_norm(ln) for ln in ltext.split("\n")]
        shared = sum(1 for c in clines if any(c in l for l in lnorms))
        if shared >= need and (best is None or shared > best[0]):
            best = (shared, key, name, where)
    if best is None:
        return None
    _shared, key, name, where = best
    raw = [ln.strip() for _no, text, _n in lib.songs[key]
           for ln in text.split("\n") if ln.strip()]
    return name, where, raw


def _sanity_lib_word_check(week, limit=6):
    """Fingerprint each content song's lyrics against the built-in hymn library.

    Returns (diffs, compared): `diffs` is [(label, collection, pairs)] where
    `pairs` is [(content_line, library_line, diff)] for lines that closely
    resemble a library line but are not identical — a likely 用錯字 (the library
    edition may differ, so these are hints); `compared` counts the songs that
    matched a library hymn and so could be checked."""
    try:
        lib = lyric_search.get_library()
    except Exception:
        return [], 0
    if not len(lib):
        return [], 0
    resp = week.get("response") or {}
    songs = [(h.get("title") or f"詩歌敬拜 第 {i + 1} 首", h)
             for i, h in enumerate(week.get("hymns") or [])]
    if resp.get("title") or resp.get("verses"):
        songs.append((resp.get("title") or "詩歌回應", resp))
    diffs, compared = [], 0
    for label, cfg in songs:
        hit = _sanity_lib_song(cfg, lib)
        if not hit:
            continue
        compared += 1
        _name, where, raw = hit
        lib_lines = [(ln, _sanity_lib_norm(ln)) for ln in raw]
        lib_norms = [x for _ln, x in lib_lines if len(x) >= 2]
        pairs = []
        for cl in _sanity_expected_lyrics(cfg):
            c = _sanity_lib_norm(cl)
            if len(c) < 3:
                continue
            if c in lib_norms or any(c in x for x in lib_norms):
                continue
            best, best_r = None, 0.0
            for ln, x in lib_lines:
                if len(x) < 2:
                    continue
                r = difflib.SequenceMatcher(None, c, x, autojunk=False).ratio()
                if r > best_r:
                    best, best_r = (ln, x), r
            if best is not None and best_r >= 0.6:
                pairs.append((cl, best[0], _sanity_char_diff(c, best[1])))
        if pairs:
            diffs.append((label, where, pairs[:limit]))
    return diffs, compared


def _sanity_slide_text(s):
    """A song lyric slide's text with its `N/M` page label dropped, normalised
    (punctuation/space removed) so wrapping and full-width marks do not matter."""
    return _sanity_norm_text("".join(
        ln for ln in s["lines"] if not _PAGE_LABEL_RE.match(ln.strip())))


def _sanity_slide_duplicates(week, threshold=0.85, limit=4):
    """Find pairs of a song's own lyric slides that are almost identical — the
    same slide printed twice with a word changed (a likely typo, an extra／missing
    word, or a second edition that survived trimming).  Exactly identical slides
    (a refrain sung the same way each time) are expected and not reported.

    Returns (dups, compared): `dups` is [(label, pairs)] with `pairs` =
    [(page_a, page_b, ratio, diff)]; `compared` counts the songs that had at least
    two lyric slides to compare."""
    dups, compared = [], 0
    for label, _cfg, slides in _sanity_song_sources(week):
        if not slides or len(slides) < 3:      # title card + 2 lyric slides
            continue
        body = slides[1:]
        texts, nos = [], []
        for k, s in enumerate(body):
            texts.append(_sanity_slide_text(s))
            pl = _sanity_page_label(s["lines"])
            nos.append(pl[0] if pl else k + 2)
        if sum(1 for t in texts if t) < 2:
            continue
        compared += 1
        pairs = []
        for i in range(len(texts)):
            if not texts[i]:
                continue
            for j in range(i + 1, len(texts)):
                if not texts[j] or texts[i] == texts[j]:
                    continue
                r = difflib.SequenceMatcher(None, texts[i], texts[j],
                                            autojunk=False).ratio()
                if r >= threshold:
                    pairs.append((nos[i], nos[j], r,
                                  _sanity_char_diff(texts[i], texts[j])))
        if pairs:
            pairs.sort(key=lambda p: (-p[2], p[0], p[1]))
            dups.append((label, pairs[:limit]))
    return dups, compared


_SANITY_UPLOAD_KEYS = {"songs": "build_songs", "offering": "build_offering",
                       "response": "build_response"}


def _sanity_section_deck_infos(section, week):
    """Slides of the generated section deck that build_deck will merge into the
    full deck, resolved the same way: an upload chosen this session, else the
    saved section pptx when 合併先前各步驟投影片檔 is on, else freshly compiled
    from the assembled content.  Returns [{bg, lines, imgs, maxsize}] or None."""
    data = None
    up = sv(_SANITY_UPLOAD_KEYS.get(section, ""), None)
    if up is not None:
        try:
            data = up.getvalue()
        except AttributeError:
            try:
                with open(str(up), "rb") as fh:
                    data = fh.read()
            except OSError:
                data = None
    if data is None and sv("use_saved", True):
        path = saved_section_path(section, DATE)
        if path:
            try:
                with open(path, "rb") as fh:
                    data = fh.read()
            except OSError:
                data = None
    if data is None:
        build_sec = "hymns" if section == "songs" else section
        try:
            data, _ = build_section_pptx_bytes(build_sec, week)
        except Exception:
            return None
    try:
        return _deck_slide_infos(data)
    except Exception:
        return None


def _estimate_song_slides(verses, refrain, repeat_after_every):
    """Total slides a song needs — 1 title card + lyric pages — with the same
    refrain arithmetic the generator uses (see _hymn_slide_count)."""
    n = 1 + len(verses or [])
    if refrain:
        n += len(verses or []) if repeat_after_every else 1
    return n


def _sanity_section_counts(section, week, default_label):
    """[(label, slides)] for a single song-like section (獻詩／詩歌回應) that is
    merged from its own deck: one row per song, split on title cards, or a
    single row for the whole deck when it cannot be split.  [] when the section
    has no generated slides at all (so nothing is reported)."""
    infos = _sanity_section_deck_infos(section, week)
    if not infos:
        return []
    groups = _sanity_group_song_slides(infos)
    if groups:
        rows = []
        for k, (n, card) in enumerate(groups):
            label = card or (default_label if k == 0
                             else f"{default_label}（第 {k + 1} 首）")
            rows.append((label, n))
        return rows
    n = sum(1 for s in infos if not _sanity_is_divider(s["lines"]))
    return [(default_label, n)] if n else []


def _sanity_song_counts(week):
    """[(label, total slides incl. title card)] for every song that has slides —
    詩歌敬拜 hymns first, then 獻詩 (offering, only when a deck is merged), then
    詩歌回應.  Counts come from the GENERATED slides: each section deck (saved
    pptx when one is merged, else compiled from content).  Song names come from
    the content, falling back to the title printed on each generated title
    card.  Content estimates are the last resort when a deck cannot be grouped."""
    hymns = week.get("hymns") or []
    resp = week.get("response") or {}
    resp_has = bool(resp.get("title") or resp.get("verses"))
    out = []

    groups = _sanity_group_song_slides(
        _sanity_section_deck_infos("songs", week))
    for i, h in enumerate(hymns):
        card = None
        if groups and i < len(groups):
            n, card = groups[i]
        else:
            n = _estimate_song_slides(h.get("verses") or [], h.get("refrain"),
                                      h.get("refrain_after_every_verse", True))
        label = h.get("title") or card or f"詩歌敬拜 第 {i + 1} 首"
        out.append((label, n))
    if groups:
        for i in range(len(hymns), len(groups)):
            n, card = groups[i]
            out.append((card or f"未命名詩歌（第 {i + 1} 首）", n))

    # 獻詩（只有上載／雲端的 pptx，沒有內容欄位；有投影片才算）。
    out += _sanity_section_counts("offering", week, "獻詩")

    # 詩歌回應（有投影片或有內容才算）。
    rrows = _sanity_section_counts("response", week, "詩歌回應")
    if rrows:
        if resp.get("title"):
            _label, n = rrows[0]
            rrows[0] = (resp["title"], n)
        out += rrows
    elif resp_has:
        n = _estimate_song_slides(resp.get("verses") or [],
                                  resp.get("refrain"),
                                  resp.get("refrain_after_every_verse", False))
        out.append((resp.get("title") or "詩歌回應", n))
    return out


def _sanity_expected_image(name):
    """Path of the built-in template image `name` (communion / apostles_creed),
    mirroring generate_deck.image_path's default resolution."""
    val = IMAGE_SLIDES.get(name)
    if not val:
        return None
    if not os.path.isabs(val):
        val = os.path.join(os.getcwd(), val)
    return val if os.path.isfile(val) else None


def _sanity_image_sha(path):
    try:
        with open(path, "rb") as fh:
            return hashlib.sha1(fh.read()).hexdigest()
    except OSError:
        return None


def _sanity_communion_check(week, infos=None):
    """First-Sunday template pages: 聖餐 + 使徒信經 should be in the deck.
    Returns (ok, note) — ok False when a template page is missing.  An
    imported communion_<date>.pptx substitutes for both, so the note says so."""
    if not is_first_sunday(DATE):
        return True, ""
    if infos:
        wanted = {n: _sanity_image_sha(_sanity_expected_image(n))
                  for n in ("communion", "apostles_creed")}
        if not all(wanted.values()):
            return True, "（第一主日：樣板圖片檔無法讀取，無法比對投影片。）"
        found = set()
        for s in infos:
            for h in s.get("imgs", ()):
                for n, want in wanted.items():
                    if want and h == want:
                        found.add(n)
        labels = {"communion": "聖餐", "apostles_creed": "使徒信經"}
        missing = [n for n in ("communion", "apostles_creed") if n not in found]
        if missing:
            return False, "第一主日：投影片中未見" + "、".join(
                labels[n] for n in missing) + "樣板頁" + \
                "（若以雲端 _communion_<日期>.pptx_ 取代則可略過）。"
        return True, ""
    if not week.get("communion"):
        return False, "第一主日：內容未勾選「聖餐＋使徒信經」" + \
            "（若已由雲端 _communion_<日期>.pptx_ 取代則不必勾選）。"
    return True, ""


def _scripture_editor(prefix, section=None):
    """讀經經文 editor — step 4 of the main wizard, or the standalone
    經文投影片 flow. Field keys are `{prefix}_{field}`. With `section`
    ("scripture"), the fhl.net fetch also auto-saves to the store; the
    standalone flow passes section=None so nothing touches Google Sheets."""
    st.selectbox("書卷", bible.BOOK_NAMES, index=None,
                 placeholder="選擇書卷（也可直接在下面輸入出處）",
                 format_func=lambda n: f"{n}（{bible.book_code(n)} / "
                 f"{bible.book_en(n)}）",
                 key=K(f"{prefix}_book"),
                 on_change=_pick_book,
                 args=(K(f"{prefix}_book"), K(f"{prefix}_ref")))
    st.caption("選書卷會自動填入出處；也可直接在下面「出處」欄鍵入，例如 "
               "詩篇 34:1-3")
    ref_text_key = K(f"{prefix}_ref")
    st.session_state.setdefault(ref_text_key, sv(f"{prefix}_ref"))
    st.text_input("出處", key=ref_text_key,
                  placeholder="羅馬書 12:1-8")
    st.radio("譯本版本", VERSION_OPTIONS, horizontal=True,
             index=_ref_version_index(sv(f"{prefix}_ref")),
             key=K(f"{prefix}_version"), on_change=_pick_version,
             args=(K(f"{prefix}_version"), K(f"{prefix}_ref"),
                   K(f"{prefix}_verses"), section))
    st.button("從 fhl.net 聖經網輸入", type="primary",
              key=K(f"{prefix}_load"), use_container_width=True,
              on_click=_fetch_into,
              args=(K(f"{prefix}_ref"), K(f"{prefix}_verses"), section))
    st.caption("支援多卷書：以 `；` 分隔（如 詩篇 34:1-3；馬太福音 6:9-13）")
    c1, c2 = st.columns(2)
    ref_key, font_key = K(f"{prefix}_ref_size"), K(f"{prefix}_font_size")
    # Weeks saved before these floors were raised can carry a smaller font_size
    # (e.g. 40).  Streamlit validates a pre-set widget value against
    # min_value/max_value and raises StreamlitValueBelowMinError, so pull the
    # stored value back into range before the widget renders.
    st.session_state.setdefault(ref_key, 48)
    st.session_state.setdefault(font_key, 44)
    _clamp_widget(ref_key, REF_SIZE_MIN, REF_SIZE_MAX)
    _clamp_widget(font_key, BODY_SIZE_MIN, BODY_SIZE_MAX)
    c1.number_input("出處字型(pt)", min_value=REF_SIZE_MIN,
                    max_value=REF_SIZE_MAX, key=ref_key)
    c2.number_input("內文字型(pt)", min_value=BODY_SIZE_MIN,
                    max_value=BODY_SIZE_MAX, key=font_key)
    _layout_radio(K(f"{prefix}_layout"))
    st.caption("「整段」= 節碼以上標內嵌、整段連續排版；"
               "「列表」= 一節一行，分頁只在節與節之間。")
    verses_key = K(f"{prefix}_verses")
    st.session_state.setdefault(verses_key, sv(f"{prefix}_verses"))
    st.text_area("內容（貼上全部經文，自動分頁）",
                 key=verses_key, height=200)
    _fetch_status(f"{prefix}_verses")


def _render_scr_flow():
    """📜 經文投影片（獨立流程）— a 1-step page separate from the main wizard.
    Content lives under its own `scrflow_*` session keys (never assembled into
    the week or saved to the sheet); the generate button is inside the step."""
    if st.button("← 返回主流程", key=K("scr_flow_back")):
        st.session_state["scr_flow"] = False
        st.rerun()
    st.markdown('<div class="hero hero-sm"><div class="hero-date">📜 經文投影片'
                '<span class="brand-tail">（獨立流程）</span></div>'
                '<div class="hero-sub">單一步驟：只依輸入的經文製成 pptx · '
                '內容不會儲存至 Google Sheet</div></div>', unsafe_allow_html=True)

    st.progress(1.0, text="第 1 / 1 步驟 · 經文投影片")
    st.divider()

    _font_radio("經文投影片使用的字型", K("scrflow_font"))

    with st.container(border=True):
        section_header("1", "經文投影片", "輸入出處一鍵載入，或直接貼上整段經文")
        _scripture_editor("scrflow")

    st.divider()
    build_clicked = st.button("📜 製成經文投影片 pptx", type="primary",
                              use_container_width=True,
                              key=K("scrflow_build"))
    if build_clicked:
        try:
            week = normalize_week({
                "date": DATE,
                "call_font": sv("scrflow_font", "DFKai-SB"),
                "scripture": {
                    "ref": sv("scrflow_ref", ""),
                    "ref_size": int(sv("scrflow_ref_size", 48) or 48),
                    "font_size": int(sv("scrflow_font_size", 44) or 44),
                    "layout": str(sv("scrflow_layout", READING_LAYOUT_DEFAULT)
                                  or READING_LAYOUT_DEFAULT),
                    "verses": split_lines(sv("scrflow_verses", "")),
                },
            })
            data, count = build_section_pptx_bytes("scripture", week)
            st.session_state["scrflow_bytes"] = data
            st.session_state["scrflow_count"] = count
            st.success(f"完成：{count} 頁（獨立流程，只含讀經經文）")
            _log_action("製成經文投影片", f"{DATE} · {count} 頁")
        except Exception as exc:
            st.error(f"製成失敗：{exc}")

    data = st.session_state.get("scrflow_bytes")
    if data:
        st.download_button(
            "⬇️ 下載經文投影片 pptx",
            data=data,
            file_name=f"scripture_{DATE}.pptx",
            mime="application/vnd.openxmlformats-officedocument"
                 ".presentationml.presentation",
        )
        st.caption("此獨立檔案不會被「製成整場投影片」讀取或合併。")


_inject_css()

# --------------------------------------------------------------- sidebar --
with st.sidebar:
    st.markdown('<div class="brand">崇拜投影片'
                '<span class="brand-tail"> 編輯器</span></div>',
                unsafe_allow_html=True)

    # 新增週次 handled BEFORE the date dropdown renders: the fresh week must
    # already exist in the store, otherwise the dropdown (built from existing
    # dates only) immediately reverts to an older date and the new week is
    # never opened.
    with st.expander("＋ 新增週次"):
        new_date = st.text_input("日期（YYYY.MM.DD）",
                                 placeholder="2026.10.04", key=K("new_week_date"))
        if st.button("新增並開啟", use_container_width=True,
                     key=K("new_week_add")):
            if DATE_RE.match(new_date):
                if new_date not in store.all().keys():
                    store.save(default_week(new_date))
                    st.cache_data.clear()
                set_active(new_date)
            else:
                st.error("日期格式應為 2026.09.20")

    if st.button("🔄 重新載入 Google Sheet", use_container_width=True,
                 key=K("reload_sheet"), help="清除快取，立即重讀 Google Sheet（手動在 Sheet 改過資料時用）"):
        invalidate()
        _clear_drafts(DATE)
        st.cache_data.clear()
        st.rerun()

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
            _clear_drafts(DATE)
            st.cache_data.clear()
            st.toast(f"已儲存 {week_to_save['date']}")
            _log_action("儲存內容", week_to_save["date"])

    gen_clicked = st.button("📽 製成整場投影片（合併＋編譯）", type="primary",
                            use_container_width=True, key=K("generate"))
    gen_slot = st.container()

    st.divider()
    st.markdown('<div class="sec-title" style="font-size:1rem">📜 經文投影片'
                '<span class="brand-tail">（獨立流程）</span></div>',
                unsafe_allow_html=True)
    st.caption("單一步驟：只輸入「經文」內容製成 pptx；\n"
               "不會儲存到 Google Sheet，也不屬於整場流程。")
    if st.button("📜 開啟「經文投影片」獨立流程",
                 use_container_width=True, key=K("enter_scr_flow")):
        st.session_state["scr_flow"] = True
        st.rerun()

# ------------------------------------------------------------------ tabs --
if st.session_state.get("scr_flow"):
    _render_scr_flow()
    st.stop()

tab_edit, tab_build = st.tabs(["✏️ 編輯內容", "📽 製作成投影片"])

with tab_edit:
    st.markdown(
        f'<div class="hero"><div class="hero-date">主日：{DATE}</div>'
        f'<div class="hero-sub">崇拜投影片編輯器 · 內容自動儲存至 Google Sheet'
        f'</div></div>', unsafe_allow_html=True)

    step = _edit_step()

    _FONT_OPTIONS = [
        ("DFKai-SB", "DFKai-SB（標楷體）"),
        ("PMingLiU", "PMingLiU（新細明體）"),
        ("JhengHei", "JhengHei（正黑體）"),
        ("DengXian", "DengXian（等線體）"),
        ("LiSu", "LiSu（隸書）"),
    ]
    _font_label = dict(_FONT_OPTIONS)
    st.radio(
        "全場統一使用的字型（套用所有投影片文字）",
        options=[f for f, _ in _FONT_OPTIONS],
        format_func=lambda f: _font_label.get(f, f),
        key=K("call_font"),
        horizontal=True)

    _step_marker(step)
    st.divider()

    if step == 1:
        with st.container(border=True):
            section_header("1", "宣召經文", "貼上整段經文，自動分頁")
            _odrive_panel("psalm", "宣召經文")
            if not _cloud_deck_status("psalm", "宣召經文"):
                st.caption("尚未讀取 `psalm_<日期>.pptx`；下方仍可直接貼上經文"
                           "由網頁內容編譯。")
            _scripture_editor("psalm", "psalm")

    elif step == 2:
        with st.container(border=True):
            section_header("2", "詩歌敬拜（Hymns）")
            _odrive_panel("songs", "詩歌")
            section_deck_panel("songs", "hymns", "songs", "詩歌",
                               f"songs_slides_{DATE}.pptx",
                               btn="🎵 產生詩歌投影片")
            if not _cloud_deck_status("songs", "詩歌", "build_songs"):
                st.caption("以下內容以網頁內容現場編譯；上載 .pptx 時"
                           "以檔案為準，字型／項目符號沿用來源檔。")
            st.checkbox("歌詞文字加上陰影（有助背景圖上閱讀，寫入投影片）",
                        value=sv("hymn_text_shadow", True),
                        key=K("hymn_text_shadow"))
            _song_font_selector(
                "歌詞字級（pt）",
                hint="固定字級（預設 52pt、可選 48–52pt），不隨行長縮小；"
                     "行太長時自動換行。"
                     "標點符號以空格取代（如「，」「。」→ 空白）。"
                     "每首詩歌可用下方「背景圖」各自指定背景。")

            hymns_ui = _get_ui("hymns")
            hymn_uids = [h["_id"] for h in hymns_ui]
            # Seed each song expander's open/closed state once, so the summary
            # strip's 前往編輯 can drive it; the first song starts open.
            for i, h in enumerate(hymns_ui):
                st.session_state.setdefault(
                    K(f"hymn_exp_{h['_id']}"), i == 0)
            if hymns_ui:
                st.markdown("**本週詩歌一覽**")
                summary = st.columns(len(hymns_ui))
                for i, h in enumerate(hymns_ui):
                    uid = h["_id"]
                    badge = _song_badge(uid, h)
                    open_now = bool(st.session_state.get(K(f"hymn_exp_{uid}")))
                    with summary[i], st.container(border=True):
                        st.markdown(
                            "<div style='font-size:.72rem;color:var(--muted);"
                            "letter-spacing:.05em'>詩歌 %d</div>"
                            "<div style='font-weight:700;margin:.1rem 0'>%s</div>"
                            % (i + 1,
                               html.escape(_song_title(uid, h) or "（未命名）")),
                            unsafe_allow_html=True)
                        if badge.startswith("⚠"):
                            st.markdown(
                                "<div style='color:#B45309;font-size:.82rem'>"
                                f"{badge}</div>", unsafe_allow_html=True)
                        else:
                            st.caption(badge)
                        if open_now:
                            st.markdown(
                                "<div style='color:var(--accent);font-size:.78rem'>"
                                "▾ 目前展開</div>", unsafe_allow_html=True)
                        st.button("前往編輯", key=K(f"hymn_goto_{uid}"),
                                  use_container_width=True,
                                  on_click=_focus_hymn,
                                  args=(uid, hymn_uids))

            for i, h in enumerate(hymns_ui):
                uid = h["_id"]
                badge = _song_badge(uid, h)
                # The label is part of a keyed expander's identity, so it must
                # stay constant: a live title/段數 here changes the widget id on
                # every keystroke and Streamlit remounts the panel shut.  The
                # live title/status go on the line inside and in the strip above.
                with st.expander(f"詩歌 {i + 1}", key=K(f"hymn_exp_{uid}"),
                                 on_change="rerun"):
                    st.caption(
                        f"{_song_title(uid, h) or '（未命名）'} · {badge}")
                    _lyric_search_panel("hymn", uid)
                    _lyric_editor("hymn", uid)
                    c1, c2 = st.columns(2)
                    c1.text_input("曲名",
                                  value=sv(f"hymn_{uid}_title", h.get("title")),
                                  key=K(f"hymn_{uid}_title"))
                    c2.text_input("英文名",
                                  value=sv(f"hymn_{uid}_subtitle",
                                           h.get("subtitle")),
                                  key=K(f"hymn_{uid}_subtitle"))
                    m1, m2, m3 = st.columns(3)
                    m1.text_input("詩集",
                                  value=sv(f"hymn_{uid}_source", h.get("source")),
                                  key=K(f"hymn_{uid}_source"))
                    m2.text_input("曲",
                                  value=sv(f"hymn_{uid}_music", h.get("music")),
                                  key=K(f"hymn_{uid}_music"))
                    m3.text_input("詞",
                                  value=sv(f"hymn_{uid}_lyricist",
                                           h.get("lyricist")),
                                  key=K(f"hymn_{uid}_lyricist"))
                    with st.expander("🎨 背景圖"):
                        _bg_picker(f"hymn_{uid}_bg", h.get("bg", ""))
                    st.divider()
                    with st.expander("▶ 預覽這首詩歌的 slides"):
                        pcol1, pcol2 = st.columns([0.42, 0.58])
                        if pcol1.button("🎬 產生這首詩歌的 slide",
                                        key=K(f"hymn_{uid}_preview_gen"),
                                        use_container_width=True):
                            try:
                                data, cnt = build_hymn_pptx_bytes(
                                    assemble_week(DATE), assemble_hymn(DATE, h))
                                st.session_state[K(f"hymn_{uid}_preview_data")] = data
                                st.session_state[K(f"hymn_{uid}_preview_count")] = cnt
                                st.session_state[K(f"hymn_{uid}_preview_html")] = \
                                    _hymn_preview_html(hymn_preview(data))
                            except Exception as exc:
                                st.error(f"產生失敗：{exc}")
                        if pcol2.button("🗑 清除預覽",
                                        key=K(f"hymn_{uid}_preview_clear"),
                                        use_container_width=True,
                                        disabled=not st.session_state.get(
                                            K(f"hymn_{uid}_preview_data"))):
                            for kk in ("_preview_data", "_preview_count",
                                       "_preview_html"):
                                st.session_state.pop(K(f"hymn_{uid}{kk}"), None)
                            st.rerun()
                        _render_hymn_preview(uid, h.get("title", ""))
                    st.divider()
                    if st.columns([1, 2])[0].button(
                            "🗑 移除這首詩歌", key=K(f"hymn_{uid}_del")):
                        st.session_state[K("del_hymn_pending")] = i
                        st.rerun()
            if st.session_state.get(K("del_hymn_pending")) is not None:
                _confirm_del_hymn()
            if st.button("＋ 新增詩歌", key=K("add_hymn")):
                _add_item("hymns", {"_id": uuid.uuid4().hex[:8], "title": "",
                                    "subtitle": "", "source": "", "music": "",
                                    "lyricist": "", "bg": "",
                                    "refrain": "", "repeat": True,
                                    "stanzas": [""]})
                st.rerun()
            _finalize_stamp("hymns_finalized", "詩歌敬拜")

    elif step == 3:
        with st.container(border=True):
            section_header("3", "獻詩（Offering）")
            _odrive_panel("offering", "獻詩")
            section_deck_panel("offering", None, "offering", "獻詩", None)
            if not _cloud_deck_status("offering", "獻詩", "build_offering"):
                st.info("此節只接受現成 .pptx：上載後製成整場時以檔案為準；"
                        "未上載則此節不加入。上載檔在 Server 重啟後可能遺失，"
                        "請當日完成整場製作。")

    elif step == 4:
        with st.container(border=True):
            section_header("4", "讀經經文", "貼上整段經文，自動分頁")
            _odrive_panel("scripture", "讀經經文")
            if not _cloud_deck_status("scripture", "讀經經文"):
                st.caption("尚未讀取 `scripture_<日期>.pptx`；下方仍可直接貼上經文"
                           "由網頁內容編譯。")
            _scripture_editor("scripture", "scripture")

    elif step == 5:
        with st.container(border=True):
            section_header("5", "講道信息（Sermon）")
            _odrive_panel("sermon", "講道")
            section_deck_panel("sermon", "sermon", "sermon", "講道",
                               f"sermon_{DATE}.pptx",
                               btn="🗣 產生講道投影片")
            if not _cloud_deck_status("sermon", "講道", "build_sermon"):
                st.info("此節以現成 .pptx 為準；上載後製成整場時以該檔案為準。"
                        "上載檔在 Server 重啟後可能遺失，請當日完成整場製作。")

    elif step == 6:
        with st.container(border=True):
            section_header("6", "詩歌回應（Response Hymn）",
                           "單首回應詩歌，緊接信息之後（可選）")
            _odrive_panel("response", "詩歌回應")
            section_deck_panel("response", "response", "response", "詩歌回應",
                               f"response_{DATE}.pptx",
                               btn="🎶 產生詩歌回應投影片")
            if not _cloud_deck_status("response", "詩歌回應", "build_response"):
                st.caption("此節**可選**：不需要可直接按「下一步」略過。"
                           "要加入時，可填寫以下內容現場編譯；"
                           "或在上方直接上載現成 .pptx 合併（以檔案為準）。")
            _song_font_selector(
                "歌詞字級（pt）",
                hint="與「詩歌敬拜」共用同一設定（預設 52pt、可選 48–54pt）。")
            _lyric_search_panel("resp", "")
            for name in ("resp_title", "resp_subtitle", "resp_source",
                         "resp_music", "resp_lyricist", "resp_bg"):
                st.session_state.setdefault(K(name), "")
            _lyric_editor("resp", "")
            c1, c2 = st.columns(2)
            c1.text_input("曲名", value=sv("resp_title"), key=K("resp_title"))
            c2.text_input("英文名", value=sv("resp_subtitle"),
                          key=K("resp_subtitle"))
            m1, m2, m3 = st.columns(3)
            m1.text_input("詩集", value=sv("resp_source"), key=K("resp_source"))
            m2.text_input("曲", value=sv("resp_music"), key=K("resp_music"))
            m3.text_input("詞", value=sv("resp_lyricist"),
                          key=K("resp_lyricist"))
            with st.expander("🎨 背景圖"):
                _bg_picker("resp_bg", "")
            _finalize_stamp("resp_finalized", "詩歌回應")

    elif step == 7:
        with st.container(border=True):
            section_header("7", "聖餐（Communion）")
            _odrive_panel("communion", "聖餐")
            if not _cloud_deck_status("communion", "聖餐"):
                st.caption("若從雲端取回 `communion_<日期>.pptx`，會**取代**內建的"
                           "聖餐＋使徒信經兩頁；此時不必另外勾選聖餐。")
            st.checkbox("聖餐（第一主日：加聖餐＋使徒信經頁）",
                        value=sv("communion", is_first_sunday(DATE)),
                        key=K("communion"))
            if is_first_sunday(DATE) and not sv("communion", False):
                # The date is a first Sunday but this week was saved unchecked;
                # say so plainly rather than claiming it is pre-checked.
                st.caption("此日期為當月第一主日，**建議勾選**聖餐＋使徒信經；"
                           "本週目前為未勾選狀態（沿用已儲存設定）。")
            elif is_first_sunday(DATE):
                st.caption("此日期為當月第一主日，已預設勾選聖餐＋使徒信經。")

    elif step == 8:
        with st.container(border=True):
            section_header("8", "家事分享（Announcements）")
            _odrive_panel("announcements", "家事分享")
            section_deck_panel("announcements", "announcements",
                               "announcements", "家事分享",
                               f"announcements_{DATE}.pptx",
                               btn="📋 產生家事分享投影片",
                               widget="ann")
            if not _cloud_deck_status("announcements", "家事分享", "build_ann"):
                st.info("此節以現成 .pptx，或以已輸入的家事分享內容編譯；"
                        "上載後以檔案為準。")

    elif step == 9:
        with st.container(border=True):
            _odrive_panel("video", "家事MP4", is_video=True)
            render_video_panel("s9", "9")
        st.caption("產生 MP4 後，按左側「📽 製成整場投影片」即會自動併入"
                   "整場投影片的第一頁（勾選「併入第一頁」時，自動循環播放）。")

    st.divider()
    _step_nav(step)
    if step >= len(EDIT_STEPS):
        st.success("全部步驟已完成 ✔　在左側面板按「📽 製成整場投影片」，"
                   "或到「📽 製作成投影片」頁查看「製成選項與來源狀態」。")

with tab_build:
    st.markdown(
        f'<div class="hero hero-sm"><div class="hero-date">製作成投影片</div>'
        f'<div class="hero-sub">{DATE} · 上載與產生已在「✏️ 編輯內容」各節，'
        f'此頁檢查來源狀態；或以網頁內容現場編譯'
        f'</div></div>', unsafe_allow_html=True)

    # Persist uploaded section decks (uploaded in the 編輯內容 wizard steps) so
    # the section status survives a page refresh (Streamlit clears file_uploader
    # state). Read straight from the uploader session keys.
    persist_uploads(DATE, {
        "songs": st.session_state.get(K("build_songs")),
        "sermon": st.session_state.get(K("build_sermon")),
        "announcements": st.session_state.get(K("build_ann")),
        "offering": st.session_state.get(K("build_offering")),
        "response": st.session_state.get(K("build_response")),
    })

    ui_ann = _get_ui("ann")
    ann_images = {}
    if ui_ann:
        with st.expander("家事分享圖片（圖片型家事分享，可選）"):
            for i, a in enumerate(ui_ann):
                if a.get("mode") == "image":
                    up = st.file_uploader(
                        f"家事分享 {i + 1} 圖片：{a.get('image_name') or '(未上載)'}",
                        type=["jpg", "jpeg", "png", "webp"],
                        key=K(f"build_annimg_{i}"), accept_multiple_files=False)
                    if up is not None:
                        ann_images[i] = up

    with st.container(border=True):
        section_header("A", "步驟狀態")
        use_saved = st.checkbox(
            "合併先前各步驟投影片檔", value=True,
            key=K("use_saved"),
            help="勾選：製成整場時，直接合併先前個別產生的各節 pptx 檔"
                 "（含已上載的檔案），不需要重新編譯。"
                 "取消：全部改用網頁上編輯的內容現場編譯。")
        def _row(label, sec, upload_key=None, fallback=None):
            """One 步驟狀態 line.  A section with a deck on disk is named by that
            deck — by the filename it was fetched/uploaded under, not the local
            re-write — with a ✖ 移除 button.  `fallback` is used for sections
            with no deck, so text-only ones still report their content."""
            text, present = _deck_present(sec, upload_key)
            if not present:
                if fallback is None:
                    st.markdown(f"🧩 **{label}** — 由網頁內容編譯")
                else:
                    st.markdown(fallback)
                return
            note = ("" if use_saved
                    else "　⚠️「合併先前各步驟投影片檔」已取消，實際不會用到。")
            r1, r2 = st.columns([0.78, 0.22], vertical_alignment="center")
            r1.markdown(f"✅ **{label}** — {text}{note}")
            r2.button("✖ 移除", key=K(f"clear_{sec}"),
                      use_container_width=True,
                      on_click=_signal_clear, args=(DATE, sec),
                      help=f"移除{label}的上載／已讀取檔案並刪除已儲存的 pptx，"
                           "恢復用網頁內容編譯")

        def _verses_fallback(_pfx, _lbl):
            """Status for a scripture section with no deck: report the verse text
            that will be compiled instead."""
            _ref = sv(f"{_pfx}_ref", "").strip()
            _verses = split_lines(sv(f"{_pfx}_verses", ""))
            if _ref and _verses:
                return f"✅ **{_lbl}** — {_ref}（{len(_verses)} 行）"
            if _ref:
                return f"⚠️ **{_lbl}** — 出處「{_ref}」已填，內容空白"
            return f"🧩 **{_lbl}** — 未填寫"

        # 順序與「✏️ 編輯內容」一致
        _row("📖 宣召經文", "psalm",
             fallback=_verses_fallback("psalm", "📖 宣召經文"))
        _row("🎵 詩歌敬拜", "songs", "build_songs")
        _row("🙌 獻詩", "offering", "build_offering")
        _row("📖 讀經經文", "scripture",
             fallback=_verses_fallback("scripture", "📖 讀經經文"))
        _row("🗣 講道信息", "sermon", "build_sermon")
        _row("🎶 詩歌回應", "response", "build_response")
        _communion_fallback = (
            "✅ **🍞 聖餐＋使徒信經** — 已加入"
            if sv("communion", is_first_sunday(DATE))
            else "⬜ **🍞 聖餐＋使徒信經** — 未加入"
                 + ("（第一主日建議勾選）" if is_first_sunday(DATE) else ""))
        _row("🍞 聖餐＋使徒信經", "communion", fallback=_communion_fallback)
        _row("📋 家事分享", "announcements", "build_ann")
        vid_ok = (st.session_state.get(K("video_data")) is not None
                  or saved_video_mp4(DATE) is not None)
        if vid_ok:
            first_deck = st.session_state.get(K("video_first_deck_s9"))
            if first_deck is None:
                first_deck = saved_video_meta(DATE)
            vid_text = ("✅ **🎞 家事MP4** — 已產生"
                        + ("，將併入第一頁" if bool(first_deck)
                           else "，未併入第一頁"))
        else:
            vid_text = ("🧩 **🎞 家事MP4** — 未產製"
                        "（請到「編輯內容 → 家事MP4」填入內容並產製）")
        vr1, vr2 = st.columns([0.78, 0.22], vertical_alignment="center")
        vr1.markdown(vid_text)
        if vid_ok:
            vr2.button("✖ 移除", key=K("clear_video_btn"),
                       use_container_width=True,
                       on_click=_signal_clear, args=(DATE, "video"),
                       help="移除已產製的家事MP4並刪除已儲存的 mp4 檔，"
                            "恢復以網頁內容重新產製")

    st.markdown(
        '<div style="background-color:#d8f3dc;color:#14532d;'
        'border:1px solid #52b788;border-radius:8px;padding:12px 16px;'
        'font-weight:600;">完成檢查後，請按左側「📽 製成整場投影片'
        '（合併＋編譯）」產生整場投影片。</div>',
        unsafe_allow_html=True)

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
            st.session_state["video_first_applied"] = False
            week = assemble_week(DATE)
            data, slide_count = build_deck(
                week,
                songs_pptx=pick(st.session_state.get(K("build_songs")),
                                None, "詩歌檔名"),
                sermon_pptx=pick(st.session_state.get(K("build_sermon")),
                                 None, "講道檔名"),
                announcements_pptx=pick(st.session_state.get(K("build_ann")),
                                        None, "家事分享檔名"),
                offering_pptx=pick(st.session_state.get(K("build_offering")),
                                   None, "獻詩檔名"),
                response_pptx=pick(st.session_state.get(K("build_response")),
                                   None, "詩歌回應檔名"),
                announcement_images=ann_images,
                use_saved=use_saved,
            )
            want_video = st.session_state.get(K("video_first_deck_s9"))
            if want_video is None:
                want_video = saved_video_meta(DATE)
            if want_video:
                mp4b = st.session_state.get(K("video_data")) \
                    or _saved_mp4_bytes(DATE)
                if mp4b is None:
                    snap = st.session_state.get(K("video_snap_s9"))
                    if snap and snap.get("kind") != "none":
                        kn = snap.get("ns") or ""
                        mp4b, err = build_video_bytes(
                            snap["kind"], snap["items"], snap["zip_map"],
                            snap["pptx"],
                            st.session_state.get(K(f"video_nums{kn}"),
                                                 snap.get("nums_auto") or "1"),
                            st.session_state.get(K(f"video_secs{kn}"), 10.0),
                            st.session_state.get(K(f"video_fps{kn}"), 15),
                            ann_images)
                        if err:
                            st.warning(f"第一頁的影片未加入：{err}")
                if mp4b:
                    data = prepend_video_slide(data, mp4b,
                                               video_poster_png(mp4b))
                    slide_count += 1
                    st.session_state["video_first_applied"] = True
                else:
                    st.warning("未產製 MP4，無法併入第一頁"
                               "（請先在「✏️ 編輯內容 → 家事MP4」產生影片）。")
            st.session_state["built_bytes"] = data
            st.session_state["built_name"] = f"Sunday_Service_{DATE}.pptx"
            st.session_state["built_count"] = slide_count
            st.session_state["built_date"] = DATE
            if st.session_state.get("video_first_applied"):
                st.success(f"完成：{slide_count} 頁，影片已併入第一頁。")
            else:
                st.success(f"完成：{slide_count} 頁")
            _log_action("製成整場投影片", f"{DATE} · {slide_count} 頁")
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
                   "檔案在瀏覽器下載，不會儲存在Server。")

    st.divider()
    if st.button("🔍 檢查整場投影片",
                 key=K("sanity_check"), use_container_width=True,
                 help="比對「讀經經文」與「講道信息」的經文出處、檢查所有詩歌"
                      "（詩歌敬拜／獻詩／詩歌回應）至少 4 頁（含標題卡）、"
                      "比對投影片歌詞與內容是否一致、對照內建歌詞庫提示可能"
                      "用錯字的地方、找出同一首歌內幾乎相同（只差一個字）的"
                      "重複投影片，並確認第一主日含有聖餐＋使徒信經樣板頁。"
                      "會同時核對內容欄位與已產生的投影片。"):
        week = assemble_week(DATE)
        has_deck = bool(built and st.session_state.get("built_date") == DATE)
        infos = _deck_slide_infos(built) if has_deck else None
        rep = _sanity_check_scripture(week, built if has_deck else None)
        if rep["status"] == "ok":
            st.success(f"✅ 相符：讀經「{rep['scr']}」＝ 講道「{rep['ser']}」")
        elif rep["status"] == "contained":
            st.warning(
                f"⚠️ 範圍重疊：讀經「{rep['scr']}」與講道「{rep['ser']}」"
                "出自同一經卷但範圍不同，請確認讀經範圍是否涵蓋講道經文。")
        elif rep["status"] == "mismatch":
            st.error(
                f"❌ 不一致：讀經「{rep['scr']}」vs 講道「{rep['ser']}」——"
                "經文出處不同，請確認。")
        else:
            missing = ("讀經" if not _sanity_ref_norm(rep["scr"]) else "") \
                + ("講道" if not _sanity_ref_norm(rep["ser"]) else "")
            if not _sanity_ref_norm(rep["scr"]) and not _sanity_ref_norm(rep["ser"]):
                st.info("ℹ️ 讀經與講道都尚未填寫經文出處，無法比對。")
            else:
                st.info(f"ℹ️ {missing}尚未填寫經文出處，無法完整比對。"
                        "填寫後可重新檢查。")
        st.caption(
            "讀經：{}　·　講道：{}　·　{}".format(
                f"「{rep['scr']}」" if rep["scr"]
                else "（未填寫）",
                f"「{rep['ser']}」" if rep["ser"]
                else "（未填寫）",
                "檢查" + ("已產生的整場投影片及內容欄位" if has_deck
                          else "內容欄位（尚未產生整場投影片）")))
        if has_deck:
            st.caption("出處來源：" + ("讀經＝內容欄位" if rep["scr_src"] == "week"
                                       else "讀經＝投影片")
                       + "、"
                       + ("講道＝內容欄位" if rep["ser_src"] == "week"
                          else "講道＝投影片（信息前四張）"))
        # 詩歌頁數：所有詩歌（詩歌敬拜／獻詩／詩歌回應），取自產生的投影片。
        song_counts = _sanity_song_counts(week)
        short = [(lab, n) for lab, n in song_counts if n < 4]
        if short:
            st.warning("⚠️ 以下詩歌少於 4 頁（含標題卡）："
                       + "、".join(f"「{lab}」只有 {n} 頁"
                                   for lab, n in short)
                       + "，請確認是否需補充歌詞段落。")
        if song_counts:
            st.caption("詩歌頁數（含獻詩／回應，取自產生的投影片）：" + "、".join(
                f"「{lab}」{n} 頁" + (" ⚠️" if n < 4 else "")
                for lab, n in song_counts))
        # 歌詞比對：投影片上的歌詞是否與內容一致（修正後是否已重新產生）。
        lyric_diffs, n_cmp = _sanity_lyric_check(week)
        for label, missing, extra in lyric_diffs:
            bits = []
            if missing:
                bits.append(f"內容有 {len(missing)} 行未見於投影片（"
                            + "／".join(missing[:3])
                            + ("…" if len(missing) > 3 else "") + "）")
            if extra:
                bits.append(f"投影片有 {len(extra)} 行不在內容中（"
                            + "／".join(extra[:3])
                            + ("…" if len(extra) > 3 else "") + "）")
            st.warning(f"⚠️ 「{label}」投影片歌詞與內容不符：" + "；".join(bits)
                       + "。可能是已上載／已產生的詩歌投影片尚未更新，或內容有誤"
                       "（換行／分段不同也可能造成差異）。")
        if lyric_diffs == [] and n_cmp:
            st.caption(f"📝 歌詞比對：{n_cmp} 首歌的投影片歌詞與內容一致。")
        # 用字比對：對照內建歌詞庫（lyrics/），提示可能用錯字的地方。
        word_diffs, n_word = _sanity_lib_word_check(week)
        for label, where, pairs in word_diffs:
            items = "\n".join(
                f"- 內容「{c}」／詩集「{l}」　差異：{d}" for c, l, d in pairs)
            st.warning(f"📖 「{label}」對照歌詞庫（{where}）可能有 "
                       f"{len(pairs)} 處用字不同：\n\n{items}\n\n"
                       "（詩集版本可能與你所用不同，請自行核對。）")
        if not word_diffs and n_word:
            st.caption(f"📖 用字對照：{n_word} 首歌與歌詞庫相符。")
        # 重複／幾近相同的投影片：同一首歌內幾乎一樣（只差一個字）的兩頁——可能是
        # 重複的段落、未修剪乾淨的第二個版本，或打錯字。完全相同的重複（副歌每次
        # 一樣）屬正常，不列出。
        dup_slides, n_dup = _sanity_slide_duplicates(week)
        for label, pairs in dup_slides:
            items = "\n".join(
                f"- 第 {a} 頁 ↔ 第 {b} 頁（相似度 {r * 100:.0f}%）　"
                f"差異：{d}" for a, b, r, d in pairs)
            st.warning(f"🔁 「{label}」有 {len(pairs)} 對幾乎相同的投影片"
                       f"（可能重複，或只差一個字）：\n\n{items}")
        if not dup_slides and n_dup:
            st.caption(f"🔁 重複投影片：{n_dup} 首歌沒有幾乎相同的頁。")
        # 第一主日：聖餐＋使徒信經 樣板頁。
        comm_ok, comm_note = _sanity_communion_check(week, infos)
        if not comm_ok:
            st.warning("⚠️ " + comm_note)
        elif comm_note:
            st.caption("📌 " + comm_note)
