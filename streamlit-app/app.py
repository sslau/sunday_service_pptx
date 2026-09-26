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
import io
import os
import re
import tempfile
import uuid
import html
import zipfile
from datetime import date as _date

import streamlit as st

st.set_page_config(page_title="崇拜投影片編輯器", page_icon="✝",
                   layout="wide")

import bible  # noqa: E402
import gdrive  # noqa: E402
import onedrive  # noqa: E402
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


# --- 登入門（簡單密碼鎖）----------------------------------------------------
# 密碼放在 `.streamlit/secrets.toml` 的 `app_password`（被 .gitignore 排除，
# 部署時請在 Streamlit Cloud 的 Settings → Secrets 貼上同一欄位）。
# 沒有設定 app_password 時會放行，但顯示黃色提醒。

try:
    APP_PASSWORD = st.secrets.get("app_password") or None
except Exception:
    APP_PASSWORD = None


def _auth_ok():
    return bool(st.session_state.get("auth_ok", False))


if not _auth_ok():
    with st.container(border=True):
        st.subheader("🔒 崇拜投影片編輯器 — 請登入")
        pw = st.text_input("密碼", type="password", key=K("gate_password"),
                           placeholder="請輸入密碼")
        if st.button("登入", type="primary", key=K("gate_login")):
            if APP_PASSWORD is None:
                st.warning("尚未設定 app_password（在 Settings → Secrets）。"
                           "未設定密碼前先放行。")
                st.session_state["auth_ok"] = True
                st.rerun()
            elif pw == APP_PASSWORD:
                st.session_state["auth_ok"] = True
                st.rerun()
            else:
                st.error("密碼錯誤，請重試。")
    st.stop()


def sv(name, default=""):
    return st.session_state.get(K(name), default)


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
    """背景圖 selectbox + 上傳存檔 + 預覽 + 可選縮圖列（詩歌敬拜／詩歌回應共用）。"""
    current = sv(wkey, default)
    opts = song_bg_options(current)
    labels = {o: ("（依序：song_bg1…）" if not o
                  else os.path.basename(o)) for o in opts}
    st.selectbox("背景圖", opts,
                 index=opts.index(current) if current in opts else 0,
                 format_func=lambda o: labels.get(o, o),
                 key=K(wkey))
    st.caption("下拉選單僅支援文字，下方縮圖僅供確認所選背景")
    thumb_cols = st.columns(len(opts) if opts else 1)
    for col, opt in zip(thumb_cols, opts):
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
    st.caption("亦可直接放圖片進 media/（檔名 song_bg*.jpg 就會出現在選單）；"
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


def section_source(upload, name, section, date, use_saved):
    """Where a section's slides will come from, for the readiness indicator."""
    if upload is not None:
        fname = _uploaded_name(upload)
        if fname:
            return f"已上載檔案（合併）：{fname}", "ready"
        return "已上載檔案（合併）", "ready"
    path = resolve_deck_path(name) if name else None
    if path:
        return f"使用檔案：{os.path.basename(path)}", "ready"
    if use_saved:
        saved = saved_section_path(section, date)
        if saved:
            if name:
                return f"使用已上載：{os.path.basename(name)}（已儲存）", "saved"
            return f"使用已儲存：{os.path.basename(saved)}", "saved"
    return "由網頁內容編譯", "compile"


def section_ready(section, name_key, label, upload_key=None):
    """Edit-wizard caption when the section already has an uploaded/found pptx
    (mirrors the build-tab status), else None. `upload_key` is the file-uploader
    widget key so an uploaded deck's original filename can be shown."""
    up = st.session_state.get(K(upload_key)) if upload_key else None
    if up is not None:
        fname = _uploaded_name(up)
        if fname:
            return (f"✅ 已上載 {label} .pptx：**{fname}**"
                    "（製成整場時以該檔案為準）")
        return f"✅ 已上載 {label} .pptx（製成整場時以該檔案為準）"
    name = st.session_state.get(K(name_key))
    path = resolve_deck_path(name) if name else None
    if path:
        return (f"✅ 已有現成的 {label} .pptx：使用檔案 "
                f"{os.path.basename(path)}")
    if sv("use_saved", True):
        saved = saved_section_path(section, DATE)
        if saved:
            if name:
                return (f"✅ 已上載 {label} .pptx：**{os.path.basename(name)}**"
                        "（已儲存，製成整場時以該檔案為準）")
            return (f"✅ 已有現成的 {label} .pptx（已上載／已儲存），"
                    "製成整場時以該檔案為準。")
    return None


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
        except Exception as exc:
            st.error(f"產生失敗：{exc}")
    b = st.session_state.get(K(f"sec_bytes_{sec_key}"))
    if b and st.session_state.get(K(f"sec_date_{sec_key}")) == DATE:
        bc2.download_button(
            f"⬇️ 下載 {label} .pptx", data=b, file_name=dl_name,
            mime="application/vnd.openxmlformats-officedocument"
                 ".presentationml.presentation",
            use_container_width=True)


def _sa_from_secrets(secrets):
    sa = secrets.get("gcp_service_account") or {}
    if sa.get("client_email") and sa.get("private_key"):
        return dict(sa)
    return None


def _drive_fetch_section(section, date, folder_id):
    """Find the canonical `section` deck file for `date` inside a Drive folder
    and download it to the saved-deck path (same location uploads use), so the
    existing 來源狀態/merge flow picks it up automatically.  Returns the file
    name, or None when nothing matches (no download performed)."""
    sa = _sa_from_secrets(st_secrets or {})
    if not sa:
        raise RuntimeError(
            "未設定 Google 服務帳戶（secrets.toml 的 gcp_service_account）")
    want = os.path.basename(section_path(section, date))
    if not want:
        raise RuntimeError("不支援此節（無對應檔名）")
    files = gdrive.files_in_date_folder(sa, folder_id, date)
    picked = next((f for f in files if (f.get("name") or "") == want), None)
    if picked is None:
        variants = gdrive._date_variants(date)
        picked = next(
            (f for f in files
             if (f.get("name") or "").endswith(".pptx")
             and any(v in (f.get("name") or "") for v in variants)), None)
    if picked is None:
        return None
    path = section_path(section, date)
    ok, err = gdrive.download_to(sa, picked["id"], path)
    if not ok:
        raise RuntimeError(err)
    fname = picked.get("name") or ""
    st.session_state[f"{date}::build_{'ann' if section == 'announcements' else section}_name"] = \
        fname
    save_upload_name(section, date, fname)
    st.cache_data.clear()
    return fname


def _drive_fetch_video(folder_id):
    """Download the date's 家事MP4 (announcements_<date>.mp4) from the Drive
    date-subfolder into the persisted video store.
    Returns the file name, or None when nothing matches."""
    sa = _sa_from_secrets(st_secrets or {})
    if not sa:
        raise RuntimeError(
            "未設定 Google 服務帳戶（secrets.toml 的 gcp_service_account）")
    want = f"announcements_{DATE}.mp4"
    files = gdrive.files_in_date_folder(sa, folder_id, DATE)
    picked = next((f for f in files if (f.get("name") or "") == want), None)
    if picked is None:
        variants = gdrive._date_variants(DATE)
        picked = next(
            (f for f in files
             if (f.get("name") or "").endswith(".mp4")
             and any(v in (f.get("name") or "") for v in variants)), None)
    if picked is None:
        return None
    tmp = tempfile.NamedTemporaryFile(
        suffix=".mp4", prefix="gdrive_", delete=False).name
    try:
        ok, err = gdrive.download_to(sa, picked["id"], tmp)
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


def _gdrive_panel(section, label, is_video=False):
    """Expander that pulls the section's .pptx (or the 家事MP4 when is_video)
    straight from the Livestreaming Drive folder (per-date subfolder).
    Reusable across wizard steps."""
    with st.expander("☁️ 從 Google Drive 讀取%s投影片" % label):
        try:
            cfg_folder = (st.secrets.get("gdrive", {})
                          .get("folder", ""))
        except Exception:
            cfg_folder = ""
        if cfg_folder:
            folder_val = cfg_folder
            _p = (DATE or "").split(".")
            _mmdd = ("%s.%s%s" % (_p[0], _p[1], _p[2])
                     if len(_p) == 3 else DATE)
            want = (f"announcements_{DATE}.mp4" if is_video
                    else os.path.basename(section_path(section, DATE)))
            st.caption("📽 %s投影片會直接從雲端 **「Livestreaming」資料夾**"
                       "（依日期子資料夾，如 `%s`）讀取，不需手動上載。"
                       "預期檔案名稱：**`%s`**（或檔名含該日期）。"
                       % (label, _mmdd, want))
        else:
            folder_val = st.text_input(
                "Drive 資料夾連結（secrets.toml 尚未設定）",
                key=K("gdrive_folder"),
                placeholder="https://drive.google.com/drive/folders/…")
        if st.button("⚡ 讀取並儲存", type="primary",
                     key=K(f"gdrive_fetch_{section}")):
            fid = gdrive.folder_id_from(folder_val)
            if not fid:
                st.error("請貼上 Google Drive 資料夾連結"
                         "（須含 /folders/<ID>）")
            else:
                try:
                    if is_video:
                        name = _drive_fetch_video(fid)
                    else:
                        name = _drive_fetch_section(section, DATE, fid)
                except Exception as exc:
                    st.error(f"Google Drive 讀取失敗：{exc}")
                else:
                    if name:
                        st.success(f"已從 Drive 讀取：{name}")
                    else:
                        want = (f"announcements_{DATE}.mp4" if is_video
                                else os.path.basename(
                                    section_path(section, DATE)))
                        st.warning("資料夾中找不到符合 "
                                   f"{want}（或日期含 {DATE}）的檔案")


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


def _signal_clear(date, section):
    """Callback for the ✖ buttons in 製成選項與來源狀態: schedules clearing the
    section's uploaded/specified/saved deck on the next run."""
    st.session_state[f"{date}::clear_{section}"] = True


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
        "psalm_verses": "\n".join(psalm.get("verses") or []),
        "hymn_font_max": week.get("hymn_font_max", 48),
        "hymn_font_min": week.get("hymn_font_min", 48),
        "hymn_margin_in": week.get("hymn_margin_in", 0.83),
        "hymn_text_shadow": bool(week.get("hymn_text_shadow", True)),
        "call_font": week.get("call_font", "DFKai-SB") or "DFKai-SB",
        "hymns_finalized": bool(week.get("hymns_finalized", False)),
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


def _restore_lapsed_fields(date, week):
    """Re-instate stored values for editor widget keys that are absent from
    session state.  Streamlit deletes widget-created session entries as soon
    as the widget leaves the rendered tree (e.g. switching wizard steps), so a
    field whose step is not active would otherwise read as empty and wipe the
    stored week on the next save.  setdefault only fills missing keys, leaving
    values the user is currently editing untouched."""
    for name, val in _seed_values(week, date).items():
        st.session_state.setdefault(f"{date}::{name}", val)
    hymns = week.get("hymns") or []
    for i, h in enumerate(st.session_state.get(f"{date}::ui_hymns") or []):
        sw = hymns[i] if i < len(hymns) else {}
        base = f"{date}::hymn_{h['_id']}_"
        st.session_state.setdefault(base + "title", sw.get("title", ""))
        st.session_state.setdefault(base + "subtitle", sw.get("subtitle", ""))
        st.session_state.setdefault(base + "source", sw.get("source", ""))
        st.session_state.setdefault(base + "music", sw.get("music", ""))
        st.session_state.setdefault(base + "lyricist", sw.get("lyricist", ""))
        st.session_state.setdefault(base + "bg", sw.get("bg", ""))
        st.session_state.setdefault(base + "refrain",
                                    "\n".join(sw.get("refrain") or []))
        st.session_state.setdefault(
            base + "repeat", bool(sw.get("refrain_after_every_verse", True)))
        for j, stanza in enumerate(sw.get("verses", [])):
            st.session_state.setdefault(f"{base}verses_{j}",
                                        "\n".join(stanza))
    slides = (week.get("sermon") or {}).get("slides") or []
    for i, s in enumerate(st.session_state.get(f"{date}::ui_slides") or []):
        sw = slides[i] if i < len(slides) else {}
        base = f"{date}::slide_{s['_id']}_"
        st.session_state.setdefault(base + "title", sw.get("title", ""))
        st.session_state.setdefault(base + "body",
                                    "\n".join(sw.get("body") or []))
    rverses = (week.get("response") or {}).get("verses") or []
    ui_r = st.session_state.get(f"{date}::ui_resp") or []
    for j in range(max(len(ui_r), len(rverses), 1)):
        st.session_state.setdefault(
            f"{date}::resp_verses_{j}",
            "\n".join(rverses[j]) if j < len(rverses) else "")


def _seed_response_fields(date, week):
    """Seed the 詩歌回應 stanza list and per-stanza widget keys."""
    response = week.get("response") or {}
    verses = ["\n".join(v) for v in (response.get("verses") or [])] or [""]
    st.session_state[f"{date}::ui_resp"] = verses
    for j in range(max(len(verses), 20)):
        st.session_state.setdefault(
            f"{date}::resp_verses_{j}",
            verses[j] if j < len(verses) else "")


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
        for j, stanza in enumerate(h.get("stanzas", [])):
            st.session_state[f"{base}verses_{j}"] = stanza

_restore_lapsed_fields(DATE, week)


# ── clear-upload signals ────────────────────────────────────────────────────
# The ✖ buttons in 製成選項與來源狀態 set these flags; they are processed here
# at the top of the run — before the file_uploader widgets are instantiated —
# so their session keys can be safely popped and the page rerun.
_SECTION_KEYS = ("songs", "offering", "response", "sermon", "announcements")
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
    st.markdown(f"確定要移除 **{h.get('title') or '（未命名）'}** 嗎？"
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

    resp_verses = []
    for j in range(max(len(ui_resp), 1)):
        lines = split_lines(st.session_state.get(
            f"{date}::resp_verses_{j}", ""))
        if lines:
            resp_verses.append(lines)

    week = {
        "date": date,
        "psalm": {
            "ref": gv("psalm_ref", ""),
            "ref_size": int(gv("psalm_ref_size", 48)),
            "font_size": int(gv("psalm_font_size", 44)),
            "verses": split_lines(gv("psalm_verses", "")),
        },
        "hymn_font_max": int(gv("hymn_font_max", 48)),
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
    stanza_txt = [st.session_state.get(f"{base}verses_{j}", ui["stanzas"][j])
                  for j in range(max(len(ui.get("stanzas", ["", ])), 1))]
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
        "verses": [split_lines(t) for t in stanza_txt],
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


def _apply_parsed(uid_kind, uid, text, with_chorus=True):
    """Run _parse_hymn_paste and write the result into the session widget keys.

    uid_kind: 'hymn' (per-song uid) or 'resp' (single response hymn).
    uid: hymn `_id` or "" for the response section."""
    stanzas, refrain, repeat = _parse_hymn_paste(text, with_chorus=with_chorus)
    if uid_kind == "hymn":
        base = f"hymn_{uid}_"
        ui = _get_ui("hymns")
        h = next((x for x in ui if x.get("_id") == uid), None)
        if h is not None:
            h["stanzas"] = stanzas
            h["refrain"] = refrain
            h["repeat"] = repeat
    else:
        base = "resp_"
        ui = _get_ui("resp")
        ui[:] = stanzas
    for j in range(max(len(stanzas), 1)):
        st.session_state[K(f"{base}verses_{j}")] = stanzas[j]
    st.session_state[K(f"{base}refrain")] = refrain
    st.session_state[K(f"{base}repeat")] = repeat
    st.session_state[K(f"{base}chorus_applied")] = with_chorus
    st.rerun()



def _maybe_reapply_chorus(uid_kind, uid):
    """If the 副歌辨識正確 checkbox now disagrees with the last applied parse,
    re-run _apply_parsed with the new setting so each paragraph becomes its
    own slide (with_chorus=False) or the chorus grouping comes back on."""
    base = f"hymn_{uid}_" if uid_kind == "hymn" else "resp_"
    cur = st.session_state.get(K(f"{base}chorus_ok"))
    applied = st.session_state.get(K(f"{base}chorus_applied"))
    paste = st.session_state.get(K(f"{base}paste"), "")
    if cur is not None and applied is not None and cur != applied and paste.strip():
        _apply_parsed(uid_kind, uid, paste, with_chorus=cur)


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
            + "".join(body) + "</div>")
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


VERSION_OPTIONS = ["和合本", "和合本2010（和修版）", "新譯本"]
VERSION_NOTE = {"和合本": "和合本",
                "和合本2010（和修版）": "和合本2010", "新譯本": "新譯本"}

FONT_OPTIONS = [
    ("DFKai-SB", "DFKai-SB（標楷體）"),
    ("PMingLiU", "PMingLiU（新細明體）"),
    ("Microsoft JhengHei", "Microsoft JhengHei（微軟正黑體）"),
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


def _song_font_selector(label, hint=None):
    """歌詞字級下拉（48–54pt），label 與下拉同一行並緊貼。若舊週次存有
    範圍外的數值，自動補進選項，避免 selectbox 因值不在清單而報錯。"""
    cur = int(sv("hymn_font_max", 48) or 48)
    options = sorted(set(range(48, 55)) | {cur})
    c1, c2, _ = st.columns([1.3, 0.9, 5], vertical_alignment="center")
    c1.markdown(f"**{label}**")
    c2.selectbox(
        "歌詞字級",
        options=options,
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


def render_video_panel(ns="", num="D",
                       title="🎞 家事分享投影片 → MP4"):
    """The section-D MP4 panel, reusable in the 「製作成投影片」 tab (D) and as
    wizard step 9 (s9).  `ns` namespaces widget keys so both panels can render
    in the same run; the produced MP4 is stored under the shared video_data key
    so the sidebar 製成 flow always sees it."""
    section_header(num, title,
                   "上載 PNG/JPG 圖片或 ZIP，或選用週期內容，只取部分項目產製影片")
    st.caption("投影片在 **Server 端**以 Pillow 製成 H.264 影片"
               "（純 Python，雲端可執行，不佔本機軟體）。")
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
    video_imgs = st.file_uploader(
        "① 直接上載 PNG/JPG 圖片（一檔一頁）——最精確，畫面即所見",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True, key=K(f"video_imgs{ns}"))
    zip_file = st.file_uploader(
        "② 或 ZIP（內含投影片圖片；檔名編號 = 投影片編號，如 slide_1.png）",
        type=["zip"], key=K(f"video_zip{ns}"))
    vdate = date_from_filename(
        getattr(ann_pptx, "name", "") if ann_pptx is not None else "")
    src_kind, src_items, src_pptx_bytes = "none", [], None
    zip_map = {}
    if video_imgs:
        src_kind = "images"
        src_items = [u for u in video_imgs]
        st.caption(f"使用上載圖片：{len(src_items)} 張，一檔一頁。")
    elif zip_file is not None:
        src_kind = "zip"
        try:
            with zipfile.ZipFile(io.BytesIO(zip_file.getvalue())) as zf:
                entries = []
                for zi in zf.infolist():
                    if zi.is_dir() or "__MACOSX/" in zi.filename \
                            or ".DS_Store" in zi.filename:
                        continue
                    base = zi.filename.rsplit("/", 1)[-1].lower()
                    if not (base.endswith(".png") or base.endswith(".jpg")
                            or base.endswith(".jpeg")
                            or base.endswith(".webp")):
                        continue
                    entries.append((zi.filename, zf.read(zi)))
                if not entries:
                    st.caption("ZIP 內沒有 PNG/JPG/WEBP 圖片。")
                else:
                    maxnum = 0
                    for name, _b in entries:
                        m = re.search(r"(\d+)", name.rsplit("/", 1)[-1])
                        if m:
                            maxnum = max(maxnum, int(m.group(1)))
                    nextnum = maxnum + 1
                    for name, blob in entries:
                        m = re.search(r"(\d+)", name.rsplit("/", 1)[-1])
                        num = int(m.group(1)) if m else None
                        if num is None:
                            num = nextnum
                            nextnum += 1
                        if num in zip_map:
                            continue           # keep the first
                        zip_map[num] = blob
                        src_items.append(ZipSlideImage(
                            num, name.rsplit("/", 1)[-1], blob))
                    src_items.sort(key=lambda z: z.num)
                    st.caption(f"ZIP：{len(src_items)} 張圖片，"
                               "檔名編號 = 投影片編號（按編號排序）。")
        except Exception as exc:
            st.caption(f"無法讀取 ZIP：{exc}")
            src_items = []
            src_kind = "none"
    else:
        src_fname = (getattr(ann_pptx, "name", "")
                     if ann_pptx is not None else "")
        saved_ann_path = saved_section_path("announcements", DATE)
        cur_ann = (assemble_week(DATE).get("announcements") or [])
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
        "kind": src_kind, "items": src_items, "zip_map": zip_map,
        "pptx": src_pptx_bytes, "vdate": vdate, "ns": ns,
        "nums_auto": ",".join(
            str(getattr(z, "num", i + 1)) if src_kind == "zip" else str(i + 1)
            for i, z in enumerate(src_items))}
    if src_kind == "none":
        st.caption("沒有可用的家事分享來源：請上載圖片、ZIP、家事分享 .pptx，"
                   "或在「✏️ 編輯內容 → 家事分享」輸入至少一項。")
        st.text_input("投影片編號", value="", disabled=True,
                      key=K(f"video_nums_off{ns}"),
                      help="上載圖片／ppt 或輸入家事分享後即可產製")
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
                src_kind, src_items, zip_map, src_pptx_bytes,
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
            st.download_button(
                "⬇️ 下載 MP4", data=vid,
                file_name=f"announcements_{vdate or DATE}.mp4",
                mime="video/mp4", key=K(f"video_dl{ns}"))
            st.caption(f"已產生影片（約 {len(vid) / 1048576:.1f} MB）。")
            first_toggle = st.checkbox(
                "把影片併入整場投影片的第一頁\n"
                "（自動循環播放，直到點擊下一頁）",
                value=True, key=K(f"video_first_deck{ns}"))
            save_video_meta(DATE, bool(first_toggle))
            if first_toggle:
                st.caption("記得按左上「製成完整投影片」的按鈕，"
                           "影片才會被放入第一頁。")


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
    st.text_input("出處", value=sv(f"{prefix}_ref"),
                  key=K(f"{prefix}_ref"),
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
    c1.number_input("出處字型(pt)", value=48, min_value=20,
                    max_value=90, key=K(f"{prefix}_ref_size"))
    c2.number_input("內文字型(pt)", value=44, min_value=44,
                    max_value=90, key=K(f"{prefix}_font_size"))
    st.text_area("內容（貼上全部經文，自動分頁）",
                 value=sv(f"{prefix}_verses"),
                 key=K(f"{prefix}_verses"), height=200)
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
                    "verses": split_lines(sv("scrflow_verses", "")),
                },
            })
            data, count = build_section_pptx_bytes("scripture", week)
            st.session_state["scrflow_bytes"] = data
            st.session_state["scrflow_count"] = count
            st.success(f"完成：{count} 頁（獨立流程，只含讀經經文）")
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
        ("Microsoft JhengHei", "Microsoft JhengHei（微軟正黑體）"),
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
            _scripture_editor("psalm", "psalm")

    elif step == 2:
        with st.container(border=True):
            section_header("2", "詩歌敬拜（Hymns）")
            _odrive_panel("songs", "詩歌")
            _gdrive_panel("songs", "詩歌")
            section_deck_panel("songs", "hymns", "songs", "詩歌",
                               f"songs_slides_{DATE}.pptx",
                               btn="🎵 產生詩歌投影片")
            ready = section_ready("songs", "build_songs_name", "詩歌", "build_songs")
            st.caption(ready or "以下內容以網頁內容現場編譯；上載 .pptx 時"
                       "以檔案為準，字型／項目符號沿用來源檔。")
            st.checkbox("歌詞文字加上陰影（有助背景圖上閱讀，寫入投影片）",
                        value=sv("hymn_text_shadow", True),
                        key=K("hymn_text_shadow"))
            _song_font_selector(
                "歌詞字級（pt）",
                hint="固定字級（48–54pt），不隨行長縮小；行太長時自動換行。"
                     "標點符號以空格取代（如「，」「。」→ 空白）。"
                     "每首詩歌可用下方「背景圖」各自指定背景。")

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
                    _bg_picker(f"hymn_{uid}_bg", h.get("bg", ""))
                    st.text_area("貼上整首詩歌（協議：每段之間 1 個空行＝分 slide）",
                                 value="", key=K(f"hymn_{uid}_paste"),
                                 height=100)
                    st.caption("貼上歌詞後，按下方「匯入為各段」按鈕，"
                               "即可把整首詩歌切成各段（每段＝一張 slide）。")
                    if st.button("匯入為各段",
                                 key=K(f"hymn_{uid}_apply_paste")):
                        _apply_parsed("hymn", uid,
                                      st.session_state.get(
                                          K(f"hymn_{uid}_paste"), ""),
                                      with_chorus=False)
                    for j in range(max(len(h.get("stanzas", ["", ""])), 1)):
                        st.text_area(f"第 {j + 1} 段（每行一句）",
                                     value=sv(f"hymn_{uid}_verses_{j}",
                                              h["stanzas"][j]
                                              if j < len(h["stanzas"]) else ""),
                                     key=K(f"hymn_{uid}_verses_{j}"), height=90)
                    st.divider()
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
                    c1, c2 = st.columns(2)
                    if c1.button("＋ 新增一段", key=K(f"hymn_{uid}_addstanza")):
                        h["stanzas"].append("")
                    if c2.button("移除這首詩歌", key=K(f"hymn_{uid}_del")):
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
            _gdrive_panel("offering", "獻詩")
            section_deck_panel("offering", None, "offering", "獻詩", None)
            ready = section_ready("offering", "build_offering_name", "獻詩", "build_offering")
            if ready:
                st.caption(ready)
            else:
                st.info("此節只接受現成 .pptx：上載後製成整場時以檔案為準；"
                        "未上載則此節不加入。上載檔在 Server 重啟後可能遺失，"
                        "請當日完成整場製作。")

    elif step == 4:
        with st.container(border=True):
            section_header("4", "讀經經文", "貼上整段經文，自動分頁")
            _scripture_editor("scripture", "scripture")

    elif step == 5:
        with st.container(border=True):
            section_header("5", "講道信息（Sermon）")
            _odrive_panel("sermon", "講道")
            _gdrive_panel("sermon", "講道")
            section_deck_panel("sermon", "sermon", "sermon", "講道",
                               f"sermon_{DATE}.pptx",
                               btn="🗣 產生講道投影片")
            ready = section_ready("sermon", "build_sermon_name", "講道", "build_sermon")
            if ready:
                st.caption(ready)
            else:
                st.info("此節以現成 .pptx 為準；上載後製成整場時以該檔案為準。"
                        "上載檔在 Server 重啟後可能遺失，請當日完成整場製作。")

    elif step == 6:
        with st.container(border=True):
            section_header("6", "詩歌回應（Response Hymn）",
                           "單首回應詩歌，緊接信息之後（可選）")
            _odrive_panel("response", "詩歌回應")
            _gdrive_panel("response", "詩歌回應")
            section_deck_panel("response", "response", "response", "詩歌回應",
                               f"response_{DATE}.pptx",
                               btn="🎶 產生詩歌回應投影片")
            ready = section_ready("response", "build_response_name", "詩歌回應",
                              "build_response")
            st.caption(ready or "此節**可選**：不需要可直接按「下一步」略過。"
                       "要加入時，可填寫以下內容現場編譯；"
                       "或在上方直接上載現成 .pptx 合併（以檔案為準）。")
            _song_font_selector(
                "歌詞字級（pt）",
                hint="與「詩歌敬拜」共用同一設定（48–54pt）。")
            ui_resp = st.session_state.setdefault(f"{DATE}::ui_resp", [""])
            for name in ("resp_title", "resp_subtitle", "resp_source",
                         "resp_music", "resp_lyricist", "resp_bg"):
                st.session_state.setdefault(K(name), "")
            c1, c2 = st.columns(2)
            c1.text_input("曲名", value=sv("resp_title"), key=K("resp_title"))
            c2.text_input("英文名", value=sv("resp_subtitle"),
                          key=K("resp_subtitle"))
            m1, m2, m3 = st.columns(3)
            m1.text_input("詩集", value=sv("resp_source"), key=K("resp_source"))
            m2.text_input("曲", value=sv("resp_music"), key=K("resp_music"))
            m3.text_input("詞", value=sv("resp_lyricist"),
                          key=K("resp_lyricist"))
            _bg_picker("resp_bg", "")
            st.text_area("貼上整首詩歌（協議：每段之間 1 個空行＝分 slide）",
                         value="", key=K("resp_paste"), height=100)
            st.caption("貼上歌詞後，按下方「匯入為各段」按鈕，"
                       "即可把整首詩歌切成各段（每段＝一張 slide）。")
            if st.button("匯入為各段", key=K("resp_apply_paste")):
                _apply_parsed("resp", "", st.session_state.get(K("resp_paste"), ""),
                              with_chorus=False)
            for j in range(max(len(ui_resp), 1)):
                st.session_state.setdefault(
                    K(f"resp_verses_{j}"),
                    ui_resp[j] if j < len(ui_resp) else "")
                st.text_area(f"第 {j + 1} 段（每行一句）",
                             value=sv(f"resp_verses_{j}",
                                      ui_resp[j] if j < len(ui_resp) else ""),
                             key=K(f"resp_verses_{j}"), height=90)
            c1, c2 = st.columns(2)
            if c1.button("＋ 新增一段", key=K("resp_add")):
                ui_resp.append("")
                st.rerun()
            if c2.button("移除最後一段", key=K("resp_del")) and len(ui_resp) > 1:
                ui_resp.pop()
                st.rerun()
            _finalize_stamp("resp_finalized", "詩歌回應")

    elif step == 7:
        with st.container(border=True):
            section_header("7", "聖餐（Communion）")
            st.checkbox("聖餐（第一主日：加聖餐＋使徒信經頁）",
                        value=sv("communion", is_first_sunday(DATE)),
                        key=K("communion"))
            if is_first_sunday(DATE):
                st.caption("此日期為當月第一主日，已預設勾選聖餐＋使徒信經。")

    elif step == 8:
        with st.container(border=True):
            section_header("8", "家事分享（Announcements）")
            _odrive_panel("announcements", "家事分享")
            _gdrive_panel("announcements", "家事分享")
            section_deck_panel("announcements", "announcements",
                               "announcements", "家事分享",
                               f"announcements_{DATE}.pptx",
                               btn="📋 產生家事分享投影片",
                               widget="ann")
            ready = section_ready("announcements", "build_ann_name", "家事分享", "build_ann")
            if ready:
                st.caption(ready)
            else:
                st.info("此節以現成 .pptx，或以已輸入的家事分享內容編譯；"
                        "上載後以檔案為準。")

    elif step == 9:
        with st.container(border=True):
            _odrive_panel("video", "家事MP4", is_video=True)
            _gdrive_panel("video", "家事MP4", is_video=True)
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
        _wmap = {"songs": "songs", "offering": "offering",
                 "response": "response", "sermon": "sermon",
                 "announcements": "ann"}
        for label, sec in (
                ("🎵 詩歌敬拜", "songs"),
                ("🙌 獻詩", "offering"),
                ("🎶 詩歌回應", "response"),
                ("🗣 講道信息", "sermon"),
                ("📋 家事分享", "announcements")):
            up = st.session_state.get(K("build_" + _wmap[sec]))
            name = st.session_state.get(K("build_" + _wmap[sec] + "_name"))
            text, kind = section_source(up, name, sec, DATE, use_saved)
            icon = {"ready": "✅", "saved": "💾", "compile": "🧩"}[kind]
            r1, r2 = st.columns([0.78, 0.22], vertical_alignment="center")
            r1.markdown(f"{icon} **{label}** — {text}")
            if kind != "compile":
                r2.button("✖ 移除", key=K(f"clear_{sec}"),
                          use_container_width=True,
                          on_click=_signal_clear, args=(DATE, sec),
help="移除該節的上載並刪除已儲存的 pptx，"
                                "恢復用網頁內容編譯")
        if sv("communion", is_first_sunday(DATE)):
            st.markdown("✅ **聖餐＋使徒信經** — 已加入")
        else:
            st.markdown("⬜ **聖餐＋使徒信經** — 未加入"
                        + ("（第一主日建議勾選）"
                           if is_first_sunday(DATE) else ""))
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
