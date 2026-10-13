# Sunday Service Slides Editor · User Guide

Streamlit web app. Edit each Sunday's worship content in the browser and generate
and download the full-service PowerPoint (`.pptx`) right from the page. No
PowerPoint install and no local server required.

- **Storage:** week content lives in Google Sheets (the `weeks` worksheet); when
  no Sheets credentials are configured, a local run falls back to
  `data/weeks.json`.
- **Output:** the `.pptx` (and optional announcements `.mp4`) downloads straight
  from the browser, so nothing is kept on the server disk.

> This document matches the current code. If a button is not on screen, that
> section currently works mainly by uploading a ready-made file (see each
> section).

---

## 1. Getting started

1. Open the app URL in a browser.
2. If `app_password` is set, a **login screen** appears first; type the password
   and press "登入" (Log in) or Enter.
   - The password is set in `.streamlit/secrets.toml` (`app_password`), and on
     Streamlit Cloud under Settings → Secrets.
   - Opening the app on the same machine (**`localhost` / `127.0.0.1` / `::1`)
     skips the password**. The password gate is a convenience, not a security
     boundary, because the host header can be spoofed.
   - With no `app_password` set, the app stays open but shows a reminder.
   - Each login attempt is written to the `logs` worksheet; a summary of
     important actions per login session goes to the `sessions` worksheet
     (login time, IP, action count, running summary).
3. You land on the **✏️ 編輯內容** (Edit content) tab, on the **newest**
   service date. If there is no data yet, a default week is created.

### Sidebar

From top to bottom:

| Control | Purpose |
| --- | --- |
| ＋ 新增週次 (Add week) | Enter `YYYY.MM.DD` and click "新增並開啟" (Add and open) to create and switch to a new week. |
| 🔄 重新載入 Google Sheet (Reload) | Clear the cache and re-read the sheet now (use after editing the sheet by hand). |
| 主日日期 (Service date) | Switch the Sunday being edited; the edit page follows the selection. |
| 💾 儲存目前內容 (Save current content) | Save everything on the page back to the data store (Sheets/local JSON). |
| 📽 製成整場投影片（合併＋編譯）(Generate full service) | Build the whole `Sunday_Service_<date>.pptx`; on success a "⬇️ 下載 PPTX" (Download) button appears below. |
| 📜 經文投影片（獨立流程）(Scripture-only, separate flow) | Open a one-step page that makes a scripture-only deck; its content is **not** saved to Sheets and is not part of the full-service flow. |

> **Remember to click "💾 儲存目前內容".** While you type, edits live in the
> browser session only; save before switching dates or logging back in. (Using
> "從 fhl.net 聖經網輸入" to fetch scripture auto-saves once.)

---

## 2. A Sunday in a nutshell

1. Pick the **service date** on the left (or "＋ 新增週次" to add one).
2. On the "✏️ 編輯內容" tab, fill the sections using the **nine step buttons**
   at the top (any order works).
3. Where a section needs a ready-made sermon/announcements file, **upload the
   `.pptx`** in that step (or fetch it from OneDrive).
4. Go to the "📽 製作成投影片" tab and check "步驟狀態" (section status).
5. In the sidebar click "💾 儲存目前內容", then "📽 製成整場投影片".
6. Click "⬇️ 下載 PPTX" to get `Sunday_Service_<date>.pptx`.

---

## 3. Edit content (tab 1)

The edit page is a **wizard** showing one section at a time. The nine step
buttons at the top let you jump around in **any order**; below are "⬅ 上一步 /
下一步 ➡" (Previous / Next). At the very top you first pick the
**global typeface** (DFKai-SB / PMingLiU / JhengHei / DengXian / LiSu), applied
to all slide text.

### Step 1 · Call to Worship (宣召經文)

- **書卷 (Book):** picking a book fills the FHL book code into **出處
  (Reference)**; you can also type the reference directly, e.g. `詩篇 34:1-3`
  (Psalm 34:1-3).
- **譯本版本 (Translation):** 和合本 / 和合本2010（和修版）/ 新譯本. Selecting a
  version appends `（版本）` to the reference and re-fetches. You can also write
  a version note directly (several translations are supported).
- **從 fhl.net 聖經網輸入 (Load from Bible site):** fetches the passage from
  bible.fhl.net into **內容 (Content)** and auto-saves once on success.
  - **Multiple books:** separate with `；`, `;` or a newline, e.g.
    `詩篇 34:1-3；馬太福音 6:9-13`. Each book becomes its own paragraph headed
    by its "Book chapter:verse" label.
  - Ranges like `12:1-8` and cross-chapter `1:1-2:3` are supported; books accept
    full or short Chinese names (詩篇/詩, 羅馬書/羅 …).
- **出處字型 / 內文字型 (pt) (Reference / body font):** reference 20–90; body
  **minimum 44pt**, maximum 90.
- **經文排版方式 (Layout):**
  - **整段 (Paragraph):** verses run together continuously (verse numbers shown
    as superscripts; good for the call to worship).
  - **列表 (List):** one verse per line; page breaks only fall between verses
    (good for a longer reading).
- **內容 (Content):** paste the whole passage in one box; the app **splits it
  across slides automatically** (section 7).
- Top-right "🌐 從 OneDrive 讀取宣召經文投影片": with OneDrive configured, fetch
  `psalm_<date>.pptx` to replace the compiled output.

### Step 2 · Hymns (詩歌敬拜)

Each hymn keeps its lyrics in **one editor box**, shown right inside the card.
What you type is what lands on the slides: **a blank line separates slides; each
line inside a stanza is one slide line.** The box shows a running stanza count
and changes apply immediately — no separate import button, no per-stanza boxes.
Add a stanza by leaving a blank line, delete one by clearing it.

- **歌詞庫搜尋 (Lyric-library search)** inside each hymn card:
  - Type keywords, e.g. "中國啊" or "讓我們 愛". **Space-separated terms must all
    appear on the same slide**; matching ignores punctuation and spaces, so
    lines/sentences broken across slides still match.
  - Results are **grouped by song** (a song appears once), showing
    "Title · Collection · hits — lyric snippet".
  - Selecting a song opens a read-only **全首歌詞 (Whole lyrics)** pop-up with
    the search terms highlighted.
  - "**⬇️ 匯入整首詩歌 (Import whole song)**" drops the whole song (page-leading
    verse numbers already stripped) straight into the lyric editor box and
    fills **Title / English title / Songbook / Music / Lyrics** from slide 1 (a
    title card is not counted as a stanza); edit it in the box afterwards.
  - The library holds 4 collections: 教會聖詩, 粵語其他詩歌, 美樂頌, 輔助本.
- Per-hymn fields: **曲名 (Title), 英文名 (English title), 詩集 (Songbook),
  曲 (Music), 詞 (Lyrics)**.
- **背景圖 (Background):** expand "🎨 背景圖" and tick the checkbox **under a
  thumbnail** to apply it (one per hymn; the first cell "依序" = the default
  `media/song_bg1…` order). You can also upload a new background, or drop an
  image into `media/` (named `song_bg*.jpg` and it appears in the strip).
- **歌詞字級 (pt) (Lyric font size):** fixed size (**default 52pt**, 48–52pt
  selectable; it does not shrink to fit; long lines wrap).
- **歌詞文字加上陰影 (Text shadow):** keeps lyrics readable over background
  images and is written into the deck.
- **▶ 預覽這首詩歌的 slides (Preview):** "🎬 產生這首詩歌的 slide" renders an
  inline preview of each page (text and sizes, title card included) and offers
  that hymn's own `.pptx` for download.
- **Upload a hymns `.pptx`:** an uploaded file wins for the whole section (its
  fonts/bullets are preserved); leave it empty to compile live from the page.
  "🎵 產生詩歌投影片" compiles a standalone file from the page.
- "＋ 新增詩歌" adds a hymn; "🗑 移除這首詩歌" asks for confirmation first.
- "✔️ 標記『詩歌敬拜』此節為定稿 (Mark finalized)": a marker only.

> The "本週詩歌一覽" strip shows each hymn's status (e.g. "⚠ 未填曲名" no title,
> "⚠ 沒有歌詞" no lyrics, or "N 段") with a "前往編輯" (Go to edit) button.

### Step 3 · Offering (獻詩)

- This section accepts a ready-made **`.pptx` only**: upload or fetch from
  OneDrive; the file wins when building. Left empty, the section is skipped.
  Uploaded files may be lost when the server restarts, so finish the full build
  the same day.
- In the full deck, Offering comes **right after the prayer / Lord's Prayer**.

### Step 4 · Scripture Reading (讀經經文)

- Identical controls to the Call to Worship (book dropdown, translation, fhl.net
  fetch, fonts, paragraph/list layout, auto-splitting); the reading is just
  usually longer.

### Step 5 · Sermon (講道信息)

- This section is **built from a ready-made `.pptx`**: upload it here, or fetch
  `sermon_<date>.pptx` from OneDrive.
- "🗣 產生講道投影片" compiles a standalone file from the page's sermon content
  (when the week already has any).

### Step 6 · Response Hymn (詩歌回應)

- **Optional:** skip with "下一步" when not needed (the section is not added).
- A single response hymn with the same fields as Hymns: **one lyric editor box**
  (a blank line separates slides, with a running stanza count), Title / English
  title / Songbook / Music / Lyrics, background, lyric-library search.
  **The lyric font size is shared with Hymns.**
- You can also **upload a ready-made `.pptx`** (the file wins); otherwise the
  page content is compiled live.
- "✔️ 標記『詩歌回應』此節為定稿" as above.
- In the full deck the response hymn sits **after the sermon and before
  communion** (its header image is fixed to `media/response_hymn.jpg`).

### Step 7 · Communion (聖餐)

- Ticking "聖餐" adds a **Communion + Apostles' Creed** pair after the sermon /
  response.
- It is **pre-ticked on the month's first Sunday**.
- If you fetch `communion_<date>.pptx` from OneDrive, it **replaces** the built-in
  Communion + Creed pair (and is used even when the checkbox is off).

### Step 8 · Announcements (家事分享)

- This section is **built from a ready-made `.pptx`**: upload it here, or fetch
  `announcements_<date>.pptx` from OneDrive; "📋 產生家事分享投影片" compiles it
  from the page.
- In the full deck, announcements sit after the doxology and benediction, second
  to last (the "closing" slide is last).

### Step 9 · Announcements → MP4 (家事分享 → MP4)

Turn announcements into a video that can **join the full deck as slide 1** and
loop automatically until the next click. Two ways:

1. **Upload an MP4 exported by PowerPoint** (recommended): open the deck →
   File → Export → Create a Video (or "Export to video"), choose 1080p, wait,
   then upload the `.mp4` under "① 直接上載 PowerPoint 匯出的 MP4".
2. **Generate it on the server:** if you don't have a video yet, build an H.264
   MP4 in pure Python (Pillow + imageio-ffmpeg) from:
   - an uploaded/saved announcements `.pptx`, or
   - the week's announcements content.

   Set **投影片編號 (slide numbers)** (e.g. `1,3,5` or `1-3,5`), **每張停留秒數
   (seconds per slide)** (1–30) and **fps** (15/20/24/30), then click
   "🎞 產生影片（MP4）".

Afterwards:

- "⬇️ 下載 MP4" downloads `announcements_<date>.mp4`.
- Ticking "**把影片併入整場投影片的第一頁**（自動循環播放，直到點擊下一頁）" and
  then clicking "📽 製成整場投影片" places the video as the **first slide** (before
  the welcome slide).

---

## 4. Make the slides (tab 2)

This tab is for **checking where each section comes from** and setting
announcement images. The actual "generate" button is in the sidebar. (Uploads
and per-section "generate" actions live inside the Edit-content steps.)

- **步驟狀態 (Section status, A):** per-section source:
  - `✅` — uploaded / fetched from cloud / saved (the file wins), with a
    "✖ 移除" (Remove) button.
  - `🧩` — compiled from the page.
  - Scripture sections with no file show their reference and line count (or
    "未填寫" not filled).
  - Communion shows whether it is included; the announcements MP4 shows whether
    it was generated and whether it joins slide 1.
- **合併先前各步驟投影片檔 (Merge previously generated section decks)** (on by
  default): when on, the full build merges the section `.pptx` files generated/
  uploaded earlier instead of recompiling; turn it off to compile everything
  from the page.
- **家事分享圖片 (Announcement images, optional):** for image-based
  announcements, attach a jpg/png/webp per item.

When everything looks right, click "📽 製成整場投影片（合併＋編譯）" in the
sidebar. On success you'll see "完成：N 頁" and a "⬇️ 下載 PPTX" button below.

- **🔍 檢查整場投影片 (Check the whole deck) (Sanity check —
  Reading↔Sermon · Song length + lyrics + wording · First-Sunday templates):**
  after building, click this button to check at once:
  - **Reading ↔ Sermon reference:** that the 讀經經文 (Scripture reading)
    reference matches the 講道信息 (Sermon) reference — the sermon's verse is
    taken from one of the **first four slides** of the sermon section. It
    checks both the content fields and the generated deck (including sections
    merged from `.pptx`); it warns when a reference is missing or the passages
    differ, and reports a match when they agree (or one is contained in the
    other).
  - **Song length:** **every song** — 詩歌敬拜 hymns, the 獻詩 (offering) and the
    詩歌回應 (response) — should be at least 4 slides including its **title
    card**; shorter songs are flagged (usually too few lyric stanzas, or a song
    with no lyrics yet). Counts are taken from the **generated slides** — the
    saved/uploaded section deck when one is merged, else compiled from the
    content (獻詩 has no content fields, so it is counted only from its deck).
    Slides generated here are split per song by their `N/M` page labels; an
    **imported** deck (no page labels) is split on its larger-font title cards,
    skipping the 詩歌敬拜／獻詩／詩歌回應 section headers.
  - **Lyrics comparison:** the words actually on each song's slides are compared
    line by line with the lyrics entered / **corrected** in the content fields
    (punctuation and line wrapping ignored). When they differ — e.g. a lyric was
    corrected but the slides were not regenerated, or an uploaded song pptx is
    wrong — it lists how many lines are "in the content but not on the slides"
    and "on the slides but not in the content", with examples. 詩歌敬拜／詩歌回應
    have content fields to compare; 獻詩 has none, so it is not compared.
  - **Wording check (against the built-in lyric library):** each song's lyrics
    are compared line by line with the same-titled hymn in the built-in library
    (`lyrics/教會聖詩` / 粵語其他詩歌 / 美樂頌 / 輔助本) to spot **likely wrong
    characters** — a line that closely resembles a library line but is not
    identical is listed as "content … / hymnal …" with the differing characters
    (e.g. `〔赦→赫〕`). This is a **hint**, not a hard error: the hymnal edition
    may differ from yours, so please verify. Punctuation, spaces and line
    wrapping are ignored, and 祢/祂 are read as 你/他. Songs with no library
    match (or no overlapping lyrics) are skipped.
  - **First-Sunday templates:** on the month's first Sunday, it confirms the
    deck contains the 聖餐 (Communion) + 使徒信經 (Apostles' Creed) template
    pages and warns when they are missing (ignore if you replaced the built-in
    templates with a cloud `communion_<date>.pptx`).
  The button works anytime — before a deck is built it estimates from the
  content fields.

---

## 5. Full-deck slide order

(After opening `Sunday_Service_<date>.pptx`; if the video is used it is placed
first of all.)

Fixed slides + sections:

1. Opening (welcome)
2. Call-to-worship header → **Call to Worship (宣召經文)**
3. Worship header → **Hymns (詩歌敬拜)**
4. Prayer header → Lord's Prayer
5. **Offering (獻詩)** (if uploaded)
6. Scripture header → **Scripture Reading (讀經經文)**
7. Sermon header → **Sermon (講道信息)**
8. **Response Hymn (詩歌回應)** (if used)
9. **Communion + Apostles' Creed** (if ticked or a cloud file exists)
10. Doxology → Benediction
11. Announcements header → **Announcements (家事分享)**
12. Closing

Section header backgrounds come from `media/` (replace via `media/` or the
relevant background settings).

---

## 6. Scripture-only flow

The sidebar "📜 開啟『經文投影片』獨立流程" opens a **one-step** page: choose the
global font, enter a reference (fetchable from fhl.net) or paste the whole
passage, then "📜 製成經文投影片 pptx" to download `scripture_<date>.pptx`.

- Its content uses separate `scrflow_*` session keys and is **not** saved to
  Google Sheets.
- The generated file is standalone and is **not** read or merged by the full
  build.

---

## 7. How scripture auto-splitting works

- The pasted passage is treated as one continuous string; the app estimates the
  characters per line from the **font size**, derives lines per slide from the
  text-box height, and splits into slides.
- **Body font minimum 44pt** (default 44pt, adjustable upward).
- Splits avoid punctuation: a slide (and each line) **never ends on
  punctuation**, and the next slide never starts on one.
- Chinese has no spaces between words, so breaks fall mid-sentence.
- **整段 (Paragraph)** = continuous flow with superscript verse numbers;
  **列表 (List)** = one verse per line, breaks only between verses.

---

## 8. Fetching from OneDrive (optional)

If an admin has set `[onedrive]` in Secrets (tenant_id / client_id /
client_secret / site_host / site_path / folder_path), each section's
"🌐 從 OneDrive 讀取…投影片" panel lists that date's folder and downloads after
you click "⚡ 從 OneDrive 讀取並儲存".

- Expected names are the section's canonical filenames (e.g.
  `psalm_<date>.pptx`, `sermon_<date>.pptx`, `announcements_<date>.mp4`); folder
  and filenames may spell the date as `YYYY.MM.DD` / `YYYY-MM-DD` / `YYYYMMDD` /
  `YYYY.MMDD`.
- The announcements MP4 is searched across the whole folder tree (including
  subfolders).

---

## 9. Storage and data

- **Week content:** written to the Google Sheets `weeks` worksheet (`date | json
  | updatedAt`). Streamlit Community Cloud's filesystem is ephemeral, wiped on
  every restart/redeploy, so content is not kept on server files. With no Sheets
  configuration, a local run falls back to `data/weeks.json`.
- **Cache:** sheet reads are cached for ~30 seconds; after editing the sheet by
  hand, click "🔄 重新載入 Google Sheet" to re-read immediately.
- **Login records:** the `logs` (each login attempt) and `sessions` (per-session
  action summary) worksheets; with no credentials, local files
  `data/login_log.json` and `data/sessions.json`.
- **Per-section generated files and video:** stored under
  `data/decks/<date>/` on the server. On the Cloud these are cleared on restart;
  locally they persist, and the same day's full build reuses them automatically
  (unless "合併先前各步驟投影片檔" is turned off).
- **Uploaded images** (e.g. image-based announcements, custom hymn backgrounds)
  are **session-only** and not saved back to the sheet; re-upload after a
  restart.
- **New-week defaults:** a new week is seeded from `sample_week.json` (Psalm
  34:1-3, Romans 12:1-8, three empty hymns) which you can overwrite; if the date
  is the month's first Sunday, communion is pre-ticked.

---

## 10. FAQ

**Q: PowerPoint says a merged deck "needs repair"?**
A: The merge imports each section as-is (with its layouts/master/theme/images)
and fixes master/layout ID collisions. If it still happens, check that the source
`.pptx` opens cleanly in PowerPoint on its own.

**Q: Do I still need to fill the edit page after uploading a file?**
A: No. Once a section has an uploaded file, that file wins when building; mixing
is fine (e.g. hymns from a file, reading compiled from the page).

**Q: I changed the page but the output didn't change?**
A: That section may have an uploaded file, or "合併先前各步驟投影片檔" reused an
old file. Clear it with "✖ 移除" on "製作成投影片 → 步驟狀態", or untick the
option, then build again.

**Q: The download button is gone?**
A: It only shows the most recent build for the current date. Click "📽 製成整場
投影片" again.

**Q: After refreshing, are the uploaded files / video still there?**
A: Uploaded and generated section files are saved under `data/decks/<date>/` so
they survive a page refresh; on the Cloud a server restart clears them, so
re-upload or regenerate.

**Q: Some fields lose my typing after switching steps?**
A: This version keeps unsaved edits across wizard steps (a per-date draft), but
**always click "💾 儲存目前內容" before switching dates or logging back in**.

**Q: Do I need to save again after fetching scripture?**
A: A successful fhl.net fetch auto-saves once; changes to other fields still need
a manual save.

**Q: Why didn't communion appear?**
A: Communion is only added when ticked; it is **pre-ticked on the month's first
Sunday**. Otherwise tick it in "Step 7 · Communion".

---

_Filename: `user_manual.en.md`. 繁體中文版：`user_manual.md`._
