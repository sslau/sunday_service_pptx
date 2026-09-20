# Sunday Service Slides · Quick Start

Browser tool: edit content, generate `Sunday_Service_<date>.pptx`. One page.

## 1. Pick the week
Side panel → **Service date**. New week? **Add week**, enter `YYYY.MM.DD`,
"Add and open".

## 2. Edit content (tab "Edit content")
- **1. Call to Worship** — reference + font sizes (body min 44pt) + paste the
  whole passage in one box (auto-splits). Or hit **"Load from Bible site"** to
  fetch it from bible.fhl.net in 整段 form with a `[書 章:節]` prefix per verse
  (無空版); a **Book** dropdown fills the reference for you. Supports
  和合本/新譯本…, ranges like `12:1-8`, and **multiple books** separated by
  `；`/`;`/newline (merged into one passage).
- **3. Scripture Reading** — same as above.
- **2. Hymns** — per-hymn card: title, English title, source, refrain, verses;
  tick "repeat refrain after every verse".
- **4./5. Sermon & Announcements** — `.pptx` files only; upload on the other tab.
- **6. Communion** — tick to add Communion + Apostles' Creed; auto-ticked on the
  first Sunday of the month.

## 3. Generate (tab "Generate slides")
- Optionally **upload** a section file per slot (Hymns / Sermon / Announcements).
  Uploaded = used as-is; empty = compiled from the edit page. Mixing is fine.
- Local only: expand "or enter a filename" to pull from the app folder,
  `data/downloads`, or `~/Downloads`.
- Optionally replace background images or attach announcement images.
- Check **section source status**: uploaded / saved / compiled.

## 4. Build
Side panel → **Save current content**, then **Generate full service**, then
**Download PPTX** (button appears under it). Output:
`Sunday_Service_<date>.pptx`.

## Good to know
- Scripture auto-splits by font-derived capacity; **44pt minimum**; lines never
  end/start on punctuation.
- Content lives in Google Sheets (local fallback `data/weeks.json`).
- "Reuse previously saved" (default on) reuses a section built earlier that day.
- Full manual: `user_guide.en.md` / `user_guide.md`.
