# Sunday Service Slides Editor · User Guide

A browser-based tool for editing the weekly worship content and generating the
service `.pptx`. Content is stored in Google Sheets; the generated file
downloads straight from the browser, so you do not need PowerPoint installed or
a local build server.

---

## 1. Getting started

1. Open the app URL and sign in (access is controlled by the Streamlit
   Community Cloud account/permissions).
2. On first load you land on the newest date in the **Service date** list. If
   there is no data yet, a default week is created.
3. The side panel, top to bottom:

| Control | Purpose |
| --- | --- |
| Service date | Switch between Sundays; the edit page follows the selection. |
| Save current content | Save everything on the current page back to the data store (Sheets/local). |
| Generate full service (merge + compile) | Build the complete `Sunday_Service_<date>.pptx`. |
| Add week | Enter `YYYY.MM.DD` and click "Add and open" to create a new week. |

> After a successful build, the download button appears directly beneath the
> "Generate full service" button.

---

## 2. Edit content (tab 1)

The top of the page has two side-by-side scripture blocks, then hymns, then the
sermon, announcements, and communion.

### 1. Call to Worship (Psalm)
- **Reference**: e.g. `Psalm 34:1-3`, or pick the book from the **Book**
  dropdown above it and it fills in the reference for you.
- **Load from Bible site**: after entering a reference, click this button to
  fetch the passage from bible.fhl.net (信望愛) into the content box as one
  continuous paragraph with a `[書 章:節]` prefix per verse, in 無空版 (no
  spacing before 神) — the same view as the site's `&TABFLAG=2` + `SSS=2`. You
  can still edit it or paste text yourself. Use a full or short Chinese book
  name (詩篇／詩, 羅馬書／羅, 約翰一書／約一 …); ranges like `12:1-8` and
  `1:1-2:3` (cross-chapter) are supported. **Fetch several books at once** by
  separating references with `；`, `;` or a newline, e.g.
  `詩篇 100:1-2；約翰福音 3:16` — each book becomes its own paragraph headed
  by its `書 章:節` label (e.g. `詩篇 100:1-2`), laid out in order.
  Add a translation such as `（和合本）`,
  `（新譯本）` or `（現代中文譯本2019版）`; the default is 和合本 (unv).
- **Reference font / Body font**: in pt. The body font has a minimum of 44pt.
- **Content**: paste the whole passage into the single box; it is **split across
  slides automatically** (see section 4).

### 3. Scripture Reading
- Same as the Call to Worship, but usually a longer passage.
- Also just paste one continuous block of text — no manual paging needed.

### 2. Hymns
- Global settings for this section: **lyric font / minimum font / left-right
  margin**.
- Each hymn is a card with: title, English title, source, refrain, and verses.
- "Repeat refrain after every verse" inserts the refrain after each verse.
- "Add verse" adds a stanza; "Remove this hymn" deletes the whole hymn.
- Backgrounds come from `media/song_bg1.jpg` and so on; they can be replaced on
  the build page.

### 4. Sermon
- This section **only accepts a ready-made `.pptx`**: upload it on the
  "Generate slides" page. Nothing to type on the edit page.

### 5. Announcements
- Same as above: **ready-made `.pptx` only**, uploaded on the build page.

### 6. Communion
- When checked, a "Communion + Apostles' Creed" pair of slides is added after
  the sermon.
- It is **pre-checked on the first Sunday of the month**.

---

## 3. Generate slides (tab 2)

1. **Upload files to merge (optional)**: three slots for Hymns, Sermon, and
   Announcements. A section with an uploaded file is used **as the file is**
   (source formatting preserved); an empty slot is compiled live from the edit
   page. You can mix and match.
2. **Enter a filename (local runs)**: expand "or enter a filename" and the app
   looks in the app folder, `data/downloads`, and `~/Downloads`. If not found it
   falls back automatically and warns you.
3. **Replace background images (optional)**: swap the title background image of
   each section.
4. **Announcement images (optional)**: attach images if an announcement is
   image-based.
5. **Reuse previously saved/generated section slides**: on by default. When on,
   a section generated earlier for the same date is reused; turn it off to
   recompile everything from the edit page.
6. **Section source status**: verify where each section comes from before
   building:
   - Uploaded / file (file wins)
   - Saved (reuse of a previously generated file)
   - Compiled (from the edit page)
7. Back in the side panel, click "Generate full service", then
   "Download PPTX" to get `Sunday_Service_<date>.pptx`.

> The "Generate individual pptx" block produces Hymns / Sermon / Announcements
> separately if you want to send one section to a teammate first; those files
> are also saved and reused when building the full service.

---

## 4. How scripture auto-splitting works

- The passage you paste is treated as one continuous string. The app estimates
  the characters per line from the **font size**, derives lines per slide from
  the text box height, and splits into multiple slides.
- **Font size starts at 44pt**: default 44pt, adjustable upward but never below
  44pt.
- Splits avoid punctuation: a slide (and each rendered line) **never ends on a
  punctuation mark**, and the next slide never starts on one.
- Chinese has no spaces between words, so breaks fall mid-sentence. If you would
  rather break *after* sentence punctuation (the opposite rule), ask for the
  change.

---

## 5. Storage and data

- **Week content** is written to Google Sheets (the `weeks` sheet). On Streamlit
  Community Cloud the server filesystem is ephemeral, so content is not kept on
  the server.
- With no Sheets configuration, a local run falls back to `data/weeks.json`.
- Uploaded images and individually generated section files live in
  `data/decks/` on the server; on the Cloud they are cleared on restart, while
  locally they persist.

---

## 6. FAQ

**Q: PowerPoint says a merged deck "needs repair"?**
A: The merge now fixes master/layout ID collisions, so this should no longer
happen. If it still does, check that the source `.pptx` opens cleanly in
PowerPoint on its own.

**Q: Do I still need to fill the edit page after uploading a file?**
A: No. Once a section has an uploaded file, that file is used.

**Q: I changed the edit page but the output did not change?**
A: That section may have an uploaded file, or "Reuse previously saved" is on and
an old file was reused. Turn off that toggle or re-upload, then build again.

**Q: The download button is gone?**
A: Click "Generate full service" again; the download button only shows the most
recent build for the current date.
