# 崇拜投影片 · Streamlit 網頁版（純 Python）

Web editor + pptx generator for the weekly Sunday service deck. No PWA, no
Google Apps Script, no local build server. Deploy to Streamlit Community Cloud
and users log in on the output URL.

```
streamlit-app/
├── app.py              # Streamlit 編輯器 + 生成介面
├── deck_builder.py     # 呼叫 generate_deck.build() 產生 pptx（記憶體中）
├── generate_deck.py    # 既有生成器（原封不動複製）
├── store.py            # Google Sheets 儲存（gspread）/ 本機 JSON 備份
├── model.py            # 週次的資料模型、預設值、正規化
├── media/              # 節目標題背景圖、詩歌背景圖
├── blank_16x9.pptx     # 模板
├── sample_week.json    # 新增週次的預設內容
└── requirements.txt
```

## 本機執行

```
cd streamlit-app
make install     # 建立 .venv 並安裝相依套件（用預編譯 wheel）
make run         # 啟動 app（http://localhost:8501）
make verify      # 檢查 Google 試算表連線
make stop        # 停止 app
```

沒有配置 Google 試算表時會自動改用 `data/weeks.json`（本機測試用）。

手動安裝（不使用 make）：

```
python3 -m venv .venv && source .venv/bin/activate
pip install --only-binary=:all: -r requirements.txt   # macOS 需要預編譯 wheel
streamlit run app.py
```

### Windows

用 PowerShell 版安裝指令（等同 Makefile）：

```
.\setup.ps1 install   # 建立 .venv 並安裝相依套件
.\setup.ps1 run       # 啟動 app（http://localhost:8501）
.\setup.ps1 verify    # 檢查 Google 試算表連線
.\setup.ps1 stop      # 停止 app
.\setup.ps1 clean     # 移除 .venv
```

或直接雙擊 `install.bat`（安裝）、`run.bat`（啟動）。若 PowerShell 阻擋腳本，
先執行一次 `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`，bat 檔已自帶
`-ExecutionPolicy Bypass`。

## 為什麼內容必須存在 Google 試算表

Streamlit Community Cloud 的檔案系統是「暫時性」的——每次重啟／重新部署都會
清空。`data/weeks.json` 這類本機檔在雲端會被還原成上傳前的狀態，所以週次內容
一律寫入 Google 試算表（用 gspread，**不使用 Apps Script**）。生成出來的
`.pptx` 直接在瀏覽器下載，不佔伺服器磁碟。

## 部署到 Streamlit Community Cloud

1. 把整個 `streamlit-app/` 資料夾推到一個 GitHub repository（資料夾內的
   `data/` 保留──內容並非從 repo 讀取，但 `data/.gitkeep` 可維持資料夾）。
2. 在 Google Cloud Console 建立 Service Account，下載 JSON 金鑰；
   把金鑰欄位填入 `.streamlit/secrets.toml`（範例見
   `.streamlit/secrets.toml.example`）。
3. 建立（或沿用）一個 Google 試算表，把它**分享給 service account 的 email**
   （權限 Editor）。試算表內會自動建立名為 `weeks` 的分頁
   （欄位：date | json | updatedAt）。
4. `streamlit.community.cloud` → New app → 連接該 repo，選 `app.py`。
5. Settings → Secrets：貼上 `.streamlit/secrets.toml` 的內容。
6. Deploy。App 的帳號／權限由 Community Cloud 控制（app 可設為 private，
   加上 `allowed_users` 或直接鎖給自己）。

## 使用流程

- **左側面板**：切換／新增主日日期、`💾 儲存目前內容`、`📽 製成整場投影片
  （合併＋編譯）`，以及製成後的下載按鈕。
- **編輯內容**：選日期 → 嚮導分節：宣召／詩歌／**獻詩**／讀經／信息／
  **詩歌回應**／聖餐／家事分享。宣召與讀經可只填**出處**（如 `詩篇 34:1-3`、
  `羅馬書 12:1-8（和合本）`），按 **`從 fhl.net 聖經網輸入`** 由信望愛聖經網
  （bible.fhl.net）自動抓取經文，再自動分頁。
- **製成投影片**：
  1. 四個可合併的獨立 pptx：**詩歌敬拜**、**講道信息**、**家事分享**、
     **獻詩**（可選）、**詩歌回應**（可選；未上載則以編輯內容編譯）。
  2. 每一節都可先「個別產生」成獨立 .pptx（在網頁編譯，例如只有歌詞時
     直接產出 `songs_slides_<date>.pptx`），傳給其他同工或之後合併。
     - 「個別產生」的檔案會**自動儲存在 `data/decks/<date>/`**，之後同一天
       按「製成整場投影片」時會**自動沿用**，不必重新上載（可用「自動使用
       先前儲存／已產生的各節投影片」開關停用）。
  3. 合併：把該節的 .pptx **上載**（欄位裡選檔案），該節以檔案為準；留空則用
     編輯頁內容現場編譯。可混用（例如詩歌用檔案、信息用網頁內容）。
     - 合併會把該節**原封不動**搬入（連同它的版面／母片／佈景主題／圖片），
       所以字型、項目符號等**沿用來源檔格式**，亦不會令檔案損壞（PowerPoint
       不需修復）。
     - 本機自行執行時，也可**輸入檔名**（如 `sermon_2026.09.20.pptx`），
       系統會從 app 資料夾、`data/downloads`、`~/Downloads` 找檔案合併；
       找不到會自動改用已儲存檔案或網頁內容並提示。
  4. 製成前會顯示**各節來源狀態**（✅ 上載／檔案、💾 已儲存、🧩 現場編譯）
     以及是否加入聖餐＋使徒信經。**當月第一主日**會預設勾選聖餐＋使徒信經。
  5. 「製成整場投影片」合併固定頁＋四大節（詩歌／獻詩／講道／家事分享，依序：
     宣召 → 詩歌 → 禱告主禱文 → **獻詩** → 讀經 → 講道 → **詩歌回應** →
     聖餐 → 家事分享）→ 下載 `Sunday_Service_<date>.pptx`。
  6. 另可替換節目的背景圖片，或在製成時為「圖片型家事分享」附加圖片。

已知取捨：上傳的圖片（例如「圖片型家事分享」）不回存試算表，重啟後需重新
上載；內容文字與設定則全部持久化。個別產生的節檔案存於伺服器
`data/decks/`，在 Streamlit Community Cloud 上會隨重啟清除（本機則保留）。