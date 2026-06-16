# DF REG MKT Teammate — 技術說明 / 交接文件

本資料夾是網站的「產生器原始碼」備份。網站本身（`index.html`）在 repo 根目錄，已上線於：
**https://garena-df-regional.github.io/marketing-teammate/**

給 AI（Claude / ChatGPT / Gemini）查詢時，建議直接貼純文字版，命中率最高：
**https://garena-df-regional.github.io/marketing-teammate/content.md**

---

## 整體架構（東西住在哪）

| 元件 | 位置 | 說明 |
|------|------|------|
| 內容資料 | **Google Sheet** | 唯一的內容來源，所有 tab／流程／連結都在這 |
| 更新程式 | **Google Sheet 的 Apps Script** | `Code.gs`，提供「DF MKT Site → 🔄 Update Website」選單 |
| 網站本體（人看） | **GitHub repo 根目錄** `index.html` | 靜態網頁，托管於 GitHub Pages |
| 純文字版（AI 看） | **repo 根目錄** `content.md` | 同樣內容平鋪成 Markdown，AI 讀取最準 |
| 防搜尋 | `index.html` 內的 `<meta name="robots" content="noindex">` | 不被 Google 收進搜尋結果，但 AI／有網址者可讀 |
| 允許讀取 | `robots.txt`（`Allow: /`） | 讓 AI 能 fetch 網址 |
| 版型原始碼 | **本資料夾** `src/` | 改網站長相時才會用到 |

## 日常更新內容（不需碰程式）

1. 在 Google Sheet 修改／新增內容
2. 點選單 **DF MKT Site → 🔄 Update Website**
3. 約 1 分鐘後網站自動更新

> 任何有 Sheet 編輯權限的人都能按。Token 已設定在 Apps Script 專案（共用），
> 每人第一次使用需各自通過一次 Google 授權。

---

## 新增頁簽（不需碰程式）

左側選單會**自動掃描 Google Sheet 的所有分頁**，分頁順序 = 網站左側順序。
新增分頁後按一次「🔄 Update Website」就會出現，無需改程式或碰 GitHub。

頁簽類型由分頁名稱自動判斷：

| 你要的 | 分頁怎麼命名 | 分頁內容格式 | 網站呈現 |
|--------|------------|------------|---------|
| **內容頁簽**（情境細節） | `TabN 名稱`（如 `Tab14 Onboarding`） | 第 1 列標題、第 2 列欄位名（Scenario / Regional PIC / Process Steps / What to Prepare / Timeline / Common Q&A / Related Links）、第 3 列起資料 | 卡片式（可展開） |
| **通用頁簽**（連結／清單／架構表） | 任意名稱（如 `Important Folder/Sheet/Doc`） | 第 1 列當標題列，第 2 列起資料；任何欄位是連結都會自動可點 | 表格式 |

規則細節：
- **順序**：左側順序完全跟著 Google Sheet 分頁的左右順序——分頁拖到哪，網站就排到哪。
- **分隔線**：第一個 `TabN` 卡片頁之前會自動加一條分隔線（通用頁簽排在分隔線上方）。
- **顯示名**：自動從分頁名產生（`Tab1 Budget Plan & Report` → `Tab 1 · Budget Plan & Report`；通用頁簽直接用分頁全名）。
- 特殊分頁 `Index` / `Reg & Local Contactor` / `Social Media Link` 有專屬版面，**名稱請勿更動**。
- `content.md`（給 AI 的純文字版）會自動帶入新分頁。

> 技術實作：`generate_site.py` / `Code.gs` 的 `classify()` 依分頁名分類（特殊／`Tab\d+`→card／其他→generic），
> `buildData()` 按分頁順序動態產生 NAV，所以無寫死的頁簽清單。

---

## 改網站版型（需要前端知識）

版型的 source of truth 是 `src/generate_site.py` 裡的 HTML/CSS 模板。
**不要直接改 GitHub 上的 `index.html`**——下次有人按 Update Website 會被覆蓋。

流程：

```bash
# 1. 改 generate_site.py 裡的 CSS/HTML 模板
#    （注意：generate_site.py 頂部的 EXCEL 路徑要指向你本機的 Excel）

# 2. 重新產生並預覽
python generate_site.py        # 產出 index.html，瀏覽器開啟確認

# 3. 把新模板打包進 Apps Script
python build_apps_script.py    # 產出 apps-script/Code.gs

# 4. 把新的 Code.gs 全部內容，重貼進 Google Sheet 的 Apps Script，存檔
#    （這樣「Update Website」才會用新版型）
```

`build_apps_script.py` 會把 `index.html` 的版型模板以 base64 嵌入 `Code.gs`，
並把資料部分換成 `__DATA_JSON__` 佔位符，Apps Script 執行時再填入即時的 Sheet 資料。

---

## 檔案說明

| 檔案 | 用途 |
|------|------|
| `generate_site.py` | 從 Excel 產生 `index.html` + `content.md`（版型 source of truth；會把 `assets/logo.png` 以 base64 內嵌） |
| `build_apps_script.py` | 把 `index.html` 模板打包成 `Code.gs` |
| `Code.gs` | Apps Script 程式（已貼進 Google Sheet；此為備份） |
| `assets/logo.png` | 左上角 sidebar logo（已去白底透明化）；換 logo 換這張再重新產生即可 |

## Token 與帳號歸屬

- repo 已位於 GitHub Organization **`garena-df-regional`**（不綁任何個人帳號）。
- GitHub Token 存在 Apps Script 專案的 Script Properties（不在程式碼、不在 repo）。
- Token 的 Resource owner 應選 **`garena-df-regional`**（org），權限 Contents: Read and write。

### 待完成的交接動作
- [ ] 加接班人（Ruru）當 org `garena-df-regional` 的 owner（待其 GitHub 帳號）
- [ ] Token 到期前重新產生並用「🔑 Set GitHub Token」更新（建議到期日設長或無期限）
- [ ] 通知各區最終網址：`https://garena-df-regional.github.io/marketing-teammate/`

## 為什麼是公開 GitHub Pages，不是公司 intranet

核心目的是讓本地團隊把網址直接貼給 Claude / ChatGPT 查詢。intranet 需登入／IP 鎖，
AI 讀不到，故採公開網站 +「允許讀取（robots.txt）+ 不被搜尋（noindex meta）」。
內容非機密；Related Links 的 Google Sheet／Drive 連結仍需 Garena 權限才能開啟。
