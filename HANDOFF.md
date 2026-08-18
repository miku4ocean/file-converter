# HANDOFF — file-converter
更新：2026-08-18／claude（CSV 公式注入修復）

## 2026-08-18／claude（CSV 公式注入修復，跨專案同族修復之一）
`assets/js/converters/spreadsheet.js` 的 `convertToCsv()` 先前直接把儲存格值寫入 CSV，
未防範 OWASP CSV Injection（CSV 公式注入）：若欄位值以 `=`、`+`、`-`、`@`、Tab 或 `\r`
開頭，在 Excel/Google Sheets 開啟時可能被當公式執行（如 `=cmd|'/c calc'!A1`）。
新增 `SpreadsheetConverter.sanitizeCsvField()` 輔助函式，在既有的引號跳脫（`""`）之前，
先檢查trim後的值是否以上述字元開頭，是則補一個前導單引號 `'` 使其被當純文字讀取；
純既有跳脫邏輯不變。已於 `tests/unit/spreadsheet.test.js` 補上對應測試（含
`=cmd|'/c calc'!A1` payload）並跑 `node --test "tests/unit/*.test.js"` 全數通過（68/68）。
此修復為跨 5 個專案的同族批次整治（CSV 公式注入）之一。

更新：2026-08-07／claude（孤兒轉換器+測試檔清理+README同步+workers移除）

## 目前目標
純前端零安裝的多格式檔案轉換工具，目標部署至 GitHub Pages。

## 狀態
- 已完成：`index.html` 主介面 + assets/js/converters 各格式轉換器；根目錄測試 HTML 已清理
- 進行中：工作區乾淨（無 npm 管理，純靜態站）
- 驗收現況：index.html 及其引用的 assets 路徑已用本機 static server 驗證皆可載入（200）；未手動測試各轉換功能實際效果

## 本次更新（2026-07-20）
- 根目錄 46 個測試/除錯 HTML（`test-*`、`debug-*`、`final-*`、`*-fix*` 等，含 `github-ready-converter.html`／`onlyoffice-converter.html`／`libreoffice-wasm-solution.html` 等替代方案實驗頁）用 `git mv` 整批搬到新建的 `tests/` 目錄，未刪除任何內容
- 交叉確認 index.html 未連結任何被搬移的頁面，安全
- 保留原樣、未搬動：`assets/js/converters/` 下的 `cloudconvert-pdf.js`、`fixed-pdf-converter.js`、`presentation-backup.js`（未被 index.html 載入，疑似廢棄實驗，但屬 JS 非 HTML，超出本次任務範圍，留給下一手判斷是否清理）
- 未修改：`GITHUB_PAGES_SETUP.md`、`TEST_GUIDE.md`、`CONVERSION_FIX_SUMMARY.md`、`final-validation-test.js`、`quick-test.js`、`check-github-pages.js`、`validate-pdf-fix.js` 等——這些檔案內文引用了被搬移的測試頁路徑（如 `direct-pdf-test.html`），搬移後路徑已過期，但依任務範圍本次不動非 HTML 清理相關檔案，下一手可視需要更新
- 資安掃描：全庫掃過 api-key/secret/password/token 樣式，唯一命中在 `tests/stress-test-pdf.html` 內的 SQL-injection 測試字串（假資料，非真實金鑰），無實際外洩

## 2026-08-03／claude（PDF 中文字型支援修正）

地雷段原寫「PDF 中文字型支援狀態不明」，本輪評估並修正。

### 問題
6 支轉換器的 `font-family` 存在不一致的 CJK fallback：
- `direct-pdf.js`、`fixed-pdf-converter.js`、`visual-pdf-converter.js`、`presentation-backup.js`：
  部分宣告只有 `'微軟正黑體'`（Windows 繁體版限定）或完全沒有中文字型 fallback，
  在 macOS（無微軟正黑體）和 Linux 上會讓 html2canvas 截到方塊或豆腐字。
- `document.js` 的舊版路徑（L839/1270）也只有部分覆蓋。
- 同一份 `document.js` 的另外三處（L331/389/511）已經用了完整堆疊，證明不是刻意省略。

### 修法
統一成跨平台 CJK 堆疊：
- serif: `'Times New Roman', 'PingFang TC', 'Microsoft JhengHei', 'Noto Serif CJK TC', 'SimSun', serif`
- sans: `'PingFang TC', 'Microsoft JhengHei', 'Noto Sans CJK TC', 'Microsoft YaHei', Arial, sans-serif`
- mono: `'Courier New', 'PingFang TC', 'Microsoft JhengHei', monospace`

涵蓋：macOS 蘋方、Windows 微軟正黑體、Linux Noto CJK、Windows 簡體版微軟雅黑/宋體。

### 為什麼不走「下載 Noto Sans 字型嵌入 jsPDF」
`tests/chinese-pdf-test.html` 就是這個路線——從 Google Fonts 抓 Noto Sans SC woff2 再 base64 嵌入。
實測失敗（`字體載入失敗`），而且即使成功也有三個致命問題：
1. Noto Sans SC/TC 字型檔 ~5–8 MB，每次開頁面都要等
2. GitHub Pages 倉庫大小有限制
3. 離線不可用

html2canvas 路線（截 iframe 渲染結果為點陣圖）不需要嵌入字型——只要瀏覽器系統有中文字型，
截出來就是中文。macOS/Windows/iOS/Android **全部自帶中文系統字型**，
font-family 的作用只是確保系統在選字型時走到中文那一層，不會停在 Arial 然後吐方塊。

### 驗證
macOS Chrome 上用 canvas 直接繪製「繁體中文測試」六個字：
修正後的 CJK 堆疊與純 `Arial, sans-serif` 都能渲染（因為 macOS 的 sans-serif fallback
就是蘋方）。**本修正的價值在 Windows/Linux 上**——那些系統的 sans-serif fallback 不一定含
中文字型，沒有明確指定 `'Microsoft JhengHei'`/`'Noto Sans CJK TC'` 就會變成方塊。
無法在 macOS 上驗出差異，但程式碼的改動是確定正確的——每一處都從「漏掉某些平台」
變成「涵蓋全平台」，沒有行為變化只有覆蓋面擴大。

## 2026-08-07／claude（專案衛生清理）

### 孤兒轉換器搬出 assets/
`cloudconvert-pdf.js`（含 `api.cloudconvert.com` 端點，與「不上傳」承諾衝突）、
`fixed-pdf-converter.js`、`presentation-backup.js` 三檔未被 index.html 載入，
從 `assets/js/converters/` 搬入 `tests/`，避免隨 GitHub Pages 部署。

### 根目錄測試檔搬入 tests/
10 個測試/驗證檔案（`test-data.csv`、`quick-test.js`、`validate-conversions.js` 等）
從根目錄搬入 `tests/`，已確認 index.html 及生產 JS 無任何引用。

### README 同步
Phase 2 路線圖從 🚧 改為 ✅，專案結構樹補上全部六支轉換器。

### workers/ 移除
該目錄自建立以來始終為空，已刪除。

### CONVERSION_FIX_SUMMARY.md 路徑修正
測試檔路徑補上 `tests/` 前綴。

## 下一步（接手的人從這裡開始）
1. 用瀏覽器開 `index.html` 確認基本轉換功能可用（不需 npm install）
2. ✅ 已完成（2026-07-27）：更新 `GITHUB_PAGES_SETUP.md`／`TEST_GUIDE.md` 中指向舊測試頁路徑的連結（已移至 `tests/`）——8 個連結已更正
3. 依 `GITHUB_PAGES_SETUP.md` 指引設定 GitHub Pages 部署
4. **中文 PDF 驗證（需 Windows 或 Linux）**：在非 macOS 系統上做一次 TXT→PDF 轉換，
   確認中文字不是方塊。macOS 上測不出差異（系統 fallback 就有蘋方）。
5. 考慮將 CDN 函式庫（jszip/jspdf/sheetjs）vendor 進 repo 或加 Service Worker，
   以落實 AGENTS.md 標定的「可離線執行」目標

## 地雷（別踩）
- `tests/` 目錄下皆為開發過程產物（含三支搬入的廢棄轉換器），勿視為正式功能；正式入口只有根目錄 `index.html`
- ~~PDF 轉換走 Canvas/jsPDF 方案，中文字型支援需額外字型檔，目前狀態不明~~
  **2026-08-03 已修正**：不需要額外字型檔，html2canvas 路線只要系統有中文字型就能截到中文，
  各轉換器的 `font-family` 已統一成跨平台 CJK 堆疊（見上方條目）。
  `tests/chinese-pdf-test.html` 走的「下載字型嵌入 jsPDF」路線已確認不可行（見上方說明）。

## 主辦權
單線／待分派
