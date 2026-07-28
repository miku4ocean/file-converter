# progress.md — file-converter

> 本檔內容依實際讀取 `HANDOFF.md`、`AGENTS.md`、`README.md`、`CLAUDE.md`、`index.html`、`assets/js/*`、`assets/js/converters/*`、`CONVERSION_FIX_SUMMARY.md`、`GITHUB_PAGES_SETUP.md`、git 紀錄產出。查不到佐證處一律標註「未確認」，不臆造。

## A. 專案名稱
File Converter（檔案格式轉換器）

## B. 專案路徑
`/Users/leonalin/Code/file-converter`

## C. 專案簡介
純前端、零安裝的多格式檔案轉換工具。單一 `index.html` 進入點，透過 Vanilla JavaScript 狀態機在「圖片／文書／表單／簡報」四種轉檔模式間切換，所有轉換運算都在使用者瀏覽器本機完成（Canvas API、File API、Blob API），不上傳任何伺服器。目標部署管道為 GitHub Pages，repo 為 `https://github.com/miku4ocean/file-converter`（已設定 remote origin，`main` 分支）。

## D. 專案開發目的
依 README.md／HANDOFF.md：提供一個「不用裝軟體、不用開帳號、檔案不出瀏覽器」的線上格式轉換工具，讓使用者可即時完成圖片、文書、表單、簡報等常見格式互轉，避免將檔案內容上傳給不受信任的第三方線上轉檔服務。

## E. 解決使用者痛點
- 一般線上轉檔服務需要上傳檔案到陌生伺服器，可能涉及隱私或機密外洩疑慮；本工具承諾「純前端處理，檔案不上傳伺服器」。
- 多數轉檔工具需要安裝軟體或註冊帳號；本工具開啟 `index.html`（或以靜態伺服器打開）即可使用，零安裝、零帳號。
- 常見格式互轉（圖片、文書、試算表、簡報）分散在不同工具，本工具嘗試集中在單一介面完成。

## F. 專案功能細項介紹
- **圖片轉檔**（`image.js`）：JPG／PNG／GIF／BMP／WebP 互轉，支援拖拉上傳、批次處理、品質調整（10%–100%）、最大寬度限制、上傳前後預覽、即時進度條。README 標記此段為 v1.0 已完成功能。
- **文書轉檔**（`document.js`，1324 行）：TXT／DOCX／HTML／Markdown／PDF／RTF 互轉，含「內容壓縮」選項（標準／輕度／中度／高度）。依 `CONVERSION_FIX_SUMMARY.md`，DOCX 是用 JSZip 手刻真正的 Office Open XML 容器（`[Content_Types].xml`、`word/document.xml` 等），PDF 用 jsPDF 產生真正的 PDF 結構並支援中文字元，並非單純改副檔名。
- **表單轉檔**（`spreadsheet.js`，852 行）：CSV／XLSX／JSON／HTML 表格／純文字互轉，含「包含標題行」選項。XLSX 優先用 SheetJS 產生，並有自製 XLSX ZIP 結構作為回退方案。
- **簡報轉檔**（`presentation.js`，979 行）：PPTX／HTML／純文字／Markdown／PDF 互轉，含「包含備註」選項。
- **PDF 輔助模組**：`visual-pdf-converter.js`（635 行）、`direct-pdf.js`（969 行），兩者皆被 `index.html` 實際載入，用於 PDF 相關的視覺化／直接轉換路徑，細節「未確認」（未逐行深讀，僅確認為正式流程一部分）。
- **函式庫動態載入機制**（`lib-loader.js` ＋ `auto-loader.js`）：啟動時自動以 CDN `<script>` 動態載入必要函式庫（jszip／jspdf／sheetjs），並在背景載入可選函式庫（mammoth／pptxgenjs／html2canvas／pdfjs-dist），失敗不阻斷核心功能。

## G. 專案規格及 RPD
**技術棧**
- HTML5 ＋ CSS3（Flexbox／Grid，響應式）＋ Vanilla JavaScript（無框架、無建置工具、無 npm 相依）
- 瀏覽器原生 API：Canvas 2D Context、File API、Blob API、Drag & Drop API、`URL.createObjectURL`
- 外部函式庫（執行期 CDN 動態載入，非隨包）：jszip、jspdf、sheetjs (XLSX)、mammoth、pptxgenjs、html2canvas、pdfjs-dist，來源 `cdn.jsdelivr.net` ／ `cdn.sheetjs.com`

**埠／指令**
- 無建置指令。可直接雙擊開啟 `index.html`，或啟動本機靜態伺服器：`python3 -m http.server 8080`、`npx http-server -p 8080`、`npx serve .`
- AGENTS.md 明文禁區：「不得引入需後端或 Node.js 的依賴；保持零安裝、可離線執行特性」

**資料流**
1. 使用者於瀏覽器選擇檔案類型分頁 → 拖拉或選取檔案
2. `app.js` 的 `FileConverter` 狀態機驗證檔案 → 顯示檔案清單 → 產生對應轉換設定表單
3. 使用者按下「開始轉換」→ 呼叫對應 `converters/*.js` 模組，於瀏覽器記憶體內用 Canvas／File／Blob API（必要時呼叫已載入的 CDN 函式庫）完成格式轉換
4. 轉換結果透過 `URL.createObjectURL` 提供本機下載，全程無網路請求傳送檔案內容；唯一的網路請求是「載入轉換工具本身」的 CDN 函式庫程式碼（非使用者資料）

**RPD（需求／產品定義，依 README／HANDOFF 彙整）**
- 目標平台：GitHub Pages 靜態託管
- 支援瀏覽器：Chrome 60+／Firefox 55+／Safari 12+／Edge 79+（README 表列，未實測驗證，狀態「未確認」）
- 分階段路線圖：Phase 1 圖片轉換（README 標記已完成）／Phase 2 文書處理／Phase 3 音檔轉換／Phase 4 影片轉換

## H. 目前已完成項目
- `index.html` 主介面 ＋ 四種轉檔分頁（圖片／文書／表單／簡報）的完整 UI 骨架與狀態流（上傳→清單→設定→進度→下載）
- 圖片轉檔功能：README 標記 v1.0 完成，含拖拉上傳、批次、品質／尺寸調整、預覽、進度條
- 文書／表單／簡報三個轉換器模組（`document.js`／`spreadsheet.js`／`presentation.js`）已實作「真正產生對應格式檔案」的邏輯（非僅改副檔名），依 `CONVERSION_FIX_SUMMARY.md` 記載為针對先前「轉出檔案無法開啟」問題的修復成果
- CDN 函式庫動態載入機制（`lib-loader.js`／`auto-loader.js`）：必要庫優先載入、可選庫背景載入、載入失敗有 fallback 邏輯
- 2026-07-20：根目錄 46 個測試／除錯 HTML（`test-*`、`debug-*`、`final-*`、`*-fix*` 等）已用 `git mv` 整批搬到 `tests/` 目錄，並交叉確認 `index.html` 未連結任何被搬移的頁面
- 全庫資安掃描（api-key／secret／password／token 樣式）：唯一命中在 `tests/stress-test-pdf.html` 內的 SQL-injection 測試假資料，無真實金鑰外洩
- `index.html` 及其引用的 `assets/` 路徑已用本機 static server 驗證皆可載入（HTTP 200）

## I. 尚待完成項目
- **手動功能驗證未做**：HANDOFF.md 明載「未手動測試各轉換功能實際效果」，即文書／表單／簡報轉檔的實際輸出正確性尚待人工開瀏覽器逐一驗證
- **孤兒轉換器檔案的去留未決**：`cloudconvert-pdf.js`（內含真實第三方雲端上傳 API `api.cloudconvert.com`，需要 API Key）、`fixed-pdf-converter.js`、`presentation-backup.js` 三檔存在於 `assets/js/converters/`，但未被 `index.html` 載入，HANDOFF.md 註記「留給下一手判斷是否清理」
- **文件連結已過期未更新**：`GITHUB_PAGES_SETUP.md`（列出 `direct-pdf-test.html` 等舊路徑）、`TEST_GUIDE.md`、`CONVERSION_FIX_SUMMARY.md` 等仍引用搬移前的測試頁路徑，搬移到 `tests/` 後這些連結已失效，HANDOFF.md 明列為下一步待辦
- **GitHub Pages 是否已實際啟用「未確認」**：`GITHUB_PAGES_SETUP.md` 提供設定步驟，但目前是否已在 GitHub repo 設定中啟用、線上網址是否可訪問，本次未做網路查核，狀態「未確認」
- **音檔／影片轉換（README Phase 3／4）尚未實作**：`assets/js/converters/` 下未見對應 `audio.js`／`video.js` 或 FFmpeg.wasm 整合，索引頁的 4 個分頁也不含音檔／影片選項，屬於「規劃中、尚未實作」
- **`workers/` 目錄實際為空**：HANDOFF.md「地雷」段落提及「`workers/` 目錄的 Web Worker 為非同步，除錯時需注意跨執行緒訊息」，但實地檢查該目錄為 0 檔案的空資料夾，查無對應程式碼，現況與該筆記錄有落差，狀態「未確認」（不確定是筆記過時還是功能尚未落地）
- **「可離線執行」目標與實作有落差**：AGENTS.md 明訂本專案禁區為「保持零安裝、可離線執行特性」，但核心函式庫 jszip／jspdf／sheetjs 現況透過 `lib-loader.js` 於執行期向 CDN（`cdn.jsdelivr.net`／`cdn.sheetjs.com`）動態下載，斷網環境下文書／表單／簡報轉換能否運作「未確認」；僅圖片轉檔（純 Canvas API）不依賴任何外部網路

## J. 系統優化或增加功能建議
- 若要落實「可離線執行」目標，建議將必要函式庫（jszip／jspdf／sheetjs）改為 vendor 進 repo 本地載入，或加入 Service Worker 做 cache-first 快取策略，讓首次載入後可離線使用
- 建議明確處理三個孤兒轉換器檔案：若確定不用則移除，若要保留 CloudConvert 路徑則需明確標示「此路徑會將檔案上傳至第三方」，避免與 footer 承諾的「純前端、不上傳」產生認知落差或誤用風險
- 補上自動化測試：目前 `tests/` 下 46 個檔案皆為手動 debug 用 HTML，沒有可自動執行的單元測試或 CI 流程；建議至少為三個較複雜的轉換模組（`document.js`／`spreadsheet.js`／`presentation.js`，皆逾 800 行）補上基本輸出格式驗證測試
- 大檔案／批次轉換建議搬進 Web Worker（目前 `workers/` 為空目錄），避免長時間轉換卡住主執行緒導致 UI 凍結
- 校對並更新 `README.md` 開發路線圖的勾選狀態：其 Phase 2（文書處理）目前仍標示 🚧 未勾選，但 `document.js` 等模組實際上已有相當完整的實作（依 `CONVERSION_FIX_SUMMARY.md`），文件與程式碼現況不同步，容易誤導後續維護者
- 同步更新 `GITHUB_PAGES_SETUP.md`／`TEST_GUIDE.md` 中已因搬移而失效的測試頁連結
- 若持續發展音檔／影片轉換（README Phase 3／4），因涉及 FFmpeg.wasm 等較大型 WASM 依賴，建議及早評估其與「零安裝、可離線」目標之間的檔案體積與載入時間取捨
