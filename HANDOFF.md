# HANDOFF — file-converter
更新：2026-07-20／claude

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

## 下一步（接手的人從這裡開始）
1. 用瀏覽器開 `index.html` 確認基本轉換功能可用（不需 npm install）
2. ✅ 已完成（2026-07-27）：更新 `GITHUB_PAGES_SETUP.md`／`TEST_GUIDE.md` 中指向舊測試頁路徑的連結（已移至 `tests/`）——8 個連結已更正
3. 依 `GITHUB_PAGES_SETUP.md` 指引設定 GitHub Pages 部署

## 地雷（別踩）
- `tests/` 目錄下皆為開發過程產物，勿視為正式功能；正式入口只有根目錄 `index.html`
- PDF 轉換走 Canvas/jsPDF 方案，中文字型支援需額外字型檔，目前狀態不明
- `workers/` 目錄的 Web Worker 為非同步，除錯時需注意跨執行緒訊息

## 主辦權
單線／待分派
