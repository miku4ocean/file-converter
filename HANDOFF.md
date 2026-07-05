# HANDOFF — file-converter
更新：2026-07-05／claude

## 目前目標
純前端零安裝的多格式檔案轉換工具，目標部署至 GitHub Pages。

## 狀態
- 已完成：`index.html` 主介面 + workers/ 目錄下各格式轉換器；大量測試 HTML 頁面
- 進行中：工作區乾淨（所有檔案直接在根目錄，無 npm 管理）
- 驗收現況：未驗證（未手動測試各轉換功能）

## 下一步（接手的人從這裡開始）
1. 用瀏覽器開 `index.html` 確認基本功能可用（不需 npm install）
2. 清理根目錄大量測試/除錯 HTML 檔（50+ 個），只保留 `index.html` 與必要 assets
3. 依 `GITHUB_PAGES_SETUP.md` 指引設定 GitHub Pages 部署

## 地雷（別踩）
- 根目錄有大量同質測試 HTML（`debug-*.html`、`test-*.html`），是開發過程產物，勿視為正式功能
- PDF 轉換走 Canvas/jsPDF 方案，中文字型支援需額外字型檔，目前狀態不明
- `workers/` 目錄的 Web Worker 為非同步，除錯時需注意跨執行緒訊息

## 主辦權
單線／待分派
