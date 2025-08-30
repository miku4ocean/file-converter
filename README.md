# 檔案格式轉換器 (File Converter)

🚀 純前端網頁版檔案格式轉換工具，支援圖片、文書、音檔、影片格式互轉

## ✨ 特色功能

- **🖼️ 圖片轉換** - 支援 JPG、PNG、WebP 互轉
- **📄 文書處理** - PDF、Word、Excel 格式轉換 (規劃中)
- **🎵 音檔轉換** - MP3、WAV、OGG 格式轉換 (規劃中)
- **🎬 影片轉換** - MP4、WebM 格式轉換 (規劃中)
- **🔒 隱私保護** - 純前端處理，檔案不上傳伺服器
- **📱 響應式設計** - 支援桌面和行動裝置

## 🎯 當前版本功能

### 圖片轉換 (v1.0)
- ✅ 拖拉式檔案上傳
- ✅ 支援格式：JPG ↔ PNG ↔ WebP
- ✅ 品質調整：10%-100% 壓縮設定
- ✅ 尺寸調整：自訂最大寬度
- ✅ 批次處理：同時轉換多個檔案
- ✅ 即時預覽：上傳前後圖片預覽
- ✅ 進度顯示：即時轉換進度條

## 🚀 快速開始

### 在線使用
直接開啟 `index.html` 檔案即可使用

### 本地伺服器
```bash
# 使用 Python 啟動本地伺服器
python3 -m http.server 8080

# 或使用 Node.js
npx http-server -p 8080
```

然後開啟瀏覽器前往 `http://localhost:8080`

## 📁 專案結構

```
file-converter/
├── index.html                   # 主頁面
├── assets/
│   ├── css/
│   │   └── style.css           # 響應式樣式表
│   └── js/
│       ├── app.js              # 主要應用程式邏輯
│       └── converters/
│           └── image.js        # 圖片轉換核心功能
├── workers/                     # Web Workers (未來擴展)
├── README.md                   # 專案說明文件
└── .gitignore                  # Git 忽略檔案
```

## 🔧 技術架構

### 前端技術
- **HTML5** - 現代網頁結構
- **CSS3** - Flexbox/Grid 響應式布局
- **Vanilla JavaScript** - 純 JS，無框架依賴
- **Canvas API** - 圖片處理核心
- **File API** - 檔案讀取與處理
- **Drag & Drop API** - 拖拉上傳功能

### 圖片處理
- **Canvas 2D Context** - 圖片繪製與轉換
- **ImageSmoothing** - 高品質圖片縮放
- **Blob API** - 檔案格式轉換
- **URL.createObjectURL** - 預覽功能

## 📱 瀏覽器相容性

| 瀏覽器 | 版本支援 | 備註 |
|--------|----------|------|
| Chrome | 60+ | ✅ 完整支援 |
| Firefox | 55+ | ✅ 完整支援 |
| Safari | 12+ | ✅ 完整支援 |
| Edge | 79+ | ✅ 完整支援 |

## 🗺️ 開發路線圖

### Phase 1 - 圖片轉換 ✅
- [x] 基礎圖片格式轉換
- [x] 品質與尺寸調整
- [x] 拖拉上傳介面
- [x] 批次處理功能

### Phase 2 - 文書處理 🚧
- [ ] PDF 讀取與轉換
- [ ] Word 文件處理 (使用 mammoth.js)
- [ ] Excel 檔案處理 (使用 SheetJS)
- [ ] PowerPoint 支援

### Phase 3 - 音檔轉換 📋
- [ ] Web Audio API 整合
- [ ] MP3/WAV/OGG 互轉
- [ ] 音檔品質調整
- [ ] 音檔視覺化 (WaveSurfer.js)

### Phase 4 - 影片轉換 📋
- [ ] FFmpeg.wasm 整合
- [ ] 基礎影片格式轉換
- [ ] 影片壓縮與品質調整
- [ ] 影片截圖功能

## 🤝 貢獻指南

歡迎提交 Issue 和 Pull Request！

### 開發環境設置
```bash
# 克隆專案
git clone https://github.com/miku4ocean/file-converter.git
cd file-converter

# 啟動開發伺服器
python3 -m http.server 8080
```

### 提交規範
- `feat:` 新功能
- `fix:` 錯誤修復
- `docs:` 文檔更新
- `style:` 代碼格式調整
- `refactor:` 代碼重構
- `test:` 測試相關

## 📄 授權條款

MIT License - 詳見 [LICENSE](LICENSE) 檔案

## 🙏 致謝

- [Canvas API](https://developer.mozilla.org/en-US/docs/Web/API/Canvas_API)
- [File API](https://developer.mozilla.org/en-US/docs/Web/API/File)
- [Drag and Drop API](https://developer.mozilla.org/en-US/docs/Web/API/HTML_Drag_and_Drop_API)

---

⭐ 如果這個專案對你有幫助，請給它一個星星！