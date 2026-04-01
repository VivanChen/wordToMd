# Word → Markdown Converter & Reader

一個部署在 Netlify 的 Word 文件轉 Markdown 工具，支援 `.doc` 和 `.docx` 格式。

## 架構

| 元件 | 技術 | 說明 |
|------|------|------|
| 前端 | React + Vite | SPA 單頁應用 |
| .docx 轉換 | mammoth.js | 瀏覽器端直接處理 |
| .doc 轉換 | Netlify Function + word-extractor | Server-side 處理舊格式 |
| MD 儲存 | GitHub API | Push 到指定 repo |
| 部署 | Netlify | 靜態 SPA + Serverless Functions |

## 功能

- **轉換頁面**：拖曳上傳 Word 檔案，支援批次轉換
- **MD 閱讀器**：開啟本機或 GitHub 上的 .md 檔案，即時渲染預覽
- **GitHub 整合**：瀏覽 repo 中的 .md 檔案，一鍵推送轉換結果
- **Markdown 編輯**：轉換後可直接編輯 Markdown 原始碼
- **三種檢視模式**：Preview / Markdown Source / Raw HTML

## 快速開始

```bash
# 1. Clone repo
git clone https://github.com/YOUR_USER/word-to-md.git
cd word-to-md

# 2. 安裝依賴
npm install

# 3. 本地開發（含 Netlify Functions）
npx netlify dev

# 4. 建置
npm run build
```

## 部署到 Netlify

1. 將專案推送到 GitHub
2. 前往 [Netlify](https://app.netlify.com) → New site from Git
3. 選擇你的 GitHub repo
4. 設定：
   - **Build command**: `npm run build`
   - **Publish directory**: `dist`
   - Functions directory 已在 `netlify.toml` 中設定
5. 部署完成！

## 專案結構

```
word-to-md/
├── index.html              # 入口 HTML
├── netlify.toml            # Netlify 設定（含 Functions 路由）
├── package.json
├── vite.config.js
├── src/
│   ├── main.jsx            # React 入口
│   └── App.jsx             # 主應用元件
└── netlify/
    └── functions/
        └── convert-doc.mjs # .doc 轉換 Serverless Function
```

## .doc vs .docx 轉換差異

| | .docx | .doc |
|--|-------|------|
| 處理位置 | 瀏覽器端 | Netlify Function (server) |
| 引擎 | mammoth.js | word-extractor |
| 格式保真度 | 高（保留粗體、斜體、標題、表格、列表、圖片等） | 中（保留文字結構，格式有限） |
| 檔案大小限制 | 無特別限制 | Netlify Function body limit ~6MB |

## GitHub 設定

在應用的「設定」頁面填入：

- **Personal Access Token**：需要 `repo` scope
- **Owner**：GitHub 用戶名或組織名
- **Repository**：目標 repo 名稱
- **Path**：存放 .md 檔案的資料夾（預設 `docs`）
