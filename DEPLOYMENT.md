# 英語正則 POC 雲端部署說明

這個版本可以用 Docker 部署到支援 persistent disk 的平台，例如 Render、Railway、Fly.io 或一般 VPS。

## 專案結構

```text
product-poc/
  frontend/   React + Vite 前端
  backend/    Python 標準庫 API server、上傳與分析歷史資料
```

本機開發時可分開啟動：

```bash
cd product-poc
npm run serve:backend
npm run dev:frontend
```

單機 POC 或部署前測試：

```bash
cd product-poc
npm run serve:poc
```

## 必要條件

- 需要 Docker build。
- 需要一個可掛載的永久磁碟，建議掛到 `/data`。
- 單機 POC 建議只開 1 個 instance，因為登入 session 目前存在記憶體。
- 大批次 ZIP 上傳時，平台或反向代理的 request body limit 需大於 ZIP 檔案大小。

## 環境變數

| 變數 | 建議值 | 說明 |
| --- | --- | --- |
| `POC_USERNAME` | `admin` 或自訂 | 後台登入帳號 |
| `POC_PASSWORD` | 強密碼 | 後台登入密碼，正式環境不要用預設值 |
| `POC_COOKIE_SECURE` | `1` | HTTPS 環境請設為 `1` |
| `POC_RUNS_DIR` | `/data/runs` | 上傳 ZIP、分析結果、DOCX/CSV/JSON 歷史資料 |
| `MANUAL_ANALYSIS_MODE` | `auto` | 有 `GEMINI_API_KEY` 時使用 AI 歸納，否則退回規則式 |
| `GEMINI_API_KEY` | 從 Google AI Studio 建立 | Gemini API key，正式環境請設 secret |
| `GEMINI_MODEL` | `gemini-2.5-flash-lite` | Gemini 歸納模型 |
| `PORT` | 平台自動提供 | 服務監聽 port |

## 本機 Docker 測試

```bash
docker build -t essay-regularity-poc .
docker run --rm -p 8787:8787 \
  -e POC_USERNAME=admin \
  -e POC_PASSWORD='change-me' \
  -e POC_COOKIE_SECURE=0 \
  -e POC_RUNS_DIR=/data/runs \
  -e MANUAL_ANALYSIS_MODE=auto \
  -e GEMINI_MODEL=gemini-2.5-flash-lite \
  -v "$PWD/.data:/data" \
  essay-regularity-poc
```

開啟：

```text
http://127.0.0.1:8787/
```

## Render 部署建議

1. 將專案推到 GitHub。
2. Render 建立 Web Service，選擇 Docker。
3. 設定 persistent disk：
   - Mount path: `/data`
   - Size: 依保存批次決定，POC 可先用 10GB。
4. 設定環境變數：
   - `POC_USERNAME`
   - `POC_PASSWORD`
   - `POC_COOKIE_SECURE=1`
   - `POC_RUNS_DIR=/data/runs`
   - `MANUAL_ANALYSIS_MODE=auto`
   - `GEMINI_MODEL=gemini-2.5-flash-lite`
   - `GEMINI_API_KEY`
5. Deploy 後用 Render 網址登入測試。

## Production 注意事項

- 目前是 POC 架構，適合單機部署；如果要多台水平擴充，需要把 session 與歷史資料移到 DB/Object Storage。
- 533 份 PDF 的批次會吃 CPU、磁碟 I/O 與記憶體，建議至少 1-2 vCPU、1GB RAM 以上；若 ZIP 很大，需確認平台支援該上傳大小。
- 上傳與分析目前在同一個 web process 內排隊執行。正式產品可再拆成 queue worker，避免長時間分析影響 web request。
- PDF 擷取不使用 AI；手冊歸納可使用 Gemini。若 Gemini API 失敗或未設定 key，系統會自動退回規則式歸納並保留可下載結果。
