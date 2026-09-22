# price-monitor-2026

# 📱 iPhone 回收價格監控

從 Google Sheets（或本地 Excel）讀取 iPhone 回收價數據，自動生成互動式 HTML 報表，部署到 Vercel 供手機隨時查看。

---

## 專案目錄結構

```
price_monitor/
├── generate.py              # 主程式：讀數據 → 生成 HTML
├── requirements.txt         # Python 依賴套件
├── vercel.json              # Vercel 部署設定
├── README.md                # 本教學文件
├── .gitignore
├── credentials/
│   └── service_account.json # Google 憑證（自行放入，不上傳 Git）
├── data/
│   └── 價格表.xlsx           # 備用本地數據（首次執行自動建立範例）
├── public/
│   └── index.html           # 生成的互動式網頁（部署到 Vercel）
└── .github/
    └── workflows/
        └── deploy.yml       # GitHub Actions 自動化（選用）
```

---

## 第一步：本地環境設定

### 1.1 安裝 Python

確認已安裝 Python 3.10 以上：

```powershell
py --version
# 或 python --version
```

若未安裝，前往 https://www.python.org/downloads/ 下載安裝。

### 1.2 建立虛擬環境並安裝套件

在專案目錄下執行：

```powershell
cd c:\Users\ann_yue\price_monitor

# 建立虛擬環境
py -m venv .venv

# 啟動虛擬環境（Windows PowerShell）
.\.venv\Scripts\Activate.ps1

# 安裝依賴
pip install -r requirements.txt
```

### 1.3 首次測試（使用本地 Excel 備用方案）

不需要 Google 憑證也能先跑起來：

```powershell
# ⚠️ 啟用 .venv 後請用 python，不要用 py（py 會跑系統 Python，找不到 pandas）
python generate.py

# 或一鍵執行（自動使用 .venv）：
.\run.ps1
```

程式會自動建立 `data/價格表.xlsx` 範例資料，並生成 `public/index.html`。

用瀏覽器開啟 `public/index.html` 即可預覽互動式圖表。

---

## 第二步：設定 Google Sheets 數據源

### 2.1 建立 Google 試算表

1. 前往 [Google Sheets](https://sheets.google.com)，建立新試算表
2. 命名為 **「iPhone回收價格」**（或在 `generate.py` 修改 `SPREADSHEET_NAME`）
3. 第一列填入標題欄：

| Date | Time | Model | Capacity | Color | Price |
|------|------|-------|----------|-------|-------|
| 2025-09-22 | 10:00 | iPhone 16 Pro | 256GB | 原色鈦金屬 | 7200 |

> 也支援中文欄位名：日期、時段、型號、容量、顏色、回收價

### 2.2 Google Cloud Console 申請憑證

詳細步驟已寫在 `generate.py` 的 `load_from_google_sheets()` 函式註解中，摘要如下：

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 建立新專案（如 `price-monitor`）
3. 啟用 **Google Sheets API** 和 **Google Drive API**
4. 建立 **Service Account**（服務帳戶）
5. 下載 JSON 金鑰 → 重新命名為 `service_account.json`
6. 放到 `credentials/service_account.json`

### 2.3 分享試算表給 Service Account

1. 打開 `service_account.json`，找到 `"client_email"` 欄位
   ```
   "client_email": "price-monitor-bot@your-project.iam.gserviceaccount.com"
   ```
2. 在 Google Sheets 點「共用」→ 貼上上述 email → 權限選「編輯者」→ 傳送

### 2.4 再次執行測試

```powershell
py generate.py
```

若設定正確，終端機會顯示 `[INFO] 從 Google Sheets 讀取 N 筆資料`。

---

## 第三步：部署到 Vercel

### 3.1 推送到 GitHub

```powershell
# 初始化 Git（若尚未初始化）
git init
git add .
git commit -m "feat: iPhone price monitor with Plotly chart"

# 在 GitHub 建立新 repository（如 price-monitor），然後：
git remote add origin https://github.com/你的帳號/price-monitor.git
git branch -M main
git push -u origin main
```

> ⚠️ `credentials/service_account.json` 已在 `.gitignore` 中，不會被上傳。

### 3.2 連接 Vercel

1. 前往 [vercel.com](https://vercel.com)，用 GitHub 帳號登入
2. 點 **「Add New Project」**
3. 選擇你的 `price-monitor` repository
4. Vercel 會自動讀取 `vercel.json` 設定：
   - **Build Command**: `pip install -r requirements.txt && py generate.py`
   - **Output Directory**: `public`
5. （選用）在 **Environment Variables** 新增：
   - Key: `GOOGLE_SERVICE_ACCOUNT_JSON`
   - Value: 貼上整份 `service_account.json` 的內容
6. 點 **Deploy**

部署完成後，你會得到一個網址如：
```
https://price-monitor-xxxx.vercel.app
```

手機瀏覽器打開此網址即可查看價格走勢！

### 3.3 自動更新流程

每次你更新 Google Sheets 的數據後，有兩種方式刷新網頁：

| 方式 | 操作 |
|------|------|
| **手動觸發** | Vercel Dashboard → Deployments → Redeploy |
| **Push 程式碼** | 任意 git push 到 main，Vercel 自動 build 並重新生成 |
| **定時自動** | 設定 `VERCEL_DEPLOY_HOOK` secret，每 6 小時 GitHub Actions 觸發 Vercel 重建 |
| **本地預覽** | `py generate.py` → 瀏覽器開啟 `public/index.html` |

---

## 第四步：日常使用

### 手機記錄價格

1. 打開 Google Sheets App
2. 新增一列：日期、時段、型號、容量、顏色、回收價
3. 觸發重新部署（push 或等定時任務）
4. 手機打開 Vercel 網址查看最新圖表

### 修改試算表名稱

編輯 `generate.py` 頂部的設定：

```python
SPREADSHEET_NAME = "你的試算表名稱"
WORKSHEET_NAME = "Sheet1"  # 或你的工作表名稱
```

---

## 常見問題

**Q: 沒有 Google 憑證能跑嗎？**
A: 可以。程式會自動使用 `data/價格表.xlsx` 作為備用數據源。

**Q: Vercel 部署失敗（numpy 版本衝突）？**
A: 本專案已移除 `pandas`/`numpy` 依賴，改用 `openpyxl` + 純 Python 處理數據，避免 Vercel Python 3.14 與 3.12 版本不一致問題。`buildCommand` 只需 `python3 generate.py`。若 Build 需要 Google 憑證，請在 Vercel → Settings → Environment Variables 新增 `GOOGLE_SERVICE_ACCOUNT_JSON`。

**Q: 圖表篩選器沒反應？**
A: 確認瀏覽器允許 JavaScript，且網路能載入 Plotly CDN。

**Q: 想加更多 iPhone 型號？**
A: 直接在 Google Sheets 新增列即可，重新生成 HTML 後圖表會自動出現新線條。

---

## 技術棧

- **Python 3.11+** — 數據處理
- **gspread** — Google Sheets API
- **pandas + openpyxl** — 資料讀取
- **Plotly** — 互動式折線圖
- **Vercel** — 靜態網站託管
- **GitHub Actions** — CI/CD 自動化（選用）
