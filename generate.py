#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
iPhone 回收價格監控 — 從 Google Sheets（或本地 Excel）讀取數據，生成互動式 HTML 報表。

執行方式：
    python generate.py

輸出：
    public/index.html
"""

from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio

# ---------------------------------------------------------------------------
# 路徑設定
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent
CREDENTIALS_PATH = ROOT / "credentials" / "service_account.json"
FALLBACK_EXCEL = ROOT / "data" / "價格表.xlsx"
OUTPUT_HTML = ROOT / "public" / "index.html"

# Google Sheets 設定（使用 Google Sheets 時請修改這兩行）
SPREADSHEET_NAME = "iPhone回收價格"  # 試算表名稱
WORKSHEET_NAME = "Sheet1"           # 工作表名稱（預設第一個分頁）

# 欄位對應（Google Sheets / Excel 的欄位名稱）
COLUMN_MAP = {
    "date": "Date",
    "time": "Time",
    "model": "Model",
    "capacity": "Capacity",
    "color": "Color",
    "price": "Price",
}


# ===========================================================================
# 1. 讀取數據
# ===========================================================================

def load_from_google_sheets() -> pd.DataFrame:
    """
    使用 gspread 從 Google Sheets 讀取數據。

    ── Google Cloud Console 設定步驟 ──────────────────────────────────────

    ① 建立 Google Cloud 專案
       前往 https://console.cloud.google.com/
       → 左上角選單 → 「IAM 與管理」→「建立專案」→ 輸入名稱（如 price-monitor）

    ② 啟用 Google Sheets API
       → 「API 和服務」→「程式庫」→ 搜尋「Google Sheets API」→ 啟用

    ③ 建立 Service Account（服務帳戶）
       → 「IAM 與管理」→「服務帳戶」→「建立服務帳戶」
       → 名稱填 price-monitor-bot → 建立並繼續 → 完成

    ④ 下載 JSON 金鑰
       → 點進剛建立的服務帳戶 → 「金鑰」分頁 →「新增金鑰」→「建立新金鑰」→ JSON
       → 瀏覽器會下載一個 .json 檔案
       → 將它重新命名為 service_account.json
       → 放到本專案的 credentials/service_account.json

    ⑤ 分享試算表給 Service Account
       → 用記事本打開 service_account.json，找到 "client_email" 欄位
         格式類似：price-monitor-bot@your-project.iam.gserviceaccount.com
       → 在 Google Sheets 中，點右上角「共用」
       → 貼上上述 email → 權限選「編輯者」→ 傳送
       （Service Account 不是真人，不會收到通知，但已獲得讀取權限）

    ⑥ Vercel 部署時（選用）
       → 在 Vercel 專案 Settings → Environment Variables
       → 新增 GOOGLE_SERVICE_ACCOUNT_JSON，值為整份 JSON 字串
       → 程式會自動從環境變數讀取，無需上傳檔案

    ─────────────────────────────────────────────────────────────────────
    """
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]

    # 優先讀本地 JSON，其次讀 Vercel 環境變數
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if creds_json:
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=scopes)
        print("[INFO] 使用環境變數 GOOGLE_SERVICE_ACCOUNT_JSON 連線 Google Sheets")
    elif CREDENTIALS_PATH.exists():
        creds = Credentials.from_service_account_file(str(CREDENTIALS_PATH), scopes=scopes)
        print(f"[INFO] 使用本地憑證：{CREDENTIALS_PATH}")
    else:
        raise FileNotFoundError("找不到 Google 憑證")

    client = gspread.authorize(creds)
    spreadsheet = client.open(SPREADSHEET_NAME)
    worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    records = worksheet.get_all_records()
    df = pd.DataFrame(records)
    print(f"[INFO] 從 Google Sheets 讀取 {len(df)} 筆資料")
    return df


def load_from_excel() -> pd.DataFrame:
    """備用方案：讀取本地 Excel 檔案。"""
    if not FALLBACK_EXCEL.exists():
        raise FileNotFoundError(
            f"找不到備用 Excel：{FALLBACK_EXCEL}\n"
            "請建立 data/價格表.xlsx，或設定 Google 憑證。"
        )
    df = pd.read_excel(FALLBACK_EXCEL, engine="openpyxl")
    print(f"[INFO] 從本地 Excel 讀取 {len(df)} 筆資料：{FALLBACK_EXCEL}")
    return df


def load_data() -> pd.DataFrame:
    """依序嘗試 Google Sheets → 本地 Excel。"""
    has_env_creds = bool(os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON"))
    has_local_creds = CREDENTIALS_PATH.exists()

    if has_env_creds or has_local_creds:
        try:
            return load_from_google_sheets()
        except Exception as exc:
            print(f"[WARN] Google Sheets 讀取失敗：{exc}")
            print("[WARN] 改為讀取本地 Excel 備用方案…")

    print("[INFO] 未找到 Google 憑證，使用本地 Excel 備用方案")
    return load_from_excel()


# ===========================================================================
# 2. 資料清理
# ===========================================================================

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """標準化欄位名稱、型別，並建立 datetime 欄位。"""
    # 建立欄位對應（支援中英文欄位名 → 統一小寫欄位名）
    rename_map = {}
    col_aliases = {
        "date": ["Date", "日期", "date"],
        "time": ["Time", "時段", "time"],
        "model": ["Model", "型號", "model"],
        "capacity": ["Capacity", "容量", "capacity"],
        "color": ["Color", "顏色", "color"],
        "price": ["Price", "回收價", "price", "價格"],
    }
    for std_name, aliases in col_aliases.items():
        for alias in aliases:
            if alias in df.columns:
                rename_map[alias] = std_name
                break

    df = df.rename(columns=rename_map)

    required = ["date", "time", "model", "capacity", "color", "price"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"缺少必要欄位：{missing}。目前欄位：{list(df.columns)}")

    # 清理空白列
    df = df.dropna(subset=["price"]).copy()
    df["price"] = pd.to_numeric(df["price"], errors="coerce")
    df = df.dropna(subset=["price"])

    # 合併日期 + 時段
    df["date"] = df["date"].astype(str).str.strip()
    df["time"] = df["time"].astype(str).str.strip()
    df["datetime"] = pd.to_datetime(
        df["date"] + " " + df["time"],
        errors="coerce",
    )
    # 若時段解析失敗，只用日期
    mask = df["datetime"].isna()
    if mask.any():
        df.loc[mask, "datetime"] = pd.to_datetime(df.loc[mask, "date"], errors="coerce")

    df = df.dropna(subset=["datetime"]).sort_values("datetime")
    df["model"] = df["model"].astype(str).str.strip()
    df["capacity"] = df["capacity"].astype(str).str.strip()
    df["color"] = df["color"].astype(str).str.strip()
    df["datetime_label"] = df["datetime"].dt.strftime("%Y-%m-%d %H:%M")

    return df.reset_index(drop=True)


# ===========================================================================
# 3. 生成 Plotly 圖表
# ===========================================================================

# Plotly 預設調色盤（不同型號/顏色用不同線條色）
LINE_COLORS = [
    "#636EFA", "#EF553B", "#00CC96", "#AB63FA", "#FFA15A",
    "#19D3F3", "#FF6692", "#B6E880", "#FF97FF", "#FECB52",
    "#1F77B4", "#FF7F0E", "#2CA02C", "#D62728", "#9467BD",
]


def build_figure(df: pd.DataFrame) -> tuple[go.Figure, list[dict]]:
    """
    建立折線圖，每個 (型號, 容量, 顏色) 組合一條線。
    回傳 figure 與 trace 中繼資料（供前端篩選器使用）。
    """
    fig = go.Figure()
    trace_meta: list[dict] = []

    groups = df.groupby(["model", "capacity", "color"], sort=False)
    color_idx = 0

    for (model, capacity, color), group in groups:
        group = group.sort_values("datetime")
        legend_name = f"{model} | {capacity} | {color}"
        line_color = LINE_COLORS[color_idx % len(LINE_COLORS)]
        color_idx += 1

        fig.add_trace(
            go.Scatter(
                x=group["datetime"],
                y=group["price"],
                mode="lines+markers",
                name=legend_name,
                line=dict(color=line_color, width=2.5),
                marker=dict(size=7, color=line_color),
                hovertemplate=(
                    "<b>%{fullData.name}</b><br>"
                    "時間：%{x|%Y-%m-%d %H:%M}<br>"
                    "回收價：$%{y:,.0f}<extra></extra>"
                ),
                visible=True,
            )
        )
        trace_meta.append({
            "model": model,
            "capacity": capacity,
            "color": color,
            "legend": legend_name,
        })

    fig.update_layout(
        title=dict(
            text="📱 iPhone 回收價格走勢",
            font=dict(size=22, color="#1a1a2e"),
            x=0.5,
            xanchor="center",
        ),
        xaxis=dict(
            title="日期 / 時段",
            gridcolor="#e8e8e8",
            tickformat="%m/%d %H:%M",
        ),
        yaxis=dict(
            title="回收價（HKD）",
            gridcolor="#e8e8e8",
            tickformat="$,.0f",
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.35,
            xanchor="center",
            x=0.5,
            font=dict(size=11),
        ),
        hovermode="x unified",
        plot_bgcolor="#fafafa",
        paper_bgcolor="#ffffff",
        margin=dict(l=60, r=30, t=80, b=120),
        height=480,
    )

    return fig, trace_meta


# ===========================================================================
# 4. 組裝完整 HTML（含篩選器 + 數據表格）
# ===========================================================================

def build_html_page(df: pd.DataFrame, fig: go.Figure, trace_meta: list[dict]) -> str:
    """將 Plotly 圖表、篩選器、原始數據表格打包成響應式 HTML。"""

    chart_div = pio.to_html(fig, full_html=False, include_plotlyjs="cdn", div_id="price-chart")

    models = sorted(df["model"].unique())
    capacities = sorted(df["capacity"].unique())
    colors = sorted(df["color"].unique())

    def make_options(items: list[str]) -> str:
        opts = ['<option value="all">全部</option>']
        opts += [f'<option value="{item}">{item}</option>' for item in items]
        return "\n".join(opts)

    # 原始數據表格
    display_df = df[["datetime_label", "model", "capacity", "color", "price"]].copy()
    display_df.columns = ["日期時間", "型號", "容量", "顏色", "回收價"]
    display_df["回收價"] = display_df["回收價"].apply(lambda x: f"${x:,.0f}")

    table_rows = ""
    for _, row in display_df.iterrows():
        table_rows += (
            f"<tr>"
            f"<td>{row['日期時間']}</td>"
            f"<td>{row['型號']}</td>"
            f"<td>{row['容量']}</td>"
            f"<td>{row['顏色']}</td>"
            f"<td class='price'>{row['回收價']}</td>"
            f"</tr>\n"
        )

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    trace_meta_json = json.dumps(trace_meta, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>iPhone 回收價格監控</title>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
                   "Helvetica Neue", Arial, "Noto Sans TC", sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      padding: 16px;
      color: #333;
    }}

    .container {{
      max-width: 960px;
      margin: 0 auto;
    }}

    header {{
      text-align: center;
      color: #fff;
      margin-bottom: 20px;
    }}

    header h1 {{
      font-size: 1.6rem;
      font-weight: 700;
      margin-bottom: 4px;
    }}

    header p {{
      font-size: 0.85rem;
      opacity: 0.85;
    }}

    .card {{
      background: #fff;
      border-radius: 16px;
      box-shadow: 0 8px 32px rgba(0,0,0,0.12);
      padding: 20px;
      margin-bottom: 16px;
    }}

    .filters {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
      gap: 12px;
      margin-bottom: 16px;
    }}

    .filter-group label {{
      display: block;
      font-size: 0.78rem;
      font-weight: 600;
      color: #666;
      margin-bottom: 4px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }}

    .filter-group select {{
      width: 100%;
      padding: 10px 12px;
      border: 2px solid #e0e0e0;
      border-radius: 10px;
      font-size: 0.95rem;
      background: #fafafa;
      cursor: pointer;
      transition: border-color 0.2s;
      -webkit-appearance: none;
      appearance: none;
      background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%23666' d='M6 8L1 3h10z'/%3E%3C/svg%3E");
      background-repeat: no-repeat;
      background-position: right 12px center;
    }}

    .filter-group select:focus {{
      outline: none;
      border-color: #667eea;
    }}

    .stats {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
      gap: 10px;
      margin-bottom: 16px;
    }}

    .stat-box {{
      background: linear-gradient(135deg, #667eea22, #764ba222);
      border-radius: 12px;
      padding: 14px;
      text-align: center;
    }}

    .stat-box .value {{
      font-size: 1.3rem;
      font-weight: 700;
      color: #667eea;
    }}

    .stat-box .label {{
      font-size: 0.75rem;
      color: #888;
      margin-top: 2px;
    }}

    #price-chart {{
      width: 100% !important;
    }}

    .table-wrapper {{
      overflow-x: auto;
      -webkit-overflow-scrolling: touch;
    }}

    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 0.88rem;
    }}

    thead th {{
      background: #667eea;
      color: #fff;
      padding: 10px 8px;
      text-align: left;
      font-weight: 600;
      white-space: nowrap;
      position: sticky;
      top: 0;
    }}

    tbody td {{
      padding: 9px 8px;
      border-bottom: 1px solid #f0f0f0;
    }}

    tbody tr:hover {{
      background: #f8f9ff;
    }}

    td.price {{
      font-weight: 600;
      color: #667eea;
      text-align: right;
    }}

    .section-title {{
      font-size: 1.1rem;
      font-weight: 700;
      margin-bottom: 12px;
      color: #1a1a2e;
    }}

    footer {{
      text-align: center;
      color: rgba(255,255,255,0.7);
      font-size: 0.75rem;
      padding: 16px 0;
    }}

    @media (max-width: 480px) {{
      header h1 {{ font-size: 1.3rem; }}
      .card {{ padding: 14px; border-radius: 12px; }}
      .stat-box .value {{ font-size: 1.1rem; }}
    }}
  </style>
</head>
<body>
  <div class="container">
    <header>
      <h1>📱 iPhone 回收價格監控</h1>
      <p>最後更新：{generated_at}</p>
    </header>

    <!-- 篩選器 -->
    <div class="card">
      <div class="filters">
        <div class="filter-group">
          <label for="filter-model">型號</label>
          <select id="filter-model" onchange="applyFilters()">
            {make_options(models)}
          </select>
        </div>
        <div class="filter-group">
          <label for="filter-capacity">容量</label>
          <select id="filter-capacity" onchange="applyFilters()">
            {make_options(capacities)}
          </select>
        </div>
        <div class="filter-group">
          <label for="filter-color">顏色</label>
          <select id="filter-color" onchange="applyFilters()">
            {make_options(colors)}
          </select>
        </div>
      </div>

      <!-- 摘要統計 -->
      <div class="stats">
        <div class="stat-box">
          <div class="value" id="stat-count">{len(df)}</div>
          <div class="label">資料筆數</div>
        </div>
        <div class="stat-box">
          <div class="value" id="stat-max">${df["price"].max():,.0f}</div>
          <div class="label">最高價</div>
        </div>
        <div class="stat-box">
          <div class="value" id="stat-min">${df["price"].min():,.0f}</div>
          <div class="label">最低價</div>
        </div>
        <div class="stat-box">
          <div class="value" id="stat-avg">${df["price"].mean():,.0f}</div>
          <div class="label">平均價</div>
        </div>
      </div>

      <!-- Plotly 折線圖 -->
      {chart_div}
    </div>

    <!-- 原始數據表格 -->
    <div class="card">
      <div class="section-title">📋 原始數據</div>
      <div class="table-wrapper">
        <table id="data-table">
          <thead>
            <tr>
              <th>日期時間</th>
              <th>型號</th>
              <th>容量</th>
              <th>顏色</th>
              <th>回收價</th>
            </tr>
          </thead>
          <tbody id="table-body">
            {table_rows}
          </tbody>
        </table>
      </div>
    </div>

    <footer>
      Powered by Plotly · 數據來源 Google Sheets
    </footer>
  </div>

  <script>
    const TRACE_META = {trace_meta_json};

    function applyFilters() {{
      const model    = document.getElementById('filter-model').value;
      const capacity = document.getElementById('filter-capacity').value;
      const color    = document.getElementById('filter-color').value;

      const visibility = TRACE_META.map(t => {{
        if (model    !== 'all' && t.model    !== model)    return 'legendonly';
        if (capacity !== 'all' && t.capacity !== capacity) return 'legendonly';
        if (color    !== 'all' && t.color    !== color)    return 'legendonly';
        return true;
      }});

      Plotly.restyle('price-chart', {{ visible: visibility }});

      // 同步篩選表格
      filterTable(model, capacity, color);
    }}

    function filterTable(model, capacity, color) {{
      const rows = document.querySelectorAll('#table-body tr');
      rows.forEach(row => {{
        const cells = row.querySelectorAll('td');
        const rModel    = cells[1].textContent;
        const rCapacity = cells[2].textContent;
        const rColor    = cells[3].textContent;
        const show =
          (model    === 'all' || rModel    === model) &&
          (capacity === 'all' || rCapacity === capacity) &&
          (color    === 'all' || rColor    === color);
        row.style.display = show ? '' : 'none';
      }});
    }}
  </script>
</body>
</html>"""
    return html


# ===========================================================================
# 5. 建立範例 Excel（若不存在）
# ===========================================================================

def create_sample_excel() -> None:
    """若 data/價格表.xlsx 不存在，自動建立範例資料供測試。"""
    FALLBACK_EXCEL.parent.mkdir(parents=True, exist_ok=True)

    sample_data = {
        "Date": [
            "2025-09-20", "2025-09-20", "2025-09-20",
            "2025-09-21", "2025-09-21", "2025-09-21",
            "2025-09-22", "2025-09-22", "2025-09-22",
            "2025-09-20", "2025-09-21", "2025-09-22",
        ],
        "Time": [
            "10:00", "14:00", "18:00",
            "10:00", "14:00", "18:00",
            "10:00", "14:00", "18:00",
            "12:00", "12:00", "12:00",
        ],
        "Model": [
            "iPhone 16 Pro", "iPhone 16 Pro", "iPhone 16 Pro",
            "iPhone 16 Pro", "iPhone 16 Pro", "iPhone 16 Pro",
            "iPhone 16 Pro", "iPhone 16 Pro", "iPhone 16 Pro",
            "iPhone 16", "iPhone 16", "iPhone 16",
        ],
        "Capacity": [
            "256GB", "256GB", "256GB",
            "256GB", "256GB", "256GB",
            "256GB", "256GB", "256GB",
            "128GB", "128GB", "128GB",
        ],
        "Color": [
            "原色鈦金屬", "原色鈦金屬", "原色鈦金屬",
            "原色鈦金屬", "原色鈦金屬", "原色鈦金屬",
            "原色鈦金屬", "原色鈦金屬", "原色鈦金屬",
            "黑色", "黑色", "黑色",
        ],
        "Price": [
            7200, 7150, 7100,
            7050, 7000, 6980,
            6950, 6900, 6850,
            5200, 5150, 5100,
        ],
    }
    pd.DataFrame(sample_data).to_excel(FALLBACK_EXCEL, index=False, engine="openpyxl")
    print(f"[INFO] 已建立範例 Excel：{FALLBACK_EXCEL}")


# ===========================================================================
# 主程式
# ===========================================================================

def main() -> int:
    # Windows 終端機預設 cp1252，需切換 UTF-8 才能正確顯示中文
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    print("=" * 60)
    print("  iPhone 回收價格監控 — HTML 報表生成器")
    print("=" * 60)

    # 確保範例 Excel 存在
    if not FALLBACK_EXCEL.exists():
        create_sample_excel()

    # 讀取 & 清理數據
    df = load_data()
    df = normalize_dataframe(df)
    print(f"[INFO] 有效資料 {len(df)} 筆，"
          f"型號 {df['model'].nunique()} 種，"
          f"日期範圍 {df['datetime'].min()} ~ {df['datetime'].max()}")

    # 生成圖表
    fig, trace_meta = build_figure(df)

    # 組裝 HTML
    html = build_html_page(df, fig, trace_meta)

    # 寫出檔案
    OUTPUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"[OK] 已生成：{OUTPUT_HTML}")
    print(f"     用瀏覽器開啟此檔案即可預覽，或部署到 Vercel。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
