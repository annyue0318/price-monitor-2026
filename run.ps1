# 一鍵執行：自動使用 .venv 內的 Python，避免 py 啟動器跑錯解譯器
$venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"

if (-not (Test-Path $venvPython)) {
    Write-Host "找不到 .venv，正在建立虛擬環境並安裝套件…" -ForegroundColor Yellow
    py -m venv .venv
    & $venvPython -m pip install -r requirements.txt
}

& $venvPython generate.py
