# Windows PowerShell 启动脚本（稳定模式）
$ErrorActionPreference = "Stop"
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$port = 8010
$listeners = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
if ($listeners) {
  $pids = $listeners | Select-Object -ExpandProperty OwningProcess -Unique
  foreach ($pid in $pids) {
    Stop-Process -Id $pid -Force -ErrorAction SilentlyContinue
  }
  Start-Sleep -Milliseconds 500
}

if (-not (Test-Path ".venv")) {
  python -m venv .venv
}

.\.venv\Scripts\Activate.ps1

$depCheck = python -c "import fastapi,uvicorn,pydantic" 2>$null; $LASTEXITCODE
if ($depCheck -ne 0) {
  Write-Host "Missing dependencies. Installing from requirements.txt ..." -ForegroundColor Yellow
  pip install -r requirements.txt
}

Write-Host "Starting API at http://127.0.0.1:$port" -ForegroundColor Cyan
uvicorn backend:app --host 0.0.0.0 --port $port

