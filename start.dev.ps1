# Windows PowerShell 启动脚本（开发热更新）
$ErrorActionPreference = "Stop"
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$port = 8010
$listeners = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
if ($listeners) {
  $pids = $listeners | Select-Object -ExpandProperty OwningProcess -Unique
  foreach ($pid in $pids) {
    Write-Host "Port $port in use by PID $pid. Killing..." -ForegroundColor Yellow
    Stop-Process -Id $pid -Force
  }
  Start-Sleep -Milliseconds 500
}

if (-not (Test-Path ".venv")) {
  python -m venv .venv
}

.\.venv\Scripts\Activate.ps1

$ok = $true
try {
  python -c "import fastapi, uvicorn, pydantic" | Out-Null
} catch {
  $ok = $false
}
if (-not $ok) {
  Write-Host "Missing dependencies. Run: .venv\\Scripts\\python -m pip install -r requirements.txt" -ForegroundColor Red
  exit 1
}

Write-Host "Starting DEV API at http://127.0.0.1:$port" -ForegroundColor Cyan
uvicorn backend:app --host 0.0.0.0 --port $port --reload

