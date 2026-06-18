@echo off
chcp 65001 >nul
setlocal
set PORT=8010

for /f "tokens=5" %%a in ('netstat -ano ^| findstr /R /C:":%PORT% .*LISTENING"') do (
  echo Port %PORT% in use by PID %%a. Killing...
  taskkill /F /PID %%a >nul 2>nul
)

if not exist .venv (
  python -m venv .venv
)

call .venv\Scripts\activate.bat

REM Validate runtime dependencies before startup.
python -c "import fastapi,uvicorn,pydantic,multipart,openpyxl,bs4,playwright,chinese_calendar" 1>nul 2>nul
if errorlevel 1 (
  echo Missing dependencies. Trying to install from requirements.txt ...
  pip install -r requirements.txt || (
    echo Failed to install dependencies due to network issue.
    echo Please retry later or configure pip mirror.
    exit /b 1
  )
)

echo Starting API at http://0.0.0.0:%PORT%
uvicorn backend:app --host 0.0.0.0 --port %PORT%

