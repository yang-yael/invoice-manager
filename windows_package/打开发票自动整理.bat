@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "web_app.exe" (
  start "" "web_app.exe"
  exit /b
)

if exist "invoice_system\web_app.py" (
  py -3 "invoice_system\web_app.py"
  pause
  exit /b
)

echo 未找到 web_app.exe 或 invoice_system\web_app.py
pause
