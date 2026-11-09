@echo off
cd /d "%~dp0"

if exist "web_app.exe" (
  start "" "%~dp0web_app.exe"
  exit /b 0
)

if exist "invoice_system\web_app.py" (
  where py >nul 2>nul
  if %errorlevel%==0 (
    py -3 "invoice_system\web_app.py"
    pause
    exit /b 0
  )

  where python >nul 2>nul
  if %errorlevel%==0 (
    python "invoice_system\web_app.py"
    pause
    exit /b 0
  )
)

echo web_app.exe was not found.
echo Python was not found either.
echo.
echo If this is a no-Python computer, build web_app.exe on another Windows computer first.
echo Run build_exe.cmd on a Windows computer with Python, then copy this whole folder here.
pause
exit /b 1
