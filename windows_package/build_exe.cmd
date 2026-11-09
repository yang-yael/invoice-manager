@echo off
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  set PY=py -3
) else (
  where python >nul 2>nul
  if %errorlevel%==0 (
    set PY=python
  ) else (
    echo Python was not found.
    echo Install Python on this build computer only, then run build_exe.cmd again.
    pause
    exit /b 1
  )
)

set APP=
for /r "%~dp0" %%F in (web_app.py) do (
  set APP=%%F
  goto found_app
)

:found_app
if "%APP%"=="" (
  echo web_app.py was not found under this folder.
  echo Please make sure invoice_system\web_app.py exists.
  echo Current folder:
  cd
  echo.
  echo Files here:
  dir
  pause
  exit /b 1
)

echo Found app:
echo %APP%

echo Installing build dependencies...
%PY% -m pip install --upgrade pip
%PY% -m pip install openpyxl pyinstaller
if not %errorlevel%==0 (
  echo Failed to install dependencies.
  pause
  exit /b 1
)

echo Building web_app.exe...
%PY% -m PyInstaller --onefile --name web_app "%APP%"
if not %errorlevel%==0 (
  echo Build failed.
  pause
  exit /b 1
)

if not exist "dist\web_app.exe" (
  echo dist\web_app.exe was not created.
  pause
  exit /b 1
)

copy /Y "dist\web_app.exe" "web_app.exe" >nul
echo Done. web_app.exe was created.
echo You can now copy this whole folder to computers without Python.
pause
exit /b 0
