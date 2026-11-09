@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo 正在安装打包依赖...
py -3 -m pip install --upgrade pip
py -3 -m pip install openpyxl pyinstaller

echo 正在构建 web_app.exe...
py -3 -m PyInstaller --onefile --name web_app --paths invoice_system invoice_system\web_app.py

if not exist "dist\web_app.exe" (
  echo 构建失败，请检查上面的错误信息。
  pause
  exit /b 1
)

copy /Y "dist\web_app.exe" "web_app.exe" >nul
echo 构建完成：web_app.exe
echo 之后其他 Windows 电脑只需要这个文件夹，不需要安装 Python。
pause
