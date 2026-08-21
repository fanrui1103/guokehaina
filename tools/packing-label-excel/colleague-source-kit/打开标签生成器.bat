@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>&1
if %errorlevel%==0 (
  python app.py
  if errorlevel 1 pause
  goto :eof
)

where py >nul 2>&1
if %errorlevel%==0 (
  py -3 app.py
  if errorlevel 1 pause
  goto :eof
)

echo 本文件夹是源码版，需要先安装 Python。
echo 安装时请勾选 Add python.exe to PATH。
echo 装好后在本文件夹执行：pip install -r requirements.txt
echo 若只想用、不改代码，请用「发给同事」里的 exe。
pause
