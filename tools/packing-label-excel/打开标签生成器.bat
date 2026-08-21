@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "%~dp0发给同事\益佳通标签生成器.exe" (
  start "" "%~dp0发给同事\益佳通标签生成器.exe"
  goto :eof
)

if exist "%~dp0dist\益佳通标签生成器.exe" (
  start "" "%~dp0dist\益佳通标签生成器.exe"
  goto :eof
)

where python >nul 2>&1
if %errorlevel%==0 (
  python app.py
  if errorlevel 1 pause
  goto :eof
)

echo 没有找到可运行的程序。
pause
