@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 出货报告生成工具
echo.
echo 正在启动出货报告生成工具...
echo 启动后请用浏览器打开： http://127.0.0.1:8765
echo 不要关闭本窗口（关闭就等于退出工具）
echo.
where python >nul 2>nul
if errorlevel 1 (
  echo 未找到 Python，请先安装 Python 3，并勾选 Add to PATH。
  pause
  exit /b 1
)
start "" http://127.0.0.1:8765
python app.py
pause
