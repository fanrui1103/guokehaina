@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动送货单出货报告工具...
python app.py
pause
