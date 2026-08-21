@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在准备小程序（第一次可能会装一点组件）...
python -m pip install -r requirements.txt -q
echo 正在打开网页...
python app.py
pause
