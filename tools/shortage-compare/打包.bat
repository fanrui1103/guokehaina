@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在安装打包工具...
python -m pip install -q pyinstaller pandas openpyxl flask
echo 正在打包成 exe（可能要几分钟）...
python -m PyInstaller --noconfirm --clean --onedir --console --name "欠料对照小程序" --add-data "templates;templates" --hidden-import compare --collect-submodules flask --collect-submodules jinja2 --collect-submodules openpyxl app.py
copy /Y "使用说明.txt" "dist\欠料对照小程序\使用说明.txt" >nul
echo.
echo 完成。发给同事的是这个文件夹：
echo   dist\欠料对照小程序
echo 请把整个文件夹打成压缩包再发，不要只发 exe。
pause
