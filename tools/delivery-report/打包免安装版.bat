@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在打包免安装版，请稍候（可能需要几分钟）...
python -m pip install -q pyinstaller flask pymupdf openpyxl
if errorlevel 1 (
  echo pip 安装失败
  pause
  exit /b 1
)

if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist

python -m PyInstaller --noconfirm --clean --onedir --console ^
  --name "送货单出货报告工具" ^
  --add-data "templates;templates" ^
  --hidden-import=excel_parser ^
  --hidden-import=report_gen ^
  --hidden-import=template_match ^
  --hidden-import=size_gen ^
  --collect-all=fitz ^
  --collect-all=flask ^
  --collect-all=openpyxl ^
  app.py

if errorlevel 1 (
  echo 打包失败
  pause
  exit /b 1
)

echo.
echo 打包完成：dist\送货单出货报告工具\
pause
