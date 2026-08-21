@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo 没有找到 Python。请先安装 Python，安装时勾选 Add python.exe to PATH。
    pause
    exit /b 1
)

python -c "import rapidocr, openpyxl, cv2, PIL, pyzbar" >nul 2>nul
if errorlevel 1 (
    echo 第一次使用，正在安装需要的组件，请稍等...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo 安装失败，请把上面的报错发给技术人员。
        pause
        exit /b 1
    )
)

where pythonw >nul 2>nul
if not errorlevel 1 (
    start "" pythonw src\app.py
    exit /b 0
)

python src\app.py
if errorlevel 1 pause
