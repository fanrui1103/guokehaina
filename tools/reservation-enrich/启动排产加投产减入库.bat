@echo off
chcp 65001 >nul
cd /d "%~dp0"
python add_diff_column.py
if errorlevel 1 pause
