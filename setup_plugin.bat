@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

echo.
echo ========================================================
echo   PPT AutoMake - Plugin Setup
echo   Adds "Diagram Convert" ribbon tab to PowerPoint
echo ========================================================
echo.

:: Check Python
where python >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo [ERROR] Python is not installed.
    echo         https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

:: Move to project directory
cd /d "%~dp0"

:: Run setup
python setup_plugin.py

echo.
pause
