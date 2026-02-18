@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

echo.
echo ========================================================
echo   PPT AutoMake - Plugin Setup
echo   Adds "Diagram Convert" ribbon tab to PowerPoint
echo ========================================================
echo.

:: Detect Python - try 'py' (Python Launcher) first, then 'python'
set "PY="
py -c "import sys; sys.exit(0)" >nul 2>&1
if !ERRORLEVEL! equ 0 (
    set "PY=py"
)
if "!PY!"=="" (
    python -c "import sys; sys.exit(0)" >nul 2>&1
    if !ERRORLEVEL! equ 0 (
        set "PY=python"
    )
)
if "!PY!"=="" (
    echo [ERROR] Python is not installed or not working.
    echo.
    echo   Fix options:
    echo     1. Install Python from https://www.python.org/downloads/
    echo        IMPORTANT: Check "Add Python to PATH" during install.
    echo.
    echo     2. If already installed, disable Microsoft Store alias:
    echo        Settings - Apps - App execution aliases
    echo        Turn OFF "python.exe" and "python3.exe"
    echo.
    pause
    exit /b 1
)
echo [INFO] Using: !PY!

:: Move to project directory
cd /d "%~dp0"

:: Run setup
!PY! setup_plugin.py

echo.
pause
