@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title PPT AutoMake Server (HTTP Dev)

echo.
echo ========================================
echo   PPT AutoMake - Dev Server (HTTP)
echo ========================================
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

cd /d "%~dp0"

:: Check dependencies
!PY! -c "import flask; import flask_cors; import pptx" >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo Installing packages...
    !PY! -m pip install -r requirements.txt
    if !ERRORLEVEL! neq 0 (
        echo [ERROR] Package install failed.
        pause
        exit /b 1
    )
)

echo.
echo   Server: http://localhost:5000
echo   Test in browser directly.
echo   Press Ctrl+C to stop.
echo.

!PY! run_server.py --no-ssl

pause
