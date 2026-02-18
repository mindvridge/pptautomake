@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

echo.
echo ========================================
echo   PPT AutoMake - Diagram Converter
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

:: Move to project directory
cd /d "%~dp0"

:: Check dependencies
!PY! -c "import pptx; import transformers" >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo [SETUP] Installing required packages...
    !PY! -m pip install -r requirements.txt
    if !ERRORLEVEL! neq 0 (
        echo [ERROR] Package install failed.
        echo         Try running: !PY! -m pip install -r requirements.txt
        pause
        exit /b 1
    )
)

!PY! -c "import win32com.client" >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo [SETUP] Installing pywin32...
    !PY! -m pip install pywin32
    if !ERRORLEVEL! neq 0 (
        echo [ERROR] pywin32 install failed.
        pause
        exit /b 1
    )
)

echo [INFO] PowerPoint must be open with a presentation.
echo        First run will download the vision model (~3.5GB).
echo.

:: Run
!PY! -m src.ppt_plugin %*

echo.
pause
