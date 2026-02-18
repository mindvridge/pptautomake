@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

echo.
echo ========================================
echo   PPT AutoMake - Diagram Converter
echo ========================================
echo.

:: Check Python
where python >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo [ERROR] Python is not installed.
    echo         https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Move to project directory
cd /d "%~dp0"

:: Check dependencies
python -c "import pptx; import transformers" >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo [SETUP] Installing required packages...
    pip install -r requirements.txt --quiet
)

python -c "import win32com.client" >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo [SETUP] Installing pywin32...
    pip install pywin32 --quiet
)

echo [INFO] PowerPoint must be open with a presentation.
echo        First run will download the vision model (~3.5GB).
echo.

:: Run
python -m src.ppt_plugin %*

echo.
pause
