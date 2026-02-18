@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title PPT AutoMake Server

echo.
echo ========================================
echo   PPT AutoMake - Add-in Server (HTTPS)
echo ========================================
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

:: Check dependencies
echo [1/3] Checking dependencies...
python -c "import flask; import flask_cors; import pptx" >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo       Installing packages...
    python -m pip install -r requirements.txt --quiet
    if !ERRORLEVEL! neq 0 (
        echo [ERROR] Package install failed.
        pause
        exit /b 1
    )
)
echo       OK

:: Check API key
echo [2/3] Checking API key...
if "%GEMINI_API_KEY%"=="" (
    if "%ANTHROPIC_API_KEY%"=="" (
        echo.
        echo [WARN] No API key found.
        echo        Set GEMINI_API_KEY or ANTHROPIC_API_KEY for diagram analysis.
        echo.
        echo   Example:
        echo     set GEMINI_API_KEY=AIza...
        echo.
    ) else (
        echo       ANTHROPIC_API_KEY found
    )
) else (
    echo       GEMINI_API_KEY found
)

:: Start server
echo [3/3] Starting server...
echo.
echo ----------------------------------------
echo   Server: https://localhost:5000
echo   Health: https://localhost:5000/api/health
echo ----------------------------------------
echo.
echo   To use in PowerPoint:
echo     1. Insert - My Add-ins - Upload Custom Add-in
echo     2. Select addin\manifest.xml
echo     3. Home tab - Click "Diagram Convert"
echo.
echo   Press Ctrl+C to stop server.
echo ========================================
echo.

python run_server.py

pause
