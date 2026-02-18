@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title PPT AutoMake Server (HTTP Dev)

echo.
echo ========================================
echo   PPT AutoMake - Dev Server (HTTP)
echo ========================================
echo.

:: Check Python
where python >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo [ERROR] Python is not installed.
    pause
    exit /b 1
)

cd /d "%~dp0"

:: Check dependencies
python -c "import flask; import flask_cors; import pptx" >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo Installing packages...
    pip install -r requirements.txt --quiet
)

echo.
echo   Server: http://localhost:5000
echo   Test in browser directly.
echo   Press Ctrl+C to stop.
echo.

python run_server.py --no-ssl

pause
