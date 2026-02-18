@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title PPT AutoMake Server (HTTP Dev)

echo.
echo ========================================
echo   PPT AutoMake - 개발 서버 (HTTP)
echo ========================================
echo.

:: Python 감지 및 자동 설치
call "%~dp0_setup_python.bat"
if "!PY!"=="" (
    pause
    exit /b 1
)

cd /d "%~dp0"

:: 의존성 확인
!PY! -c "import flask; import flask_cors; import pptx" >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo 패키지 설치 중...
    !PY! -m pip install -r requirements.txt
    if !ERRORLEVEL! neq 0 (
        echo [오류] 패키지 설치 실패.
        pause
        exit /b 1
    )
)

echo.
echo   서버: http://localhost:5000
echo   브라우저에서 직접 테스트 가능.
echo   중지: Ctrl+C
echo.

!PY! run_server.py --no-ssl

pause
