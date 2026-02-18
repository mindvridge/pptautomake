@echo off
chcp 65001 >nul 2>&1
title PPT AutoMake - 플러그인 설치

echo.
echo ========================================================
echo   PPT AutoMake - 원클릭 설치
echo   PowerPoint에 "도식 변환" 리본 탭을 자동 추가합니다
echo ========================================================
echo.

:: Python 확인
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo        https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

:: 프로젝트 디렉토리로 이동
cd /d "%~dp0"

:: 설치 실행
python setup_plugin.py

echo.
pause
