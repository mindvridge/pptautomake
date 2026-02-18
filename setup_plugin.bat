@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

echo.
echo ========================================================
echo   PPT AutoMake - 플러그인 설정
echo   PowerPoint에 "도식 변환" 리본 탭 추가
echo ========================================================
echo.

:: Python 감지 및 자동 설치
call "%~dp0_setup_python.bat"
if "!PY!"=="" (
    pause
    exit /b 1
)

:: 프로젝트 디렉토리로 이동
cd /d "%~dp0"

:: 설정 실행
!PY! setup_plugin.py

echo.
pause
