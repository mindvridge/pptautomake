@echo off
chcp 65001 >nul 2>&1
title PPT 도식 자동변환

echo.
echo ========================================
echo   PPT 도식 자동변환
echo   서버 없이 PowerPoint에서 직접 실행
echo ========================================
echo.

:: Python 확인
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo        https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 프로젝트 디렉토리로 이동
cd /d "%~dp0"

:: 의존성 확인
python -c "import pptx; import anthropic; import win32com.client" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [설치] 필요한 패키지 설치 중...
    pip install -r requirements.txt --quiet
    pip install pywin32 --quiet
)

:: ANTHROPIC_API_KEY 확인
if "%ANTHROPIC_API_KEY%"=="" (
    echo [경고] ANTHROPIC_API_KEY가 설정되지 않았습니다.
    echo        도식 분석 기능에 API 키가 필요합니다.
    echo.
)

echo [안내] PowerPoint가 열려 있어야 합니다.
echo        현재 열린 프레젠테이션을 자동으로 처리합니다.
echo.

:: 실행
python -m src.ppt_plugin %*

echo.
pause
