@echo off
chcp 65001 >nul 2>&1
title PPT 도식 자동변환

echo.
echo ========================================
echo   PPT 도식 자동변환
echo   서버/API 키 불필요 - 로컬 실행
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
python -c "import pptx; import transformers; import win32com.client" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [설치] 필요한 패키지 설치 중...
    pip install -r requirements.txt --quiet
    pip install pywin32 --quiet
)

echo [안내] PowerPoint가 열려 있어야 합니다.
echo        현재 열린 프레젠테이션을 자동으로 처리합니다.
echo        최초 실행 시 비전 모델을 자동 다운로드합니다 (약 3.5GB).
echo.

:: 실행
python -m src.ppt_plugin %*

echo.
pause
