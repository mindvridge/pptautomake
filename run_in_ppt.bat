@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

echo.
echo ========================================
echo   PPT AutoMake - 도식 변환기
echo ========================================
echo.

:: Python 감지 및 자동 설치
call "%~dp0_setup_python.bat"
if "!PY!"=="" (
    pause
    exit /b 1
)

:: 프로젝트 디렉토리로 이동
cd /d "%~dp0"

:: 의존성 확인
!PY! -c "import pptx; import transformers" >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo [설치] 필수 패키지 설치 중...
    !PY! -m pip install -r requirements.txt
    if !ERRORLEVEL! neq 0 (
        echo [오류] 패키지 설치 실패.
        echo        직접 실행: !PY! -m pip install -r requirements.txt
        pause
        exit /b 1
    )
)

!PY! -c "import win32com.client" >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo [설치] pywin32 설치 중...
    !PY! -m pip install pywin32
    if !ERRORLEVEL! neq 0 (
        echo [오류] pywin32 설치 실패.
        pause
        exit /b 1
    )
)

echo.
echo [정보] PowerPoint에서 프레젠테이션을 열어두세요.
echo        첫 실행 시 비전 모델 다운로드 (~3.5GB).
echo.

:: 실행
!PY! -m src.ppt_plugin %*

echo.
pause
