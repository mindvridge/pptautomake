@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title PPT AutoMake Server

echo.
echo ========================================
echo   PPT AutoMake - Add-in 서버 (HTTPS)
echo ========================================
echo.

:: Python 감지 - 'py' (Python Launcher) 우선, 'python' 대체
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
    echo [오류] Python이 설치되지 않았거나 정상 작동하지 않습니다.
    echo.
    echo   해결 방법:
    echo     1. https://www.python.org/downloads/ 에서 Python 설치
    echo        중요: 설치 시 "Add Python to PATH" 체크 필수!
    echo.
    echo     2. 이미 설치된 경우, Microsoft Store 앱 별칭 비활성화:
    echo        설정 - 앱 - 앱 실행 별칭
    echo        "python.exe" 와 "python3.exe" 끄기
    echo.
    pause
    exit /b 1
)
echo [정보] Python: !PY!

:: 프로젝트 디렉토리로 이동
cd /d "%~dp0"

:: 의존성 확인
echo [1/3] 의존성 확인 중...
!PY! -c "import flask; import flask_cors; import pptx" >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo       패키지 설치 중...
    !PY! -m pip install -r requirements.txt
    if !ERRORLEVEL! neq 0 (
        echo [오류] 패키지 설치 실패.
        pause
        exit /b 1
    )
)
echo       완료

:: API 키 확인
echo [2/3] API 키 확인 중...
if "%GEMINI_API_KEY%"=="" (
    if "%ANTHROPIC_API_KEY%"=="" (
        echo.
        echo [경고] API 키가 설정되지 않았습니다.
        echo        도식 분석을 위해 GEMINI_API_KEY 또는 ANTHROPIC_API_KEY를 설정하세요.
        echo.
        echo   예시:
        echo     set GEMINI_API_KEY=AIza...
        echo.
    ) else (
        echo       ANTHROPIC_API_KEY 확인됨
    )
) else (
    echo       GEMINI_API_KEY 확인됨
)

:: 서버 시작
echo [3/3] 서버 시작 중...
echo.
echo ----------------------------------------
echo   서버: https://localhost:5000
echo   상태: https://localhost:5000/api/health
echo ----------------------------------------
echo.
echo   PowerPoint에서 사용하기:
echo     1. 삽입 - 내 추가 기능 - 사용자 지정 추가 기능 업로드
echo     2. addin\manifest.xml 선택
echo     3. 홈 탭 - "도식 변환" 클릭
echo.
echo   서버 중지: Ctrl+C
echo ========================================
echo.

!PY! run_server.py

pause
