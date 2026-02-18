@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

echo.
echo ========================================================
echo   PPT AutoMake - 플러그인 설정
echo   PowerPoint에 "도식 변환" 리본 탭 추가
echo ========================================================
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

:: 설정 실행
!PY! setup_plugin.py

echo.
pause
