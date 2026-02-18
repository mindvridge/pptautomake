@echo off
chcp 65001 >nul 2>&1
title PPT 도식 자동변환 서버

echo.
echo ========================================
echo   PPT 도식 자동변환 - Add-in 서버
echo ========================================
echo.

:: Python 확인
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo        https://www.python.org/downloads/ 에서 설치하세요.
    echo.
    pause
    exit /b 1
)

:: 프로젝트 디렉토리로 이동
cd /d "%~dp0"

:: 의존성 확인 및 설치
echo [1/3] 의존성 확인 중...
python -c "import flask; import flask_cors; import pptx; import anthropic" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo       패키지 설치 중...
    pip install -r requirements.txt --quiet
    if %ERRORLEVEL% neq 0 (
        echo [오류] 패키지 설치에 실패했습니다.
        pause
        exit /b 1
    )
)
echo       완료

:: ANTHROPIC_API_KEY 확인
echo [2/3] API 키 확인 중...
if "%ANTHROPIC_API_KEY%"=="" (
    echo.
    echo [경고] ANTHROPIC_API_KEY 환경변수가 설정되지 않았습니다.
    echo        도식 분석 기능을 사용하려면 API 키가 필요합니다.
    echo.
    echo   설정 방법:
    echo     set ANTHROPIC_API_KEY=sk-ant-...
    echo   또는 시스템 환경변수에 추가하세요.
    echo.
) else (
    echo       완료
)

:: 서버 시작
echo [3/3] 서버 시작 중...
echo.
echo ----------------------------------------
echo   서버 주소: https://localhost:5000
echo   상태 확인: https://localhost:5000/api/health
echo ----------------------------------------
echo.
echo   PowerPoint에서 사용하려면:
echo     1. 삽입 → 내 추가 기능 → 사용자 지정 추가 기능 업로드
echo     2. addin\manifest.xml 선택
echo     3. 홈 탭 → "도식 변환" 버튼 클릭
echo.
echo   서버를 종료하려면 Ctrl+C 를 누르세요.
echo ========================================
echo.

python run_server.py

pause
