@echo off
chcp 65001 >nul 2>&1
title PPT 도식 자동변환 서버 (HTTP 개발용)

echo.
echo ========================================
echo   PPT 도식 자동변환 - 개발용 서버 (HTTP)
echo ========================================
echo.

:: Python 확인
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [오류] Python이 설치되어 있지 않습니다.
    pause
    exit /b 1
)

cd /d "%~dp0"

:: 의존성 확인
python -c "import flask; import flask_cors; import pptx" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo 패키지 설치 중...
    pip install -r requirements.txt --quiet
)

echo.
echo   서버 주소: http://localhost:5000
echo   브라우저에서 바로 테스트할 수 있습니다.
echo   종료: Ctrl+C
echo.

python run_server.py --no-ssl

pause
