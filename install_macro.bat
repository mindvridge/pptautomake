@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

echo.
echo ========================================
echo   PPT AutoMake - 매크로 설치
echo ========================================
echo.

:: 프로젝트 디렉토리로 이동
cd /d "%~dp0"
set "PROJECT_PATH=%cd%"

echo [1/2] 프로젝트 경로 등록 중...
setx PPTAUTOMAKE_PATH "%PROJECT_PATH%" >nul 2>&1
echo       경로: %PROJECT_PATH%
echo       환경변수 PPTAUTOMAKE_PATH에 저장됨
echo.

echo [2/2] VBA 매크로 설치 안내
echo.
echo   PowerPoint에서 아래 순서로 진행하세요:
echo.
echo   1. PowerPoint 열기
echo   2. Alt+F11 (VBA 편집기)
echo   3. [삽입] - [모듈] 클릭
echo   4. 아래 파일 내용 복사 붙여넣기:
echo.
echo      %PROJECT_PATH%\addin\PPTAutoMake.bas
echo.
echo   5. VBA 편집기 닫기 (Alt+Q)
echo   6. 파일 - 다른 이름으로 저장 - .pptm 형식
echo.
echo   매크로 실행 방법:
echo     - Alt+F8 - PPTAutoMake_Run 선택 - 실행
echo.
echo ----------------------------------------
echo   사용 가능한 매크로:
echo     PPTAutoMake_Run            전체 변환
echo     PPTAutoMake_Analyze        분석만
echo     PPTAutoMake_SelectedSlides 선택 슬라이드
echo ----------------------------------------
echo.

:: PPTAutoMake.bas 메모장으로 열기
echo   VBA 코드 파일을 메모장에서 여는 중...
start notepad "%PROJECT_PATH%\addin\PPTAutoMake.bas"

echo.
pause
