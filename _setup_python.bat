@echo off
:: ============================================
:: Python 자동 감지 및 설치 스크립트
:: 다른 bat 파일에서 call 로 호출
:: 호출 후 !PY! 변수에 Python 경로 설정됨
:: ============================================
:: 주의: setlocal 사용하지 않음 (호출자의 변수 공간 사용)

set "PY="

:: ──────────────────────────────────────
:: 1단계: 기존 Python 감지
:: ──────────────────────────────────────

:: py (Python Launcher) 시도
py -c "import sys; sys.exit(0)" >nul 2>&1
if !ERRORLEVEL! equ 0 (
    set "PY=py"
    goto :_sp_done
)

:: python 시도
python -c "import sys; sys.exit(0)" >nul 2>&1
if !ERRORLEVEL! equ 0 (
    set "PY=python"
    goto :_sp_done
)

:: 직접 경로 탐색
call :_sp_search_paths
if not "!PY!"=="" goto :_sp_done

:: ──────────────────────────────────────
:: 2단계: 자동 설치
:: ──────────────────────────────────────

echo.
echo [알림] Python이 감지되지 않았습니다.
echo [설치] Python을 자동으로 설치합니다...
echo.

set "_SP_INSTALLED=0"

:: winget 시도 (Windows 10/11 기본 패키지 관리자)
where winget >nul 2>&1
if !ERRORLEVEL! equ 0 (
    echo [설치] winget으로 Python 3.12 설치 중...
    echo        (몇 분 소요될 수 있습니다)
    echo.
    winget install Python.Python.3.12 --silent --scope user --accept-package-agreements --accept-source-agreements
    if !ERRORLEVEL! equ 0 (
        set "_SP_INSTALLED=1"
        echo.
        echo [설치] winget 설치 완료!
    ) else (
        echo [경고] winget 설치 실패. 직접 다운로드를 시도합니다...
    )
)

:: curl로 다운로드 (winget 실패 시)
if "!_SP_INSTALLED!"=="0" (
    where curl >nul 2>&1
    if !ERRORLEVEL! equ 0 (
        echo [설치] Python 설치 파일 다운로드 중...
        curl -L -o "%TEMP%\python_installer.exe" "https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe" 2>nul
        if exist "%TEMP%\python_installer.exe" (
            echo [설치] Python 설치 중 (약 1-2분 소요)...
            "%TEMP%\python_installer.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1
            if !ERRORLEVEL! equ 0 (
                set "_SP_INSTALLED=1"
                echo [설치] 설치 완료!
            ) else (
                echo [경고] 자동 설치 실패. 수동으로 설치해주세요.
            )
            del "%TEMP%\python_installer.exe" >nul 2>&1
        ) else (
            echo [경고] 다운로드 실패.
        )
    )
)

:: bitsadmin으로 다운로드 (curl도 없을 때)
if "!_SP_INSTALLED!"=="0" (
    where bitsadmin >nul 2>&1
    if !ERRORLEVEL! equ 0 (
        echo [설치] bitsadmin으로 Python 다운로드 중...
        bitsadmin /transfer "PythonInstall" /priority high "https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe" "%TEMP%\python_installer.exe" >nul 2>&1
        if exist "%TEMP%\python_installer.exe" (
            echo [설치] Python 설치 중 (약 1-2분 소요)...
            "%TEMP%\python_installer.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1
            if !ERRORLEVEL! equ 0 (
                set "_SP_INSTALLED=1"
                echo [설치] 설치 완료!
            )
            del "%TEMP%\python_installer.exe" >nul 2>&1
        )
    )
)

:: 설치 실패
if "!_SP_INSTALLED!"=="0" (
    echo.
    echo [오류] Python 자동 설치에 실패했습니다.
    echo.
    echo   수동 설치 방법:
    echo     1. https://www.python.org/downloads/ 에서 Python 다운로드
    echo     2. 설치 시 "Add Python to PATH" 반드시 체크!
    echo     3. 설치 후 이 파일을 다시 실행
    echo.
    goto :_sp_done
)

:: ──────────────────────────────────────
:: 3단계: 설치 후 재감지
:: ──────────────────────────────────────

echo [설치] Python 경로 갱신 중...

:: 현재 세션 PATH에 일반적인 Python 경로 추가
set "PATH=!PATH!;%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts"
set "PATH=!PATH!;%LOCALAPPDATA%\Programs\Python\Python313;%LOCALAPPDATA%\Programs\Python\Python313\Scripts"
set "PATH=!PATH!;C:\Python312;C:\Python312\Scripts;C:\Python313;C:\Python313\Scripts"

:: py 재시도
py -c "import sys; sys.exit(0)" >nul 2>&1
if !ERRORLEVEL! equ 0 (
    set "PY=py"
    goto :_sp_done
)

:: python 재시도
python -c "import sys; sys.exit(0)" >nul 2>&1
if !ERRORLEVEL! equ 0 (
    set "PY=python"
    goto :_sp_done
)

:: 직접 경로 탐색
call :_sp_search_paths
if not "!PY!"=="" goto :_sp_done

echo.
echo [오류] Python 설치 후에도 감지되지 않습니다.
echo        명령 프롬프트를 닫고 다시 열어서 실행해주세요.
echo.
goto :_sp_done

:: ──────────────────────────────────────
:: 서브루틴: 일반 경로에서 Python 검색
:: ──────────────────────────────────────
:_sp_search_paths
:: AppData Python (사용자 설치)
for /f "tokens=*" %%d in ('dir /b /ad "%LOCALAPPDATA%\Programs\Python\Python3*" 2^>nul') do (
    if exist "%LOCALAPPDATA%\Programs\Python\%%d\python.exe" (
        set "PY=%LOCALAPPDATA%\Programs\Python\%%d\python.exe"
        exit /b 0
    )
)
:: C:\Python3xx (전체 사용자 설치)
for /f "tokens=*" %%d in ('dir /b /ad "C:\Python3*" 2^>nul') do (
    if exist "C:\%%d\python.exe" (
        set "PY=C:\%%d\python.exe"
        exit /b 0
    )
)
:: Program Files
for /f "tokens=*" %%d in ('dir /b /ad "%ProgramFiles%\Python3*" 2^>nul') do (
    if exist "%ProgramFiles%\%%d\python.exe" (
        set "PY=%ProgramFiles%\%%d\python.exe"
        exit /b 0
    )
)
exit /b 1

:_sp_done
if not "!PY!"=="" (
    echo [정보] Python: !PY!
)
