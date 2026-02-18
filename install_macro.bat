@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

echo.
echo ========================================
echo   PPT AutoMake - Macro Install
echo ========================================
echo.

:: Move to project directory
cd /d "%~dp0"
set "PROJECT_PATH=%cd%"

echo [1/2] Registering project path...
setx PPTAUTOMAKE_PATH "%PROJECT_PATH%" >nul 2>&1
echo       Path: %PROJECT_PATH%
echo       Saved to env var PPTAUTOMAKE_PATH
echo.

echo [2/2] VBA Macro Install Guide
echo.
echo   Follow these steps in PowerPoint:
echo.
echo   1. Open PowerPoint
echo   2. Press Alt+F11 (VBA Editor)
echo   3. Click [Insert] - [Module]
echo   4. Copy and paste from this file:
echo.
echo      %PROJECT_PATH%\addin\PPTAutoMake.bas
echo.
echo   5. Close VBA Editor (Alt+Q)
echo   6. File - Save As - .pptm format
echo.
echo   Run macros in PowerPoint:
echo     - Alt+F8 - Select PPTAutoMake_Run - Run
echo.
echo ----------------------------------------
echo   Available macros:
echo     PPTAutoMake_Run            Full convert
echo     PPTAutoMake_Analyze        Analyze only
echo     PPTAutoMake_SelectedSlides Selected slides
echo ----------------------------------------
echo.

:: Open PPTAutoMake.bas in notepad
echo   Opening VBA code file in Notepad...
start notepad "%PROJECT_PATH%\addin\PPTAutoMake.bas"

echo.
pause
