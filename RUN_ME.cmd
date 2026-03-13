@echo off
setlocal

cd /d "%~dp0"

echo ==================================================
echo ROUND BAR NESTING TEMPLATE BUILDER
echo Folder: %CD%
echo ==================================================
echo.

where powershell >nul 2>nul
if errorlevel 1 (
    echo ERROR: powershell.exe was not found on PATH.
    echo.
    pause
    exit /b 1
)

if not exist "%CD%\RUN_ME.ps1" (
    echo ERROR: RUN_ME.ps1 not found in:
    echo %CD%
    echo.
    pause
    exit /b 1
)

echo Launching PowerShell builder...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%CD%\RUN_ME.ps1"
set ERRLVL=%ERRORLEVEL%

echo.
echo Builder exit code: %ERRLVL%
echo.

if exist "%CD%\build.log" (
    echo Build log:
    echo %CD%\build.log
)

if exist "%CD%\ROUND_BAR_Nesting_Template.validation.json" (
    echo Validation JSON:
    echo %CD%\ROUND_BAR_Nesting_Template.validation.json
)

echo.
pause
exit /b %ERRLVL%