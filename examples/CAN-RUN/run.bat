@echo off
setlocal

REM Auto-run: find and run the .py file in current directory
REM Works regardless of the .py filename
REM
REM To switch project config, change CAN_PROJECT below:
REM   config_dazhong = VW project (default)
REM   config_tbox    = TBOX project
set CAN_PROJECT=config_guangqi

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

set "PYTHONPATH=%SCRIPT_DIR%\..\..;%PYTHONPATH%"

cd /d "%SCRIPT_DIR%"

set "TARGET_PY="
for %%f in (*.py) do (
    set "TARGET_PY=%%f"
)

if "%TARGET_PY%"=="" (
    echo [ERROR] No .py file found in current directory
    pause
    exit /b 1
)

echo ============================================================
echo   Running: %TARGET_PY%
echo   Config: %CAN_PROJECT%
echo   Dir: %SCRIPT_DIR%
echo ============================================================
echo.

python "%TARGET_PY%"

echo.
pause
