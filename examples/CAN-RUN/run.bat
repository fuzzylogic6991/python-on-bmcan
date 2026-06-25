@echo off
setlocal

REM Auto-run: find and run the .py file in current directory
REM Works regardless of the .py filename

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
echo   Dir: %SCRIPT_DIR%
echo ============================================================
echo.

python "%TARGET_PY%"

echo.
pause
