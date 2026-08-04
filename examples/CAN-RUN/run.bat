@echo off
REM Launch run.py -- tries system Python, then 32-bit, then PATH
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0run.py"
) else if exist "%LOCALAPPDATA%\Programs\Python\Python311-32\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python311-32\python.exe" "%~dp0run.py"
) else (
    python "%~dp0run.py" 2>nul
    if errorlevel 1 (
        echo [ERROR] No Python found. Please run install_env.bat first.
        pause
        exit /b 1
    )
)
pause
