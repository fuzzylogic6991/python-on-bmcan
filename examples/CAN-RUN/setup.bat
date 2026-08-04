@echo off
REM Launch setup.py -- tries system Python, then 32-bit, then PATH
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0setup.py"
) else if exist "%LOCALAPPDATA%\Programs\Python\Python311-32\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python311-32\python.exe" "%~dp0setup.py"
) else (
    python "%~dp0setup.py" 2>nul
    if errorlevel 1 (
        echo [ERROR] No Python found. Please install Python first.
        pause
        exit /b 1
    )
)
pause
