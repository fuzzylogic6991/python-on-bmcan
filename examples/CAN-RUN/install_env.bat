@echo off
setlocal

echo ============================================================
echo   CAN-RUN Environment Setup
echo ============================================================
echo.

REM Script is in examples\CAN-RUN\, project root is two levels up
set "ROOT=%~dp0..\.."
pushd "%ROOT%"
set "ROOT=%CD%"
popd

echo [1/4] Checking Python ...
python --version
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.8+ first.
    pause
    exit /b 1
)
echo [OK] Python ready
echo.

echo [2/4] Installing custom python-can (with bmcan) ...
echo       Uninstalling PyPI version if exists ...
pip uninstall python-can -y
pip install -e "%ROOT%"
if errorlevel 1 (
    echo [ERROR] python-can install failed
    pause
    exit /b 1
)
echo [OK] python-can installed
echo.

echo [3/4] Installing dependencies (pandas / openpyxl / pyserial) ...
pip install -r "%ROOT%\examples\CAN-RUN\requirements.txt"
if errorlevel 1 (
    echo [ERROR] Dependencies install failed
    pause
    exit /b 1
)
echo [OK] Dependencies installed
echo.

echo [4/4] Verifying bmcan interface ...
python -c "import can; ok='bmcan' in can.VALID_INTERFACES; print('[OK] bmcan available' if ok else '[FAIL] bmcan not available'); exit(0 if ok else 1)"
if errorlevel 1 (
    echo.
    echo [ERROR] bmcan interface not registered.
    echo   Make sure PyPI python-can is not installed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   [OK] Setup complete! Run run.bat to start.
echo ============================================================
pause
