@echo off
chcp 65001 >nul
setlocal

echo ============================================================
echo   dazhong2.0.py 环境一键安装
echo ============================================================
echo.

REM 脚本位于 examples\dazhong2.0\，项目根目录需向上两级
set "ROOT=%~dp0..\.."
pushd "%ROOT%"
set "ROOT=%CD%"
popd

echo [1/4] 检查 Python ...
python --version
if errorlevel 1 (
    echo ✗ 未检测到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)
echo ✓ Python 已就绪
echo.

echo [2/4] 安装本地定制版 python-can（含 bmcan 接口）...
echo       如果已装过 PyPI 官方版，会先卸载 ...
pip uninstall python-can -y
pip install -e "%ROOT%"
if errorlevel 1 (
    echo ✗ python-can 安装失败
    pause
    exit /b 1
)
echo ✓ python-can 安装完成
echo.

echo [3/4] 安装其余依赖（pandas / openpyxl / pyserial 等）...
pip install -r "%ROOT%\examples\requirements.txt"
if errorlevel 1 (
    echo ✗ 依赖安装失败
    pause
    exit /b 1
)
echo ✓ 依赖安装完成
echo.

echo [4/4] 验证 bmcan 接口 ...
python -c "import can; ok='bmcan' in can.VALID_INTERFACES; print('bmcan 可用' if ok else '✗ bmcan 不可用'); exit(0 if ok else 1)"
if errorlevel 1 (
    echo.
    echo ✗ 验证失败：bmcan 接口未注册
    echo   请检查是否误装了 PyPI 官方版 python-can
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   ✓ 环境安装完成！可以运行 dazhong2.0.py 了
echo ============================================================
pause
