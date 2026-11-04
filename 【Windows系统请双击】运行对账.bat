@echo off
REM 基金账单退款对账 - Windows 一键运行
REM 双击此文件即可自动扫描 Excel、生成 HTML 报告并打开
cd /d "%~dp0"

echo ==========================================
echo 基金账单退款对账工具 (Windows)
echo ==========================================
echo.

REM 检查 Python：优先用 python，找不到则尝试 py 启动器
set "PY=python"
python --version >nul 2>nul
if errorlevel 1 set "PY=py -3"
%PY% --version >nul 2>nul
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3
    echo 下载地址：https://www.python.org/downloads/
    echo 安装时记得勾选 "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

REM 检查依赖，缺失则自动安装
%PY% -c "import pandas, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo [提示] 首次运行，正在安装依赖...
    %PY% -m pip install pandas openpyxl -q
)

echo [开始] 正在处理，稍后自动生成报告...
echo.
%PY% process.py
echo.
pause
