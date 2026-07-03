@echo off
chcp 65001 >nul
title 测试用例记录表

echo ==========================================
echo   测试用例记录表  v1.1
echo ==========================================
echo.

:: 检查 Python 3
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [错误] 未检测到 Python 3，请先安装
    echo   下载地址: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [信息] Python 版本:
python --version
echo.

:: 切换到脚本所在目录
cd /d "%~dp0"

:: 安装依赖
echo [信息] 检查依赖...
python -m pip install --quiet flask openpyxl 2>nul

:: 启动主程序
python testcase_viewer.py

echo.
pause
