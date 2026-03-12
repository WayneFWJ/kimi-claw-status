@echo off
chcp 65001 >nul
title XLSM to XLSX 转换工具
color 0B

echo =====================================
echo     XLSM to XLSX 转换工具
echo =====================================
echo.
echo 正在检查 PowerShell 执行策略...
echo.

:: 获取当前脚本所在目录
set "ScriptDir=%~dp0"
set "PsScript=%ScriptDir%Convert-XlsmToXlsx.ps1"

:: 检查 PowerShell 脚本是否存在
if not exist "%PsScript%" (
    echo [错误] 找不到 PowerShell 脚本: Convert-XlsmToXlsx.ps1
    echo.
    echo 请确保此批处理文件与 PowerShell 脚本在同一文件夹下。
    echo.
    pause
    exit /b 1
)

:: 运行 PowerShell 脚本
echo 正在启动转换程序...
echo.

powershell -ExecutionPolicy Bypass -File "%PsScript%" -FolderPath "%ScriptDir%"

:: 如果 PowerShell 脚本执行失败，显示错误
if %errorlevel% neq 0 (
    echo.
    echo [错误] 转换过程中出现错误 (代码: %errorlevel%)
    echo.
    pause
)
