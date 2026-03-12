@echo off
chcp 65001 >nul 2>&1
title XLSM to XLSX 转换工具
color 0B
cls

echo =====================================
echo     XLSM to XLSX 转换工具
echo =====================================
echo.
echo 当前文件夹: %~dp0
echo.

:: 获取当前脚本所在目录（带引号处理空格）
set "ScriptDir=%~dp0"
set "VbsScript=%ScriptDir%Convert-XlsmToXlsx.vbs"

echo 检查脚本文件...
echo VBS路径: %VbsScript%
echo.

:: 检查 VBS 脚本是否存在
if not exist "%VbsScript%" (
    echo [错误] 找不到转换脚本: Convert-XlsmToXlsx.vbs
    echo.
    echo 请确保以下两个文件在同一文件夹下：
    echo   - Convert-XlsmToXlsx.bat
    echo   - Convert-XlsmToXlsx.vbs
    echo.
    pause
    exit /b 1
)

echo [OK] VBS脚本存在
echo.
echo 正在启动转换程序...
echo =====================================
echo.

:: 运行 VBS 脚本并显示输出
cscript //nologo "%~dp0Convert-XlsmToXlsx.vbs"

echo.
echo =====================================
echo 程序执行完毕
echo.
pause
