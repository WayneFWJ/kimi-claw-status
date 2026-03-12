@echo off
chcp 65001 >nul
title XLSM to XLSX 转换工具
color 0B
cls

echo =====================================
echo     XLSM to XLSX 转换工具
echo =====================================
echo.
echo 使用方法：将此脚本放在包含 xlsm 文件的文件夹下，双击运行
echo.
echo 注意：
echo   - 需要安装 Microsoft Excel
echo   - 转换后的 xlsx 文件不包含宏
echo   - 原始 xlsm 文件不会被删除
echo.
echo =====================================
echo.

:: 获取当前脚本所在目录
set "ScriptDir=%~dp0"
set "VbsScript=%ScriptDir%Convert-XlsmToXlsx.vbs"

:: 检查 VBS 脚本是否存在
if not exist "%VbsScript%" (
    echo [错误] 找不到转换脚本: Convert-XlsmToXlsx.vbs
    echo.
    echo 请确保以下两个文件在同一文件夹下：
    echo   - Convert-XlsmToXlsx.bat  ^(本文件^)
    echo   - Convert-XlsmToXlsx.vbs ^(核心脚本^)
    echo.
    pause
    exit /b 1
)

:: 运行 VBS 脚本
echo 正在启动转换程序...
echo.
cscript //nologo "%VbsScript%"

:: 退出
echo.
echo 程序已结束。
pause
