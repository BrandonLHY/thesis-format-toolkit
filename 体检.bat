@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
set "DIR=%~dp0"
set "SRC=%~1"
set "SPEC=%DIR%specs\sysu_undergraduate_thesis.json"

rem ============================================================
rem  论文格式体检 —— 只检查，不修改文件
rem
rem  用法：
rem    1) 把论文 .docx 文件拖到本文件图标上松手
rem    2) 或双击本文件，再按提示把文件拖进窗口
rem
rem  输出：体检报告 HTML，用浏览器打开
rem ============================================================

echo ============================================
echo  论文格式体检（只读，不修改文件）
echo ============================================
echo.

set "PY="
where py >nul 2>nul && (
    py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
)
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo [错误] 没有找到 Python。
    echo 请先安装 Python 3.10 或更高版本：https://www.python.org/downloads/
    echo 安装时务必勾选 "Add Python to PATH"。
    echo.
    pause
    exit /b 1
)

if "%SRC%"=="" (
    echo 请把论文 .docx 文件拖进本窗口，然后按回车：
    set /p "SRC="
    set "SRC=!SRC:"=!"
)
if not exist "%SRC%" (
    echo.
    echo [错误] 找不到文件：%SRC%
    echo.
    pause
    exit /b 1
)

for %%F in ("%SRC%") do set "BASE=%%~nF"
set "OUT=%DIR%体检报告_%BASE%.html"

if not exist "%SPEC%" (
    echo [错误] 找不到规范文件：%SPEC%
    echo.
    pause
    exit /b 1
)

echo.
echo 正在体检……
echo   受检文件: %SRC%
echo   规范    : 中山大学本科生毕业论文格式要求
echo   报告    : %OUT%
echo.

%PY% "%DIR%report_html.py" "%SRC%" "%SPEC%" "%OUT%"

echo.
if exist "%OUT%" (
    echo 报告已生成，正在打开…
    start "" "%OUT%"
) else (
    echo [失败] 没有生成报告，请把上面的错误信息发回来。
)
echo.
pause
