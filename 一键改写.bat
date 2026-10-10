@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
set "DIR=%~dp0"
set "SRC=%~1"

rem ============================================================
rem  论文格式改写器 —— 一键改写
rem
rem  用法：
rem    1) 把论文 .docx 文件拖到本文件图标上松手
rem    2) 或双击本文件，再按提示把文件拖进窗口
rem
rem  输出：与原文件同名的「xxx_已排版.docx」，放在本工具目录下
rem  原文件不会被修改
rem ============================================================

echo ============================================
echo  论文格式改写器
echo ============================================
echo.

rem ---- 找 Python ----
rem 优先用官方启动器 py（能拿到较新的 3.x），再退回 PATH 里的 python
set "PY="
where py >nul 2>nul && (
    py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
)
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)

if not defined PY (
    echo [错误] 没有找到 Python。
    echo.
    echo 请先安装 Python 3.10 或更高版本：https://www.python.org/downloads/
    echo 安装时务必勾选 "Add Python to PATH"。
    echo.
    pause
    exit /b 1
)

rem ---- 取输入文件 ----
if "%SRC%"=="" (
    echo 请把论文 .docx 文件拖进本窗口，然后按回车：
    set /p "SRC="
    rem 拖拽进 cmd 的路径会带引号，去掉
    set "SRC=!SRC:"=!"
)

if not exist "%SRC%" (
    echo.
    echo [错误] 找不到文件：%SRC%
    echo.
    pause
    exit /b 1
)

rem ---- 输出路径 ----
for %%F in ("%SRC%") do set "BASE=%%~nF"
set "OUT=%DIR%%BASE%_已排版.docx"

echo.
echo 正在改写……
echo   输入: %SRC%
echo   输出: %OUT%
if exist "%DIR%我的预设.json" (
    echo   预设: 我的预设.json
) else if exist "%DIR%预设.json" (
    echo   预设: 预设.json
) else (
    echo   预设: 内置默认
)
echo.

%PY% "%DIR%rewrite_cli.py" "%SRC%" "%OUT%"
set "RC=%ERRORLEVEL%"

echo.
if exist "%OUT%" (
    echo 改写完成，正在打开成稿…
    start "" "%OUT%"
) else (
    echo [失败] 没有生成输出文件，退出码 %RC%
    echo 请把上面的错误信息发回来。
)
echo.
pause
