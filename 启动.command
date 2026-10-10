#!/bin/bash
# 论文格式改写器 —— macOS / Linux 启动脚本
# 双击本文件即可（macOS 上如果提示没有执行权限，先在终端执行：
#     chmod +x 启动.command
# ）

cd "$(dirname "$0")" || exit 1

echo "=========================================="
echo " 论文格式改写器"
echo "=========================================="

# 找一个可用的 python3
PY=""
for cand in python3 python3.13 python3.12 python3.11 python3.10 python; do
    if command -v "$cand" >/dev/null 2>&1; then
        # 确认版本 >= 3.10
        if "$cand" -c 'import sys; sys.exit(0 if sys.version_info>=(3,10) else 1)' 2>/dev/null; then
            PY="$cand"
            break
        fi
    fi
done

if [ -z "$PY" ]; then
    echo
    echo "[错误] 没找到 Python 3.10 或更高版本。"
    echo
    echo "macOS 上安装方法（任选一种）："
    echo "  1) 打开「终端」，执行：  xcode-select --install"
    echo "  2) 或到 https://www.python.org/downloads/ 下载安装"
    echo
    echo "装好后重新双击本文件即可。"
    echo
    read -r -p "按回车键关闭…" _
    exit 1
fi

echo "使用 Python: $("$PY" --version 2>&1)"
echo

"$PY" 启动.py
RC=$?

if [ $RC -ne 0 ]; then
    echo
    echo "程序退出，代码 $RC"
    read -r -p "按回车键关闭…" _
fi
