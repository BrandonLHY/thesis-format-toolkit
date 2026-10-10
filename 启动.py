# -*- coding: utf-8 -*-
# 论文格式改写器 —— 把 Word 论文按格式规范重新排版
# Copyright (C) 2026 林弘毅
#
# 本程序是自由软件：你可以再分发和/或修改它，遵循自由软件基金会发布的
# GNU 通用公共许可证条款，要么版本 3，要么（按你的选择）任何更新的版本。
#
# 本程序的分发是希望它有用，但不提供任何担保，甚至不提供适销性
# 或特定用途适用性的默示担保。详见 GNU 通用公共许可证。
#
# 你应该已经随本程序收到了一份 GNU 通用公共许可证的副本。
# 如果没有，见 <https://www.gnu.org/licenses/>。
"""
启动脚本（跨平台）
==================
给拿到源码的人用：自动检查并安装依赖，然后启动程序。

用法：
    python 启动.py            正常启动图形界面
    python 启动.py --check    只检查环境，不启动
    python 启动.py --cli 论文.docx    不开界面，直接改写（命令行模式）

为什么需要它：手工跑 `pip install python-docx lxml` 对非技术用户是门槛。
本脚本把这一步自动化，遇到权限问题会给出明确的手动指令而不是一堆报错。
"""
from __future__ import annotations

import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REQUIRED = [("docx", "python-docx"), ("lxml", "lxml")]


def log(msg: str = ""):
    print(msg, flush=True)


def check_deps() -> list[str]:
    """返回缺失的包名（pip 名）。"""
    missing = []
    for mod, pipname in REQUIRED:
        try:
            __import__(mod)
        except ImportError:
            missing.append(pipname)
    return missing


def try_install(pkgs: list[str]) -> bool:
    """尝试安装。先普通装，失败再试 --user，最后给出手动指令。"""
    base = [sys.executable, "-m", "pip", "install"]
    attempts = [
        (base + pkgs, "普通安装"),
        (base + ["--user"] + pkgs, "安装到用户目录"),
    ]
    for cmd, label in attempts:
        log("  正在%s：%s" % (label, " ".join(pkgs)))
        try:
            r = subprocess.run(cmd, timeout=600)
            if r.returncode == 0:
                log("  安装成功（%s）" % label)
                return True
            log("  %s 失败（退出码 %d）" % (label, r.returncode))
        except Exception as e:
            log("  %s 出错：%r" % (label, e))
    return False


def manual_hint(pkgs: list[str]) -> str:
    joined = " ".join(pkgs)
    return (
        "\n请手动执行下面这行命令，然后再运行本脚本：\n\n"
        "    %s -m pip install --user %s\n\n"
        "如果提示找不到 pip，请先安装 Python 3：https://www.python.org/downloads/\n"
        % (sys.executable, joined)
    )


def main() -> int:
    log("=" * 58)
    log("论文格式改写器 —— 启动检查")
    log("=" * 58)
    log("Python: %s" % sys.version.split()[0])
    log("平台  : %s" % sys.platform)
    log("目录  : %s" % HERE)

    # Python 版本底线：代码用了 3.10+ 的语法（X | None）
    if sys.version_info < (3, 10):
        log("\n[错误] 需要 Python 3.10 或更高版本，当前是 %s。" % sys.version.split()[0])
        log("请到 https://www.python.org/downloads/ 安装新版后重试。")
        return 2

    missing = check_deps()
    if missing:
        log("\n缺少依赖：%s" % ", ".join(missing))
        log("正在尝试自动安装…")
        if not try_install(missing):
            log(manual_hint(missing))
            return 3
        missing2 = check_deps()
        if missing2:
            log("\n[错误] 安装后仍然缺少：%s" % ", ".join(missing2))
            log(manual_hint(missing2))
            return 3
    log("\n依赖检查: 通过")

    # 字体探测（仅 macOS 需要替换字体）
    try:
        from platform_fonts import installed_fonts, is_macos, resolve_font
        if is_macos():
            have = installed_fonts()
            log("系统字体探测: 找到 %d 个" % len(have))
            if not have:
                log("  [提示] 字体探测失败，程序会沿用原始字体名，")
                log("         若成稿里中文显示为默认字体，请手动在界面里选 Mac 字体。")
            else:
                for probe in ("宋体", "黑体", "楷体"):
                    log("  %s -> %s" % (probe, resolve_font(probe)))
        else:
            log("系统字体探测: 本平台无需替换字体")
    except Exception as e:
        log("字体探测跳过：%r" % e)

    if "--check" in sys.argv:
        log("\n环境检查完毕（--check 模式，不启动界面）")
        return 0

    log("\n正在启动…")
    try:
        import main as entry
        if "--cli" in sys.argv:
            i = sys.argv.index("--cli")
            sys.argv = [sys.argv[0]] + sys.argv[i + 1:]
            return entry.cli_main() if hasattr(entry, "cli_main") else 1
        import tkinter as tk
        import app as appmod
        root = tk.Tk()
        appmod.App(root)
        root.mainloop()
        return 0
    except Exception:
        import traceback
        log("\n启动失败：")
        log(traceback.format_exc())
        log("\n请把上面的错误信息发回来。")
        return 1


if __name__ == "__main__":
    sys.exit(main())
