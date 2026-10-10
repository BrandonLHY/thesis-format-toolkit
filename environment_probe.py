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
运行环境探测
============
用于判断当前环境能做什么，避免在无图形显示的环境（CI 服务器、纯 SSH 会话）
把「无法建窗口」误报成程序缺陷。

踩过的坑
--------
GitHub Actions 的 ubuntu-latest runner 没有 X 显示。自检里「创建 tkinter 窗口」
这一步会抛：

    TclError: no display name and no $DISPLAY environment variable

导致整套检验失败、CI 标红。这类失败不是代码问题，而是环境限制，
自检应该识别并跳过，而不是判为失败。
"""
from __future__ import annotations

import os
import sys



# 保证中文输出在任何终端都不崩。
# Windows 上 stdout 默认编码可能是 cp1252，直接 print 中文会抛
# UnicodeEncodeError 并让脚本崩掉（GitHub Actions 的 windows runner 实测踩过）。
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
def has_display() -> bool:
    """当前环境是否可能创建图形窗口。

    判定依据：
      * Windows / macOS：桌面系统默认有显示
      * 其它（Linux 等）：必须有 DISPLAY 或 WAYLAND_DISPLAY 环境变量
        （CI runner、纯 SSH 会话通常没有）
    """
    if sys.platform in ("win32", "darwin"):
        return True
    return bool(os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY"))


def can_create_window() -> tuple[bool, str]:
    """真的尝试建一次窗口，返回 (是否成功, 原因)。

    比单纯看环境变量可靠：有些环境变量存在但窗口仍建不起来。
    """
    if not has_display():
        return False, "无图形显示（缺少 $DISPLAY / $WAYLAND_DISPLAY）"
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        root.destroy()
        return True, ""
    except Exception as e:
        return False, "%s: %s" % (type(e).__name__, str(e)[:100])


if __name__ == "__main__":
    print("平台      : %s" % sys.platform)
    print("有显示    : %s" % has_display())
    okk, why = can_create_window()
    print("能建窗口  : %s  %s" % (okk, why))
