# -*- coding: utf-8 -*-
"""
给源码加/更新 GPL-3.0 版权头
============================
幂等，可重复运行。

★ 两次踩坑记录（都写下来避免重犯）
  1. 检测关键字用了英文 "GNU General Public License"，但版权头正文是中文
     「GNU 通用公共许可证」——已有头没被识别，重复插入变成两份。
  2. 用正则整块匹配删除旧头时写成 `while prev != text` 的循环，
     正则不匹配导致文本不变、循环永不退出，空转烧完了 CPU。

  现在改为**逐行处理**：先按行找到头部注释块的行号区间，切片删除，不用正则。

用法:
    python add_license_header.py --name "林弘毅"
"""
from __future__ import annotations

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

FILES = [
    "main.py", "app.py", "rewrite_cli.py", "rewriter.py",
    "region_detector.py", "numbering.py", "superscript.py",
    "platform_fonts.py", "audit_engine.py", "report_html.py",
    "build_exe.py", "启动.py", "打包给Mac.py",
    "selftest_synthetic.py", "selftest_platform.py",
    "verify_rewrite.py",
]

# 判定「这是版权头的一部分」的行内特征（语言无关 + 中文两种头都能命中）
HEADER_KEYS = (
    "论文格式改写器 ——",
    "Copyright (C) 2026",
    "自由软件",
    "GNU 通用公共许可证",
    "GNU General Public License",
    "任何担保",
    "默示担保",
    "gnu.org/licenses",
)


def is_header_line(line: str) -> bool:
    s = line.strip()
    if not s.startswith("#"):
        return False
    body = s.lstrip("#").strip()
    if body == "":
        return True                      # 版权头里的空注释行
    return any(k in body or k in s for k in HEADER_KEYS)


def make_header(name: str) -> list[str]:
    return [
        "# 论文格式改写器 —— 把 Word 论文按格式规范重新排版",
        "# Copyright (C) 2026 %s" % name,
        "#",
        "# 本程序是自由软件：你可以再分发和/或修改它，遵循自由软件基金会发布的",
        "# GNU 通用公共许可证条款，要么版本 3，要么（按你的选择）任何更新的版本。",
        "#",
        "# 本程序的分发是希望它有用，但不提供任何担保，甚至不提供适销性",
        "# 或特定用途适用性的默示担保。详见 GNU 通用公共许可证。",
        "#",
        "# 你应该已经随本程序收到了一份 GNU 通用公共许可证的副本。",
        "# 如果没有，见 <https://www.gnu.org/licenses/>。",
    ]


def process(path: str, name: str) -> str:
    with open(path, encoding="utf-8") as f:
        lines = f.read().splitlines(keepends=True)

    n_before = len(lines)

    # ---- 1. 保护 shebang 与编码声明 ----
    head_keep = 0
    if lines and lines[0].startswith("#!"):
        head_keep = 1
    if len(lines) > head_keep and lines[head_keep].lstrip().startswith("# -*- coding"):
        head_keep += 1

    # ---- 2. 从 head_keep 开始，吃掉连续的「头部注释块」----
    i = head_keep
    removed = 0
    while i < len(lines):
        s = lines[i]
        if s.strip() == "":
            # 空行：只有在后面还跟着版权行时才属于头部
            j = i + 1
            while j < len(lines) and lines[j].strip() == "":
                j += 1
            if j < len(lines) and is_header_line(lines[j]):
                i = j
                continue
            break
        if is_header_line(s):
            i += 1
            continue
        break
    end = i
    removed = end - head_keep

    body = lines[head_keep:end]
    rest = lines[end:]

    # ---- 3. 组装：shebang/编码 + 新头 + 正文 ----
    header = [h + "\n" for h in make_header(name)]
    new_lines = lines[:head_keep] + header + rest

    changed = (new_lines != lines)
    if changed:
        with open(path, "w", encoding="utf-8") as f:
            f.writelines(new_lines)

    if not changed:
        return "已是最新，未改动"
    if removed:
        return "已更新（清理旧头 %d 行）" % removed
    return "已添加"


def main() -> int:
    if "--name" not in sys.argv:
        print('用法: python add_license_header.py --name "你的姓名"')
        return 1
    name = sys.argv[sys.argv.index("--name") + 1]
    print("版权人: %s" % name)
    print("=" * 68)
    missing = 0
    for fn in FILES:
        p = os.path.join(HERE, fn)
        if not os.path.exists(p):
            print("  [缺失] %-22s" % fn)
            missing += 1
            continue
        print("  [%s] %-22s" % (process(p, name), fn))
    print("=" * 68)

    # ---- 自检：每个文件恰好一份版权头，且首行位置正确 ----
    bad = []
    for fn in FILES:
        p = os.path.join(HERE, fn)
        if not os.path.exists(p):
            continue
        text = open(p, encoding="utf-8").read()
        n_copy = text.count("# Copyright (C) 2026")
        n_title = text.count("# 论文格式改写器 ——")
        if n_copy != 1:
            bad.append((fn, "版权行 %d 份" % n_copy))
        if n_title != 1:
            bad.append((fn, "标题行 %d 份" % n_title))
        if "本程序作者" in text:
            bad.append((fn, "仍含占位名"))
    if bad:
        print("!! 异常: %s" % bad)
        return 1
    print("自检：%d 个文件各恰好一份版权头，无占位名残留 ✓" % (len(FILES) - missing))
    return 0 if missing == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
