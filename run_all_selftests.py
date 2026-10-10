# -*- coding: utf-8 -*-
"""
一键跑全部自检
==============
按顺序执行所有测试套件，汇总结果。CI 和本地都用它。

用法: python run_all_selftests.py
"""
from __future__ import annotations

import os
import subprocess
import sys


# 保证中文输出在任何终端都不崩。
# Windows 上 stdout 默认编码可能是 cp1252，直接 print 中文会抛
# UnicodeEncodeError 并让脚本崩掉（GitHub Actions 的 windows runner 实测踩过）。
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
HERE = os.path.dirname(os.path.abspath(__file__))

SUITES = [
    ("验收检验（端到端）", ["验收检验.py"], True),
    ("对抗性检验 A（边角与健壮性）", ["selftest_adversarial.py"], True),
    ("对抗性检验 B（保真度与元数据）", ["selftest_fidelity.py"], True),
    ("假想论文自检", ["selftest_synthetic.py"], True),
    ("跨平台字体自检", ["selftest_platform.py"], True),
]


def run(name: str, args: list[str]) -> tuple[bool, str]:
    print("\n" + "=" * 76)
    print("运行: %s" % name)
    print("=" * 76)
    r = subprocess.run([sys.executable] + args, cwd=HERE,
                       capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    tail = (r.stdout or "").strip().splitlines()[-6:]
    for line in tail:
        print("   %s" % line)
    if r.returncode != 0 and r.stderr:
        print("   stderr: %s" % r.stderr.strip()[-400:])
    return r.returncode == 0, (r.stdout or "")


def main() -> int:
    print("=" * 76)
    print("全部自检")
    print("=" * 76)
    print("Python %s / %s" % (sys.version.split()[0], sys.platform))

    results = []
    for name, args, _ in SUITES:
        okk, out = run(name, args)
        results.append((name, okk, out))
        print("   -> %s" % ("通过" if okk else "失败"))

    # 程序自检（含 GUI，Linux 无显示会失败，故标记为可选）
    print("\n" + "=" * 76)
    print("运行: 程序自检（含 GUI 构建）")
    print("=" * 76)
    r = subprocess.run([sys.executable, "main.py", "--selftest"], cwd=HERE,
                       capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    gui_ok = r.returncode == 0
    for line in (r.stdout or "").strip().splitlines()[-4:]:
        print("   %s" % line)
    if not gui_ok and sys.platform not in ("win32", "darwin"):
        print("   [提示] 非 Windows/macOS 平台无图形显示，GUI 自检失败属预期")
    results.append(("程序自检", gui_ok or sys.platform not in ("win32", "darwin"),
                    r.stdout or ""))

    print("\n" + "=" * 76)
    print("汇总")
    print("=" * 76)

    total_checks = 0
    for name, okk, out in results:
        import re
        m = None
        for pat in (r"通过 (\d+) 项", r"通过(\d+)项"):
            m = re.search(pat, out or "")
            if m:
                break
        n = int(m.group(1)) if m else 0
        total_checks += n
        print("  %-34s %s   %s 项" % (
            name, "通过" if okk else "失败", n if n else "-"))

    n_fail = sum(1 for _, okk, _ in results if not okk)
    print("\n  合计检查项: %d" % total_checks)
    print("  失败套件  : %d" % n_fail)
    print("  结论: %s" % ("全部通过" if n_fail == 0 else "有失败，需修复"))
    print("=" * 76)
    return 0 if n_fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
