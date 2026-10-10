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
macOS 环境模拟自检
==================
在**任意平台**上模拟 macOS，跑一遍全部检验套件，找出「只在 macOS 上才会失败」
的断言。

★ 为什么需要它
-------------
真机 CI 每次只暴露第一个失败的步骤。之前加上 macOS 任务后，CI 反复失败，
每轮只看到一个套件的问题，来回好几趟才修完：

  1. 验收检验（字体断言写死 Windows 字体名）
  2. 保真度检验（自定义预设断言写死「楷体」）

而这个模拟脚本一次就能把这些问题全部列出来，不必等真机 CI 一轮一轮地跑。

它同时也能反过来发现问题：模拟脚本里用到的那层「平台差异」如果和真机行为
不一致，说明平台抽象的边界有问题。

原理
----
把 sys.platform 改成 "darwin"、让 platform_fonts.is_macos() 返回 True、
并用一组固定的 macOS 字体名替换 installed_fonts()，
然后在子进程里逐个跑套件。

用法: python selftest_macos_sim.py
"""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 保证中文输出在任何终端都不崩。
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))

# 要模拟的套件（跳过本模拟脚本自己）
SUITES = [
    "验收检验.py",
    "selftest_adversarial.py",
    "selftest_fidelity.py",
    "selftest_synthetic.py",
    "selftest_platform.py",
    "selftest_fontparse.py",
    # macOS 真机诊断脚本本身也要能模拟 —— 它只在 macOS 上执行，
    # 上一轮 CI 就是因为它对「_items 为空」的输入处理不当而失败。
    "selftest_macos_live.py",
]

# 真实的 macOS 家族名（依据 Apple 官方「macOS 内置字体」清单）
MACOS_FAMILY = [
    "songti sc", "heiti sc", "kaiti sc", "pingfang sc",
    "stsong", "stheiti", "stfangsong", "hiragino sans gb",
    "arial unicode ms", "times new roman", "arial",
    "helvetica neue", "courier new", "lisong pro", "lantinghei sc",
]

# 受限环境下退回扫描字体目录时只能拿到**文件名**。
# 注意它们与家族名并不一致（Songti.ttc -> songti），
# 这正是字体解析出错的根源，也是必须单独模拟这一模式的原因。
MACOS_DIRSCAN = [
    "songti", "heiti", "kaiti", "pingfang", "stsong", "stheiti",
    "hiragino sans gb", "arial unicode ms", "times new roman", "arial",
    "helvetica neue", "courier new", "lisong pro", "lantinghei sc",
    "stix twotext",
]

RUNNER_TEMPLATE = '''# -*- coding: utf-8 -*-
"""临时运行器：模拟 macOS 后执行指定脚本。由 selftest_macos_sim.py 生成。"""
import os, sys, runpy

# 注意：本文件在临时目录里，不能用 __file__ 推项目路径，由生成方写死
HERE = %(here)r
sys.path.insert(0, HERE)
import platform_fonts as pf

# 三个层面都要改，缺一个就会测不准：
#   1) sys.platform —— 套件里有 sys.platform == "darwin" 的分支
#   2) pf.is_macos() —— 库内部判断平台
#   3) pf.installed_fonts() —— 决定字体替换能否发生
#      真机 runner 上 system_profiler 往往失败、只能扫到文件名，
#      所以两种情形都要模拟（见下面的 MODE）。
sys.platform = "darwin"
pf.is_macos = lambda: True

MODE = %(mode)r
_FAMILY = %(family)r      # system_profiler 给出的真实家族名
_DIRSCAN = %(dirscan)r    # 扫描字体目录得到的文件名

if MODE == "profiler":
    _HAVE = set(_FAMILY)
else:
    _HAVE = set(_DIRSCAN)

def _fake_fonts():
    return set(_HAVE)

# 必须保留 cache_clear：有套件会调用它
_fake_fonts.cache_clear = lambda: None
pf.installed_fonts = _fake_fonts

# ---- 让 selftest_macos_live.py 也能在这里跑 ----
# 它只在 macOS 上执行，且会直接调用 system_profiler，
# 所以把 subprocess.run 也按模式伪造一份。
import subprocess as _sp
_real_run = _sp.run


class _FakeProc:
    def __init__(self, rc, out, err=""):
        self.returncode, self.stdout, self.stderr = rc, out, err


def _fake_run(cmd, *a, **k):
    try:
        argv = list(cmd)
    except TypeError:
        return _real_run(cmd, *a, **k)
    if argv and argv[0] == "system_profiler":
        return _FakeProc(0, %(sp_json)r)
    if argv and argv[0] == "sw_vers":
        return _FakeProc(0, "ProductName:\\t\\tmacOS\\nProductVersion:\\t\\t14.0\\n")
    return _real_run(cmd, *a, **k)


_sp.run = _fake_run

target = sys.argv[1]
sys.argv = [os.path.join(HERE, target)]
runpy.run_path(os.path.join(HERE, target), run_name="__main__")
'''

# 伪造的 system_profiler 输出：对应两种真实情形
SP_JSON_PROFILER = (
    '{"SPFontsDataType": [{"_name": "字体", "_items": ['
    '{"_name": "Songti SC", "path": "/System/Library/Fonts/Supplemental/Songti.ttc",'
    ' "type": "TrueType", "enabled": "yes", "typefaces": ['
    '{"_name": "Songti SC Regular", "family": "Songti SC",'
    ' "fullname": "Songti SC Regular", "style": "Regular"}]}]}]}'
)
# ★ 这一种才是真机 runner 上的实际形态：JSON 结构合法，但 _items 为空。
#   上一轮 CI 就是因为诊断脚本没考虑这种输入而失败。
SP_JSON_EMPTY = '{"SPFontsDataType": [{"_name": "字体", "_items": []}]}'

SP_JSON = {"profiler": SP_JSON_PROFILER, "dirscan": SP_JSON_EMPTY}


MODE_LABEL = {
    "profiler": "system_profiler 正常（拿到家族名）",
    "dirscan": "system_profiler 失败、退回扫描目录（只有文件名）",
}


def main() -> int:
    tmpd = tempfile.mkdtemp(prefix="macsim_")
    runners = {}
    for mode in ("profiler", "dirscan"):
        p = os.path.join(tmpd, "_runner_%s.py" % mode)
        with open(p, "w", encoding="utf-8") as f:
            f.write(RUNNER_TEMPLATE % {
                "here": HERE, "mode": mode,
                "family": MACOS_FAMILY, "dirscan": MACOS_DIRSCAN,
                "sp_json": SP_JSON[mode],
            })
        runners[mode] = p

    print("=" * 76)
    print("macOS 环境模拟自检")
    print("=" * 76)
    print("在 %s 上模拟 macOS，跑 %d 个套件 × %d 种环境"
          % (sys.platform, len(SUITES), len(runners)))
    print("  两种环境都要过：真机 runner 上 system_profiler 通常是失败的那种。")
    print()

    bad: list[tuple[str, str]] = []
    for mode, runner in runners.items():
        print("--- 环境: %s" % MODE_LABEL[mode])
        for s in SUITES:
            if not os.path.exists(os.path.join(HERE, s)):
                print("  [跳过] %s（文件不存在）" % s)
                continue
            r = subprocess.run([sys.executable, runner, s], cwd=HERE,
                               capture_output=True, text=True,
                               encoding="utf-8", errors="replace")
            if r.returncode == 0:
                print("  [通过] %s" % s)
                continue
            print("  [失败] %s" % s)
            lines = (r.stdout or "").splitlines()
            shown = 0
            for i, ln in enumerate(lines):
                if "失败明细" in ln:
                    for x in lines[i + 1:i + 14]:
                        if x.strip().startswith("✗"):
                            print("         %s" % x.strip())
                            shown += 1
            if not shown and r.stderr:
                print("         stderr: %s" % r.stderr.strip()[-300:])
            bad.append((mode, s))
        print()

    print("=" * 76)
    if bad:
        print("结果: %d 处失败" % len(bad))
        print("      断言写死了 Windows 行为，需要按平台解析期望值：")
        for mode, s in bad:
            print("        - [%s] %s" % (mode, s))
        print("=" * 76)
        return 1
    total = len(SUITES) * len(runners)
    print("结果: 全部 %d 项（%d 套件 × %d 环境）通过" % (total, len(SUITES), len(runners)))
    print("=" * 76)
    return 0


if __name__ == "__main__":
    sys.exit(main())
