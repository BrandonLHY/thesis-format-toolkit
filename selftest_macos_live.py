# -*- coding: utf-8 -*-
"""macOS 真机诊断（只在 macOS 上跑，尤其是 CI 的 macos runner）。

作用
----
1. 打印真实的 system_profiler 输出结构摘要
   —— 用来交叉验证 tests/fixtures/fonts_macos_real.json 里的录制夹具
      是否和真机结构一致（夹具结构错的话，Windows 上的解析测试就是自欺欺人）
2. 对比「真机探测到的字体」与「夹具里的字体」，检查关键字体是否都在
3. 验证关键字体名在真机上确实可解析

非 macOS 平台会直接跳过并以 0 退出，不影响其他平台的 CI。

用法: python selftest_macos_live.py
产物: macOS诊断报告.txt（CI 会作为 artifact 上传）
"""
from __future__ import annotations

import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
REPORT = os.path.join(HERE, "macOS诊断报告.txt")
FIXTURE = os.path.join(HERE, "tests", "fixtures", "fonts_macos_real.json")

PASS = 0
FAIL = 0
FAILURES: list[str] = []
LINES: list[str] = []


def emit(s: str = "") -> None:
    print(s)
    LINES.append(s)


def ok(name: str, cond: bool, detail: str = "") -> None:
    global PASS, FAIL
    if cond:
        PASS += 1
        emit("  [通过] %s" % name)
    else:
        FAIL += 1
        FAILURES.append("%s  %s" % (name, detail))
        emit("  [失败] %s  %s" % (name, detail))


def main() -> int:
    if sys.platform != "darwin":
        print("当前平台是 %s，这个诊断只在 macOS 上有意义，跳过。" % sys.platform)
        return 0

    emit("=" * 74)
    emit("macOS 真机诊断")
    emit("=" * 74)

    # ---------- 1. 环境 ----------
    emit("\n[1] 环境")
    try:
        r = subprocess.run(["sw_vers"], capture_output=True, text=True, timeout=10)
        emit(r.stdout.strip() or "(sw_vers 无输出)")
    except Exception as e:
        emit("  sw_vers 失败: %r" % e)
    emit("  Python: %s" % sys.version.split()[0])
    emit("  platform: %s %s" % (sys.platform, sys.version.split()[2]))

    # ---------- 2. 真实 system_profiler 结构 ----------
    emit("\n[2] 真实 system_profiler 结构（用于验证录制夹具）")
    raw = ""
    try:
        r = subprocess.run(["system_profiler", "SPFontsDataType", "-json"],
                           capture_output=True, text=True, timeout=120)
        raw = r.stdout or ""
        emit("  返回码: %d   输出长度: %d 字节" % (r.returncode, len(raw)))
        if r.stderr:
            emit("  stderr: %s" % r.stderr.strip()[:200])
    except Exception as e:
        emit("  调用失败: %r" % e)

    data = None
    if raw:
        try:
            data = json.loads(raw)
            emit("  顶层键: %s" % list(data.keys()))
        except Exception as e:
            emit("  !! JSON 解析失败: %r" % e)

    if isinstance(data, dict):
        arr = data.get("SPFontsDataType") or []
        emit("  SPFontsDataType 是数组，长度 %d" % len(arr))
        if arr and isinstance(arr[0], dict):
            emit("  第一项的子键: %s" % list(arr[0].keys()))
            items = arr[0].get("_items") or []
            emit("  _items 长度: %d" % len(items))
            if items and isinstance(items[0], dict):
                emit("  第一个字体的子键: %s" % list(items[0].keys()))
                tf = items[0].get("typefaces") or []
                emit("  typefaces 长度: %d" % len(tf))
                if tf and isinstance(tf[0], dict):
                    emit("  第一个 typeface 的子键: %s" % list(tf[0].keys()))
                # 检查夹具假设的关键字段是否真的存在
                for key in ("_name", "path", "type", "enabled", "typefaces"):
                    ok("字体条目有 %s 字段" % key, key in items[0],
                       "真机没这个字段，夹具假设错了")
                if tf and isinstance(tf[0], dict):
                    for key in ("_name", "family", "fullname", "style"):
                        ok("typeface 有 %s 字段" % key, key in tf[0],
                           "真机没这个字段，夹具假设错了")

    # ---------- 3. 解析结果 ----------
    emit("\n[3] 用自己的解析函数处理真机输出")
    from platform_fonts import parse_system_profiler_fonts, installed_fonts, _norm
    live_names = parse_system_profiler_fonts(raw)
    emit("  解析出 %d 个字体名" % len(live_names))
    ok("真机输出能解析出字体", len(live_names) > 0, "解析结果为空")

    inst = installed_fonts()
    emit("  installed_fonts() 返回 %d 个" % len(inst))
    ok("installed_fonts() 在 macOS 上有结果", len(inst) > 0, "为空")

    # ---------- 4. 与夹具交叉验证 ----------
    emit("\n[4] 真机字体 vs 录制夹具")
    # ★ 断言策略：只对「程序真正依赖的核心字体」严格断言。
    #   夹具是一个合理的超集，个别字体在不同 macOS 版本上可能有出入
    #   （例如某些字体随版本增删），把每个夹具字体都当成硬性要求会误报。
    #   核心字体缺失才是真问题——那会让候选链全部落空。
    CORE = {
        "songti sc", "heiti sc", "pingfang sc", "kaiti sc",
        "stsong", "stheiti", "stfangsong",
        "times new roman", "arial", "helvetica neue", "courier new",
    }
    if os.path.exists(FIXTURE):
        with open(FIXTURE, encoding="utf-8") as f:
            fixture_names = parse_system_profiler_fonts(f.read())
        emit("  夹具里 %d 个字体名" % len(fixture_names))
        emit("  真机解析出 %d 个字体名" % len(live_names))

        core_missing = sorted(n for n in CORE if n not in live_names)
        if core_missing:
            ok("核心字体在真机上都存在", False,
               "缺失 %s —— 候选链会落空，属真问题" % core_missing)
        else:
            ok("核心字体在真机上都存在（%d 个）" % len(CORE), True)

        # 夹具整体可信度：只报告，不判失败
        missing = sorted(n for n in fixture_names if n not in live_names)
        if missing:
            emit("  提示：夹具里有 %d 个字体在真机上没有"
                 "（不同 macOS 版本字库有差异，不判失败）：" % len(missing))
            for m in missing[:10]:
                emit("       %s" % m)
        else:
            emit("  夹具里的字体在真机上全部存在")
        emit("  真机独有（夹具未收录）: %d 个（正常，夹具只是子集）"
             % len(live_names - fixture_names))
    else:
        ok("夹具文件存在", False, FIXTURE)

    # ---------- 5. 关键字体名可解析 ----------
    emit("\n[5] 关键中文字体名在真机上的可用性")
    from platform_fonts import FONT_ALIASES
    for role, chain in sorted(FONT_ALIASES.items()):
        if not any("\u4e00" <= ch <= "\u9fff" for ch in role):
            continue
        hit = [c for c in chain if _norm(c) in inst]
        emit("     %-12s 命中 %s" % (role, hit[:3] if hit else "无"))
        ok("角色 %s 在真机上有可用候选" % role, bool(hit), "无可用候选")

    # ---------- 6. GUI 可用性（macOS runner 上 tkinter 可能有问题）----------
    emit("\n[6] 图形界面可用性（不判失败，仅记录）")
    try:
        from environment_probe import can_create_window
        can, why = can_create_window()
        emit("  can_create_window: %s  %s" % (can, why))
        if can:
            import tkinter as tk
            emit("  Tk 版本: %s" % tk.TkVersion)
            emit("  结论: 这个 runner 上可以跑 GUI 检验")
        else:
            emit("  结论: 这个 runner 上 GUI 不可用，属于已知情况"
                 "（macOS runner 的 tkinter 常有问题），不算失败")
    except Exception as e:
        emit("  探测异常（不算失败）: %r" % e)

    # ---------- 汇总 ----------
    emit("\n" + "=" * 74)
    emit("结果：通过 %d 项，失败 %d 项" % (PASS, FAIL))
    if FAILURES:
        emit("\n失败明细：")
        for f in FAILURES:
            emit("  ✗ %s" % f)
    else:
        emit("全部通过")
    emit("=" * 74)

    try:
        with open(REPORT, "w", encoding="utf-8") as f:
            f.write("\n".join(LINES))
        print("\n报告已写入: %s" % REPORT)
    except Exception as e:
        print("写报告失败: %r" % e)

    return 1 if FAILURES else 0


if __name__ == "__main__":
    sys.exit(main())
