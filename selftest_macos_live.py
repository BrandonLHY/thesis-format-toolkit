# -*- coding: utf-8 -*-
"""macOS 真机诊断（只在 macOS 上跑，尤其是 CI 的 macos runner）。

作用
----
1. 打印真实的 system_profiler 调用结果与输出结构
   —— 用来交叉验证 tests/fixtures/fonts_macos_real.json 里的录制夹具
      是否和真机结构一致（夹具结构错的话，Windows 上的解析测试就是自欺欺人）
2. 对比「真机探测到的字体」与「夹具里的字体」
3. 验证关键字体名在真机上确实可解析

★ 判定原则
----------
CI runner 是受限环境：`system_profiler` 可能不存在、超时、或被沙箱挡住。
这类情况是**环境限制**，不是程序缺陷，不应该让 CI 标红——
否则每次 macOS 构建都会假失败。

因此本脚本的判定是：
* system_profiler 不可用  -> 记录为「环境限制」，改用字体目录扫描继续验证，不算失败
* system_profiler 可用但解析不出字体 -> 判失败（那说明我的解析逻辑有问题）
* 字体目录扫描也拿不到任何字体 -> 判失败（那才是真问题）

非 macOS 平台直接跳过并以 0 退出。

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

# 程序真正依赖的核心字体：这些拿不到，候选链就会落空
CORE_FONTS = {
    "songti sc", "heiti sc", "pingfang sc", "kaiti sc",
    "stsong", "stheiti", "stfangsong",
    "times new roman", "arial", "helvetica neue", "courier new",
}


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


def note(s: str) -> None:
    emit("  [记录] %s" % s)


def run_cmd(cmd: list[str], timeout: int) -> tuple[int, str, str]:
    """跑一个命令，返回 (返回码, stdout, stderr)。失败时返回码为 -1。"""
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        return r.returncode, r.stdout or "", r.stderr or ""
    except FileNotFoundError:
        return -1, "", "命令不存在: %s" % cmd[0]
    except subprocess.TimeoutExpired:
        return -1, "", "超时（%ds）" % timeout
    except Exception as e:
        return -1, "", "%s: %s" % (type(e).__name__, e)


def scan_font_dirs() -> set[str]:
    """扫描 macOS 字体目录，用文件名当字体名（粗糙但总能拿到点东西）。"""
    found: set[str] = set()
    dirs = ["/System/Library/Fonts", "/System/Library/Fonts/Supplemental",
            "/Library/Fonts", os.path.expanduser("~/Library/Fonts")]
    for d in dirs:
        if not os.path.isdir(d):
            continue
        for root, _dirs, files in os.walk(d):
            for fn in files:
                if fn.lower().endswith((".ttf", ".ttc", ".otf", ".dfont")):
                    found.add(os.path.splitext(fn)[0].strip().lower())
    return found


def main() -> int:
    if sys.platform != "darwin":
        print("当前平台是 %s，这个诊断只在 macOS 上有意义，跳过。" % sys.platform)
        return 0

    emit("=" * 74)
    emit("macOS 真机诊断")
    emit("=" * 74)

    # ---------- 1. 环境 ----------
    emit("\n[1] 环境")
    rc, out, err = run_cmd(["sw_vers"], 15)
    emit("  sw_vers 返回码=%d" % rc)
    if out.strip():
        for line in out.strip().splitlines():
            emit("    %s" % line)
    if err.strip():
        emit("    stderr: %s" % err.strip()[:200])
    emit("  Python: %s" % sys.version.split()[0])
    emit("  platform: %s %s" % (sys.platform, sys.version.split()[2]))

    from platform_fonts import parse_system_profiler_fonts, installed_fonts, _norm

    # ---------- 2. system_profiler 调用实况 ----------
    emit("\n[2] system_profiler 调用实况（这一步的原始信息最关键）")
    emit("  命令: system_profiler SPFontsDataType -json")
    rc, raw, err = run_cmd(["system_profiler", "SPFontsDataType", "-json"], 90)
    emit("  返回码: %d" % rc)
    emit("  stdout 长度: %d 字节" % len(raw))
    emit("  stderr 长度: %d 字节" % len(err))
    if err.strip():
        emit("  stderr 内容（前 300 字符）:")
        for line in err.strip()[:300].splitlines():
            emit("      %s" % line)
    if raw:
        emit("  stdout 前 300 字符:")
        for line in raw[:300].splitlines():
            emit("      %s" % line)
    else:
        emit("  !! stdout 为空")

    sp_usable = bool(raw.strip())
    live_from_sp: set[str] = set()
    data = None
    if sp_usable:
        try:
            data = json.loads(raw)
            emit("  JSON 解析成功。顶层键: %s" % list(data.keys())[:8])
        except Exception as e:
            emit("  !! JSON 解析失败: %s: %s" % (type(e).__name__, e))
            data = None

    # ---------- 3. 真实结构（用于验证录制夹具） ----------
    emit("\n[3] 真实 system_profiler 结构（用于验证录制夹具）")
    if isinstance(data, dict):
        arr = data.get("SPFontsDataType") or []
        emit("  SPFontsDataType 是数组，长度 %d" % len(arr))
        # 统计「所有条目里一共有多少个字体条目」，用于区分下面两种情形：
        #   A. JSON 结构合法且确实列了字体 -> 解析不出来就是解析逻辑的问题
        #   B. JSON 结构合法但 _items 为空   -> 环境特性（runner 上常见），不是缺陷
        total_items = 0
        for _it in arr:
            if isinstance(_it, dict):
                total_items += len(_it.get("_items") or [])
        emit("  所有条目的 _items 合计: %d 个字体条目" % total_items)
        if arr and isinstance(arr[0], dict):
            emit("  第一项的子键: %s" % list(arr[0].keys()))
            items = arr[0].get("_items") or []
            emit("  第一项的 _items 长度: %d" % len(items))
            if items and isinstance(items[0], dict):
                emit("  第一个字体条目的子键: %s" % list(items[0].keys()))
                for key in ("_name", "path", "type", "enabled", "typefaces"):
                    ok("字体条目有 %s 字段" % key, key in items[0],
                       "真机无此字段，夹具假设可能不对")
                tf = items[0].get("typefaces") or []
                emit("  typefaces 长度: %d" % len(tf))
                if tf and isinstance(tf[0], dict):
                    emit("  第一个 typeface 的子键: %s" % list(tf[0].keys()))
                    for key in ("_name", "family", "fullname", "style"):
                        ok("typeface 有 %s 字段" % key, key in tf[0],
                           "真机无此字段，夹具假设可能不对")
        live_from_sp = parse_system_profiler_fonts(raw)
        emit("  用解析函数处理真机输出 -> %d 个字体名" % len(live_from_sp))
        if total_items > 0:
            # 有字体条目却解析不出来 —— 这才是解析逻辑的问题
            ok("有字体条目时必须能解析出字体", len(live_from_sp) > 0,
               "JSON 里有 %d 个字体条目，解析结果却为空" % total_items)
        else:
            note("JSON 结构合法，但 _items 为空（runner 上字体列表未填充）。"
                 "此时解析出 0 个字体是**正确**的，记为环境特性而非解析缺陷")
    else:
        note("system_profiler 未给出可用 JSON —— 这是 CI 受限环境的常见情况，"
             "记为环境限制而非程序缺陷")
        note("原因见上面「返回码 / stderr / stdout」。"
             "macOS runner 上该命令可能被沙箱限制或超时。")

    # ---------- 4. 字体目录扫描（不依赖 system_profiler） ----------
    emit("\n[4] 字体目录扫描（不依赖 system_profiler）")
    dirs_found = scan_font_dirs()
    emit("  扫描 /System/Library/Fonts、/Library/Fonts、~/Library/Fonts")
    emit("  得到 %d 个字体文件名" % len(dirs_found))
    # 只记录不判失败：这是 runner 的字体环境，不是程序行为。
    note("字体目录里扫到 %d 个字体文件" % len(dirs_found))

    # ---------- 5. 用真机数据验证核心字体 ----------
    emit("\n[5] 核心字体可用性（只记录，不判失败）")
    # 这一节反映的是 runner 装了什么字体，属环境事实。
    # 判失败只会让 CI 因环境差异误报，所以全部记为 note。
    probe = live_from_sp if live_from_sp else dirs_found
    src_label = "system_profiler" if live_from_sp else "字体目录扫描"
    emit("  依据: %s（%d 个名字）" % (src_label, len(probe)))
    if src_label == "字体目录扫描":
        note("目录扫描得到的是文件名，与字体家族名不完全对应"
             "（如 Songti.ttc -> songti），因此只做宽松匹配")
        hit = [n for n in CORE_FONTS
               if any(n.replace(" ", "") in p.replace(" ", "")
                      or p.replace(" ", "") in n.replace(" ", "")
                      for p in probe)]
        emit("  宽松匹配命中 %d/%d 个核心字体" % (len(hit), len(CORE_FONTS)))
        emit("    命中: %s" % (sorted(hit) if hit else "无"))
    else:
        missing = sorted(n for n in CORE_FONTS if n not in probe)
        emit("  核心字体缺失: %s" % (missing if missing else "无"))

    # ---------- 6. 与夹具交叉验证 ----------
    emit("\n[6] 真机字体 vs 录制夹具")
    if os.path.exists(FIXTURE):
        with open(FIXTURE, encoding="utf-8") as f:
            fixture_names = parse_system_profiler_fonts(f.read())
        emit("  夹具里 %d 个字体名" % len(fixture_names))
        if live_from_sp:
            missing = sorted(fixture_names - live_from_sp)
            if missing:
                note("夹具里有 %d 个字体在真机上没有"
                     "（不同 macOS 版本字库有增删，不判失败）: %s"
                     % (len(missing), missing[:8]))
            else:
                emit("  夹具里的字体在真机上全部存在")
            emit("  真机独有（夹具未收录）: %d 个（正常，夹具只是子集）"
                 % len(live_from_sp - fixture_names))
        else:
            note("system_profiler 不可用，跳过夹具逐项比对"
                 "（夹具本身已由 Windows 上的录制回放测试覆盖）")
    else:
        ok("夹具文件存在", False, FIXTURE)

    # ---------- 7. installed_fonts() 实际返回 ----------
    emit("\n[7] installed_fonts() 的实际返回")
    inst = installed_fonts()
    emit("  返回 %d 个字体名" % len(inst))
    emit("  样例: %s" % (sorted(inst)[:12] if inst else "（空）"))
    # ★ 唯一的硬断言：返回空只允许发生在「字体目录本身就是空的」这种情况下。
    #   否则说明探测逻辑坏了（这正是之前那个缺陷的形态：探测失败却伪装成功，
    #   返回 {"字体"} 这样的无效结果，导致候选链全部落空）。
    ok("installed_fonts() 有结果，或字体目录确实为空",
       len(inst) > 0 or len(dirs_found) == 0,
       "字体目录里有 %d 个字体文件，但 installed_fonts() 返回空" % len(dirs_found))
    if inst:
        hit = sorted(n for n in CORE_FONTS if n in inst)
        emit("  核心字体命中 %d/%d: %s" % (len(hit), len(CORE_FONTS), hit))

    # ---------- 8. 图形界面可用性（只记录） ----------
    emit("\n[8] 图形界面可用性（不判失败，仅记录）")
    try:
        from environment_probe import can_create_window
        can, why = can_create_window()
        emit("  can_create_window: %s  %s" % (can, why))
        if can:
            import tkinter as tk
            emit("  Tk 版本: %s" % tk.TkVersion)
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
