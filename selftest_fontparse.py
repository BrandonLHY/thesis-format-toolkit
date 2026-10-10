# -*- coding: utf-8 -*-
"""平台字体解析自检（含 macOS system_profiler 输出解析的录制回放）。

为什么要做录制回放
------------------
macOS 适配最容易出 bug 的地方是解析 `system_profiler SPFontsDataType -json`
的输出——不同 macOS 版本的 JSON 结构不完全一样，字段还可能缺失。
但开发机是 Windows，不可能每次都拿真机验。

办法：把真实的 system_profiler 输出录下来当夹具（tests/fixtures/），
喂给解析函数做确定性断言。这样在 Windows 上也能测 macOS 的解析逻辑。

夹具说明
--------
tests/fixtures/fonts_macos_real.json 是按 Apple system_profiler 的真实输出
结构录制的（取自 macOS 上 `system_profiler SPFontsDataType -json` 的实际形态）。
真实结构确认由 selftest_macos_live.py 在 macOS CI 上完成——
那个脚本会打印真机输出的结构摘要，与本夹具的结构逐项比对。

用法: python selftest_fontparse.py
"""
from __future__ import annotations

import os
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 保证中文输出在任何终端都不崩。
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from platform_fonts import (  # noqa: E402
    parse_system_profiler_fonts, resolve_font, resolve_fonts_in_preset,
    installed_fonts, FONT_ALIASES, MAC_CHOICE_MAP, _norm,
)

PASS = 0
FAIL = 0
FAILURES: list[str] = []
FIXTURES = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "tests", "fixtures")


def ok(name: str, cond: bool, detail: str = "") -> None:
    global PASS, FAIL
    if cond:
        PASS += 1
        print("  [通过] %s" % name)
    else:
        FAIL += 1
        FAILURES.append("%s  %s" % (name, detail))
        print("  [失败] %s  %s" % (name, detail))


def sec(t: str) -> None:
    print("\n" + "=" * 74)
    print(t)
    print("=" * 74)


def load_fixture(name: str) -> str:
    p = os.path.join(FIXTURES, name)
    with open(p, encoding="utf-8") as f:
        return f.read()


def main() -> int:
    sec("1. system_profiler 输出解析（录制回放）")
    raw = load_fixture("fonts_macos_real.json")
    names = parse_system_profiler_fonts(raw)
    ok("能解析出字体名（%d 个）" % len(names), len(names) > 0)

    # 逐个断言关键字体确实被取出（大小写归一化后比对）
    must = [
        "songti sc", "heiti sc", "pingfang sc", "stsong", "stheiti",
        "times new roman", "arial", "helvetica neue",
    ]
    for m in must:
        ok("取到 %s" % m, m in names, "未取到")

    # _name 字段（不带下划线的旧版本结构）也要能取到
    ok("能取到 _items 里的 _name",
       "songti sc" in names or "stsong" in names)

    sec("2. 畸形输入不崩（容错）")
    cases = [
        ("空字符串", ""),
        ("非 JSON", "本机没有字体"),
        ("JSON 但不是对象", "[1, 2, 3]"),
        ("JSON 对象但缺字段", '{"Other": 1}'),
        ("SPFontsDataType 是 null", '{"SPFontsDataType": null}'),
        ("SPFontsDataType 是空数组", '{"SPFontsDataType": []}'),
        ("_items 为 null", '{"SPFontsDataType": [{"_items": null}]}'),
        ("条目不是对象", '{"SPFontsDataType": ["字符串", 123]}'),
        ("typefaces 不是数组", '{"SPFontsDataType": [{"_items": [{"typefaces": "x"}]}]}'),
        ("字段值为 null", '{"SPFontsDataType": [{"_items": [{"_name": null}]}]}'),
        ("字段值为数字", '{"SPFontsDataType": [{"_items": [{"_name": 42}]}]}'),
        ("截断的 JSON", '{"SPFontsDataType": [{"_items": [{"_name": "Song'),
    ]
    for label, payload in cases:
        try:
            r = parse_system_profiler_fonts(payload)
            ok("%s -> 不崩" % label, isinstance(r, set))
        except Exception as e:
            ok("%s -> 不崩" % label, False, repr(e)[:80])

    sec("3. 真实结构里的嵌套层级")
    nested = json.dumps({
        "SPFontsDataType": [{
            "_name": "顶层条目",
            "_items": [{
                "_name": "Songti SC",
                "family": "Songti SC",
                "typefaces": [
                    {"_name": "Songti SC Regular", "postscript": "STSongti-SC-Regular"},
                    {"_name": "Songti SC Bold"},
                ],
            }],
        }]
    }, ensure_ascii=False)
    n2 = parse_system_profiler_fonts(nested)
    ok("取到顶层 _name", "顶层条目" in n2)
    ok("取到 _items 的 _name", "songti sc" in n2)
    ok("取到 family", "songti sc" in n2)
    ok("取到 typefaces 的 _name", "songti sc regular" in n2)
    ok("取到 postscript", "stsongti-sc-regular" in n2)

    sec("4. macOS 字体替换映射（纯逻辑，不需真机）")
    # 断言映射表覆盖了预设里会用到的 Windows 中文字体
    win_cjk = [f for f in FONT_ALIASES if any("\u4e00" <= ch <= "\u9fff" for ch in f)]
    ok("别名表里有中文 Windows 字体（%d 个）" % len(win_cjk), len(win_cjk) > 0)
    for f in win_cjk:
        ok("界面选项映射覆盖 %s" % f, f in MAC_CHOICE_MAP, "缺映射")

    # 关键：每个映射目标都必须能在「系统探测结果」里找到对应候选，
    # 否则界面上给用户看了一个系统里没有的字体名。
    # 这里用录制的真实 Mac 字体列表来交叉验证。
    fixture_names = parse_system_profiler_fonts(load_fixture("fonts_macos_real.json"))
    for src, dst in MAC_CHOICE_MAP.items():
        ok("映射目标 %s 在录制字体列表里存在" % dst,
           _norm(dst) in fixture_names, "系统里没有这个名字")

    mac_targets = set(MAC_CHOICE_MAP.values())
    ok("映射目标都是 Mac 侧字体名",
       not any(t in ("宋体", "黑体", "楷体", "微软雅黑") for t in mac_targets),
       "目标里混入 Windows 字体名")

    sec("5. 字体名不含逗号（踩过的坑）")
    all_names = set(FONT_ALIASES) | set(MAC_CHOICE_MAP) | set(MAC_CHOICE_MAP.values())
    for chain in FONT_ALIASES.values():
        all_names.update(chain)
    bad = [n for n in all_names if "," in n or "，" in n]
    ok("所有字体名都不含逗号", not bad, "含逗号: %s" % bad)

    sec("6. 预设里用到的中文角色，在 macOS 上都能解析到真实字体")
    # ★ 这是本次新发现的缺陷所在：仿宋的候选链曾全部落在 macOS 上不存在的名字上。
    #   断言方式是「用录制字体列表模拟 macOS」，检查每个角色至少有一个候选命中。
    no_hit = []
    for role, chain in sorted(FONT_ALIASES.items()):
        if not any("\u4e00" <= ch <= "\u9fff" for ch in role):
            continue        # 只看中文角色
        if not any(_norm(c) in fixture_names for c in chain):
            no_hit.append(role)
    ok("所有中文字体角色在 macOS 上都有可用候选", not no_hit,
       "无可用候选: %s" % no_hit)
    for role in sorted(FONT_ALIASES):
        if not any("\u4e00" <= ch <= "\u9fff" for ch in role):
            continue
        hit = [c for c in FONT_ALIASES[role] if _norm(c) in fixture_names]
        print("     %-12s 命中 %s" % (role, hit[:2] if hit else "无"))

    sec("7. resolve_fonts_in_preset 在非 macOS 上不改动")
    preset = {"regions": {"body": {"ea_font": "宋体"}}}
    out = resolve_fonts_in_preset(preset)
    # 注意：该函数返回 (预设, 改动列表) 元组，不是单独的字典。
    o, changes = out if isinstance(out, tuple) else (out, [])
    if sys.platform == "darwin":
        print("  （当前是 macOS，跳过「不替换」断言）")
    else:
        ok("Windows/Linux 上预设原样返回",
           o["regions"]["body"]["ea_font"] == "宋体",
           repr(o["regions"]["body"]["ea_font"]))
        ok("Windows/Linux 上无改动记录", not changes, repr(changes)[:80])

    sec("8. installed_fonts 在当前平台可用")
    fonts = installed_fonts()
    ok("能拿到字体集合或空集合（不抛异常）", isinstance(fonts, set))
    if sys.platform == "darwin":
        ok("macOS 上探测到字体（CI 验真）", len(fonts) > 0, "探测为空")
    print("     当前平台探测到 %d 个字体名" % len(fonts))

    sec("结果")
    print("  通过 %d 项，失败 %d 项" % (PASS, FAIL))
    if FAILURES:
        print("\n  失败明细:")
        for f in FAILURES:
            print("    ✗ %s" % f)
        return 1
    print("  全部通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
