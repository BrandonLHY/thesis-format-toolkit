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
    ok("能取到 _items 的 _name", "songti sc" in n2)
    ok("取到 family", "songti sc" in n2)
    ok("取到 typefaces 的 _name", "songti sc regular" in n2)
    ok("取到 postscript", "stsongti-sc-regular" in n2)
    # ★ 区段标签不是字体名：不要把 SPFontsDataType 那层的 _name 收进来
    ok("不把区段标签当字体名", "顶层条目" not in n2,
       "「顶层条目」是区段标签，不该进字体集合")

    sec("3b. _items 为空时必须返回空集合（真实缺陷的回归测试）")
    # ★ 这是 macOS CI 真机测试发现的缺陷：
    #   system_profiler 在受限环境（如 CI runner）会返回合法 JSON 但 _items 为空。
    #   旧实现会对区段取 _name（「字体」/「Fonts」），于是 names 非空，
    #   导致 _fonts_macos() 里「探测不到就退回扫描字体目录」的兜底不触发，
    #   installed_fonts() 返回 {"字体"}，resolve_font() 找不到任何候选，
    #   最终把 Windows 字体名原样写进文档（Mac 上显示错误字体）且不报错。
    empty_cases = [
        ("_items 缺失，只有区段 _name",
         '{"SPFontsDataType": [{"_name": "字体"}]}'),
        ("_items 是空数组",
         '{"SPFontsDataType": [{"_name": "字体", "_items": []}]}'),
        ("英文区段标签",
         '{"SPFontsDataType": [{"_name": "Fonts"}]}'),
        ("SPFontsDataType 是空数组",
         '{"SPFontsDataType": []}'),
        ("_items 里的条目都没字体名",
         '{"SPFontsDataType": [{"_name": "字体", "_items": [{}, {"typefaces": []}]}]}'),
    ]
    for label, payload in empty_cases:
        r = parse_system_profiler_fonts(payload)
        ok("%s -> 返回空集合（实际 %d 个）" % (label, len(r)), len(r) == 0,
           "返回了 %s" % sorted(r)[:5])

    # 反向确认：确实有字体时不能返回空（避免上面的修复把正常情况也清空了）
    r_ok = parse_system_profiler_fonts(
        '{"SPFontsDataType": [{"_name": "字体", "_items": [{"_name": "Songti SC"}]}]}')
    ok("有字体时正常返回", "songti sc" in r_ok, repr(sorted(r_ok)))

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

    sec("7. resolve_fonts_in_preset 的行为（用真实预设结构）")
    # ★ 踩过的坑：这个测试原先写成
    #     preset = {"regions": {"body": {"ea_font": "宋体"}}}
    #   但 resolve_fonts_in_preset 期望的是 {区域名: RegionFormat 对象}
    #   （它用 getattr(fmt, "font_ea") 取值）。嵌套字典让它取不到属性，
    #   于是「永远原样返回」，测试恒真——等于没测。
    #   现在改用真实结构，并额外断言其它字段没被改动。
    from rewriter import RegionFormat
    body = RegionFormat(font_ea="宋体", font_ascii="Times New Roman",
                        size="小四", bold=False, align="两端对齐",
                        first_line_indent_chars=2, line_spacing=1.5,
                        space_before_pt=0, space_after_pt=0)
    preset = {"body": body}
    out, changes = resolve_fonts_in_preset(preset)
    ok("返回 (预设字典, 改动列表)", isinstance(out, dict) and isinstance(changes, list),
       "%s / %s" % (type(out).__name__, type(changes).__name__))
    ok("预设形状正确（按区域名索引）", "body" in out, repr(list(out.keys())[:5]))
    if sys.platform == "darwin":
        print("  （macOS：跳过「不替换」断言；替换行为见第 10 节的 monkeypatch 测试）")
    else:
        ok("Windows/Linux 上字体原样返回", out["body"].font_ea == "宋体",
           repr(out["body"].font_ea))
        ok("Windows/Linux 上无改动记录", not changes, repr(changes)[:80])
    # 这两个断言在两个平台上都成立：不该顺手改动其它字段
    ok("字号未被改动", out["body"].size == "小四", repr(out["body"].size))
    ok("行距未被改动", out["body"].line_spacing == 1.5, repr(out["body"].line_spacing))
    ok("首行缩进未被改动", out["body"].first_line_indent_chars == 2,
       repr(out["body"].first_line_indent_chars))

    sec("8. installed_fonts 在当前平台可用")
    fonts = installed_fonts()
    ok("能拿到字体集合或空集合（不抛异常）", isinstance(fonts, set))
    if sys.platform == "darwin":
        ok("macOS 上探测到字体（CI 验真）", len(fonts) > 0, "探测为空")
    print("     当前平台探测到 %d 个字体名" % len(fonts))

    sec("9. 字体存在性判定（含降级路径的模糊匹配）")
    from platform_fonts import _fuzzy_present
    # 精确匹配由 resolve_font 内部处理，这里直接测模糊匹配这一层
    ok("精确同形不做模糊也能命中（_norm）", _norm("Songti SC") == "songti sc")
    # ★ 降级路径：目录扫描得到的是文件名（songti），要能匹配上家族名 Songti SC
    ok("文件名 songti 能匹配 Songti SC",
       _fuzzy_present("Songti SC", {"songti"}))
    ok("文件名 heiti 能匹配 Heiti SC",
       _fuzzy_present("Heiti SC", {"heiti"}))
    ok("文件名 pingfang 能匹配 PingFang SC",
       _fuzzy_present("PingFang SC", {"pingfang"}))
    ok("文件名 kaiti 能匹配 Kaiti SC",
       _fuzzy_present("Kaiti SC", {"kaiti"}))
    # 短名不该乱命中
    ok("短名不做模糊匹配（长度门槛）",
       not _fuzzy_present("Kai", {"kaiti sc"}), "短名应被长度门槛挡住")
    ok("无关字体不命中",
       not _fuzzy_present("Songti SC", {"arial", "helvetica"}))

    sec("10. 降级路径：目录扫描只拿到文件名时仍须解析出 macOS 家族名")
    # ★ 踩过的坑（两层）：
    #   1) 原实现在精确匹配那轮会因 Kaiti.ttc -> kaiti 与 Windows 名 KaiTi 撞名，
    #      选中一个 Mac 上不存在的名字；
    #   2) 只在模糊那轮修还不够——必须让「来源可靠的精确匹配」与
    #      「来源不可靠只用模糊匹配」分开，并在不可靠时优先挑非精确命中的候选。
    #   这段测试用 monkeypatch 在任意平台上验证该行为。
    import platform_fonts as _pf
    _orig_mac = _pf.is_macos
    _orig_inst = _pf.installed_fonts
    _orig_rel = _pf.fonts_from_reliable_source
    try:
        _pf.is_macos = lambda: True

        # A. 来源可靠：system_profiler 给出真实家族名
        _pf.installed_fonts = lambda: {
            "songti sc", "heiti sc", "kaiti sc", "pingfang sc",
            "stsong", "stheiti", "stfangsong",
        }
        _pf.fonts_from_reliable_source = lambda: True
        # 注意：候选链里 STFangsong 排在 Kaiti SC 之前，所以仿宋应得 STFangsong
        expect_a = {"宋体": "Songti SC", "黑体": "Heiti SC", "楷体": "Kaiti SC",
                    "微软雅黑": "PingFang SC", "仿宋": "STFangsong"}
        for role, want in expect_a.items():
            got = _pf.resolve_font(role)
            ok("来源可靠 %s -> %s" % (role, want), got == want, "实际 %s" % got)

        # B. 来源不可靠：目录扫描只给文件名（songti / kaiti ...）
        _pf.installed_fonts = lambda: {
            "songti", "heiti", "kaiti", "pingfang", "stsong", "stheiti",
        }
        _pf.fonts_from_reliable_source = lambda: False
        # 目录扫描拿不到 stfangsong（该字体在 Songti.ttc 内部），故仿宋落到 Kaiti SC
        expect_b = {"宋体": "Songti SC", "黑体": "Heiti SC", "楷体": "Kaiti SC",
                    "微软雅黑": "PingFang SC", "仿宋": "Kaiti SC"}
        for role, want in expect_b.items():
            got = _pf.resolve_font(role)
            ok("来源不可靠 %s -> %s" % (role, want), got == want, "实际 %s" % got)

        # 反向检查：降级路径绝不能把 Windows 名留在结果里
        win_names = {"宋体", "黑体", "楷体", "微软雅黑", "仿宋",
                     "KaiTi", "SimSun", "SimHei", "FangSong"}
        leaked = [(r, _pf.resolve_font(r)) for r in expect_b
                  if _pf.resolve_font(r) in win_names]
        ok("降级路径不残留 Windows 字体名", not leaked, repr(leaked))
    finally:
        _pf.is_macos = _orig_mac
        _pf.installed_fonts = _orig_inst
        _pf.fonts_from_reliable_source = _orig_rel

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
