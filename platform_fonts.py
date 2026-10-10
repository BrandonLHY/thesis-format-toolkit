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
跨平台字体适配
==============
同一份预设要在 Windows 和 macOS 上都能正确显示，难点全在字体名：
Windows 有「宋体/黑体/楷体」，macOS 没有这些名字。

★ 一个被实测推翻的方案（记录下来避免重犯）
------------------------------------------
曾打算写入「逗号分隔的字体回退链」，如 w:eastAsia="宋体,Songti SC"，
以为 Word 会依次尝试。用 Word COM 实测的结果是：

    写入 "不存在的字体XYZ,宋体"  ->  Word 读回 [不存在的字体XYZ,宋体]

也就是说 Word 把整个字符串当成**一个字体名**，不会回退。
字体链方案无效，必须改成：**运行时探测系统已安装的字体，只写确实存在的那个。**
"""
from __future__ import annotations

import functools
import os
import subprocess
import sys

# 逻辑字体角色 -> 各平台候选（按优先级）
FONT_ALIASES: dict[str, list[str]] = {
    "宋体": ["宋体", "SimSun", "Songti SC", "STSong", "宋体-简", "Source Han Serif SC"],
    "黑体": ["黑体", "SimHei", "Heiti SC", "STHeiti", "黑体-简", "Source Han Sans SC"],
    "楷体": ["楷体", "KaiTi", "Kaiti SC", "STKaiti", "楷体-简"],
    # ★ 仿宋这一链曾全部落在 macOS 上不存在的名字上（仿宋/FangSong/STFangsong/仿宋-简），
    #   若学校要求正文用仿宋（公文类论文常见），在 Mac 上会写入一个不存在的字体名，
    #   Word 只能自行回退，结果不可预期。
    #   STFangsong 是否随 macOS 出厂无法确证（Office 会装，但这不该作为前提），
    #   因此补一个已确证存在的中文衬线字体作为兜底。
    #   注意：Kaiti SC 是楷体风格，与仿宋不等价，属「保证可用」的降级替代，
    #   已在 README 的 Mac 说明里写明这一点。
    "仿宋": ["仿宋", "FangSong", "STFangsong", "仿宋-简", "Kaiti SC", "STSong"],
    "仿宋_GB2312": ["仿宋_GB2312", "仿宋", "FangSong", "STFangsong", "Kaiti SC", "STSong"],
    "微软雅黑": ["微软雅黑", "Microsoft YaHei", "PingFang SC", "Helvetica Neue"],
    "思源宋体": ["思源宋体", "Source Han Serif SC", "Noto Serif CJK SC", "宋体", "Songti SC"],
    "思源黑体": ["思源黑体", "Source Han Sans SC", "Noto Sans CJK SC", "黑体", "Heiti SC"],
    "方正书宋": ["方正书宋", "FZShuSong", "Songti SC", "STSong", "宋体"],
    # 西文
    "Times New Roman": ["Times New Roman", "Times", "Liberation Serif"],
    "Arial": ["Arial", "Helvetica", "Liberation Sans"],
    "Calibri": ["Calibri", "Carlito", "Helvetica Neue"],
    "Courier New": ["Courier New", "Courier", "Liberation Mono"],
    # macOS 端用户可能直接选这些名字，也要能给出去
    "Songti SC": ["Songti SC", "STSong", "宋体"],
    "Heiti SC": ["Heiti SC", "STHeiti", "黑体"],
    "Kaiti SC": ["Kaiti SC", "STKaiti", "楷体"],
    "STSong": ["STSong", "Songti SC", "宋体"],
    "STHeiti": ["STHeiti", "Heiti SC", "黑体"],
    "STKaiti": ["STKaiti", "Kaiti SC", "楷体"],
    "PingFang SC": ["PingFang SC", "Heiti SC", "黑体"],
}


def is_macos() -> bool:
    return sys.platform == "darwin"


# ============================================================
# 系统已安装字体探测
# ============================================================
# 记录字体名是从哪来的。用途见 fonts_from_reliable_source()。
_FONTS_SOURCE: str = "unknown"


def fonts_from_reliable_source() -> bool:
    """当前字体名集合是否来自**可靠的**来源。

    ★ 为什么需要这个判断
    -------------------
    macOS 上 `system_profiler` 不可用时（受限环境、CI runner、某些 MDM 策略），
    `_fonts_macos()` 会退回扫描字体目录，此时只能拿到**文件名**：

        Kaiti.ttc  -> kaiti
        Songti.ttc -> songti

    这些名字与字体家族名（Kaiti SC / Songti SC）不对应，更糟的是
    `kaiti` 归一化后**恰好等于 Windows 字体名 `KaiTi`**，
    于是在精确匹配那一轮就会命中，把一个 Mac 上根本不存在的字体名写进文档。

    所以：来源不可靠时禁用精确匹配，只用保守的模糊匹配（前缀关系 + 长度门槛）。
    """
    return _FONTS_SOURCE != "dirscan"


@functools.lru_cache(maxsize=1)
def installed_fonts() -> set[str]:
    """返回系统已安装字体名集合（小写归一化）。探测失败时返回空集合，
    此时 font_chain 会退回原字体名，行为等同改造前，不会更差。"""
    global _FONTS_SOURCE
    try:
        if sys.platform == "win32":
            _FONTS_SOURCE = "registry"
            return _fonts_windows()
        if is_macos():
            names, source = _fonts_macos()
            _FONTS_SOURCE = source
            return names
        _FONTS_SOURCE = "fontconfig"
        return _fonts_fontconfig()
    except Exception:
        _FONTS_SOURCE = "none"
        return set()


def _norm(s: str) -> str:
    return s.strip().lower()


def _fonts_windows() -> set[str]:
    names: set[str] = set()
    try:
        import winreg
    except ImportError:
        return names
    keys = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
    ]
    for root, sub in keys:
        try:
            with winreg.OpenKey(root, sub) as k:
                n = winreg.QueryInfoKey(k)[1]
                for i in range(n):
                    name, _val, _t = winreg.EnumValue(k, i)
                    # 值名形如「宋体 (TrueType)」，去掉后缀
                    base = name.split("(")[0].strip()
                    names.add(_norm(base))
        except OSError:
            continue
    fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    if os.path.isdir(fonts_dir):
        for fn in os.listdir(fonts_dir):
            names.add(_norm(os.path.splitext(fn)[0]))
    return names


def parse_system_profiler_fonts(stdout: str) -> set[str]:
    """解析 `system_profiler SPFontsDataType -json` 的输出，取出字体名集合。

    ★ 单独抽成函数是为了能测：macOS 适配最容易出问题的地方就是这个解析——
    不同 macOS 版本的 JSON 结构不一样，字段还可能缺失（有的条目只有 _name，
    没有 family；有的 typefaces 为空）。把解析逻辑独立出来，
    就能用真实录制的输出在 Windows/CI 上做确定性回归测试，
    不需要每次都有真机。

    容错设计：结构不对时不抛异常，返回能取到的部分；完全取不到就返回空集合，
    调用方会退回「扫描字体目录」，行为不会比改造前更差。
    """
    names: set[str] = set()
    if not stdout:
        return names
    try:
        import json
        data = json.loads(stdout)
    except Exception:
        return names
    if not isinstance(data, dict):
        return names

    def _take(d: dict) -> None:
        for key in ("_name", "family", "fullname", "postscript", "display_name"):
            v = d.get(key)
            if isinstance(v, str) and v.strip():
                names.add(_norm(v))

    # ★ 坑：不要对 SPFontsDataType 这一层取 _name。
    #   它的 _name 是**区段标签**（如「字体」/「Fonts」），不是字体名。
    #   旧实现在这里调了 _take(item)，后果是一旦 _items 为空（CI runner 上
    #   system_profiler 就返回这种结果），names 会变成 {"字体"} —— 非空，
    #   于是 _fonts_macos() 里「探测不到就退回扫描字体目录」的兜底不会触发，
    #   installed_fonts() 返回 {"字体"}，resolve_font() 找不到任何候选，
    #   最终把 Windows 字体名原样写进文档（Mac 上显示成错误字体）且不报错。
    #   由 macOS CI 真机测试发现。
    for item in data.get("SPFontsDataType", []) or []:
        if not isinstance(item, dict):
            continue
        for fam in item.get("_items", []) or []:
            if not isinstance(fam, dict):
                continue
            _take(fam)
            for typ in fam.get("typefaces", []) or []:
                if isinstance(typ, dict):
                    _take(typ)
    return names


def _fonts_macos() -> tuple[set[str], str]:
    """macOS：优先用 system_profiler 拿字体家族名，失败再退回文件扫描。

    返回 (字体名集合, 来源标记)。来源标记取值：
      "profiler" —— 来自 system_profiler，是真实字体家族名，可信
      "dirscan"  —— 来自扫描字体目录的文件名，**不可信**（见 fonts_from_reliable_source）
      "none"     —— 两条路都没拿到东西
    """
    names: set[str] = set()
    try:
        r = subprocess.run(["system_profiler", "SPFontsDataType", "-json"],
                           capture_output=True, text=True, timeout=25)
        if r.returncode == 0 and r.stdout:
            names = parse_system_profiler_fonts(r.stdout)
    except Exception:
        pass
    if names:
        return names, "profiler"
    # 退回：扫描字体目录里的文件名。
    # 注意这些名字与字体家族名不对应（Songti.ttc -> songti），
    # 且可能与 Windows 字体名撞名（Kaiti.ttc -> kaiti == KaiTi），
    # 所以标记为 dirscan，调用方据此禁用精确匹配。
    for d in ("/System/Library/Fonts", "/System/Library/Fonts/Supplemental",
              "/Library/Fonts", os.path.expanduser("~/Library/Fonts")):
        if not os.path.isdir(d):
            continue
        for root, _dirs, files in os.walk(d):
            for fn in files:
                if fn.lower().endswith((".ttf", ".ttc", ".otf", ".dfont")):
                    names.add(_norm(os.path.splitext(fn)[0]))
    return names, ("dirscan" if names else "none")


def _fonts_fontconfig() -> set[str]:
    names: set[str] = set()
    try:
        r = subprocess.run(["fc-list", "--format", "%{family}\n"],
                           capture_output=True, text=True, timeout=15)
        if r.returncode == 0:
            for line in r.stdout.splitlines():
                for fam in line.split(","):
                    if fam.strip():
                        names.add(_norm(fam))
    except Exception:
        pass
    return names


# ============================================================
# 字体名解析
# ============================================================
def _norm_key(s: str) -> str:
    """归一化成比对用的键：小写并去掉空格、下划线、连字符。"""
    return "".join(ch for ch in s.lower() if ch not in " _-")


def _fuzzy_present(cand: str, have: set[str]) -> bool:
    """保守的模糊匹配：候选名与系统里某个名字归一化后互为前缀。

    只用于「退回扫描字体目录」的降级路径——那种情况下拿到的是**文件名**
    （Songti.ttc -> songti），与字体家族名（Songti SC）并不一致，
    只做精确匹配的话这条路径等于没用。

    ★ 必须只在精确匹配全部失败后才调用。
      踩过的坑：把它和精确匹配混在一轮里遍历候选链，会让 Windows 名
      「KaiTi」因前缀关系抢在 macOS 名「Kaiti SC」前面被选中，
      结果把一个 Mac 上不存在的字体名写进文档。

    约束（避免误命中）：归一化后互为前缀，候选至少 5 字符、被比对名至少 4 字符。
    """
    key = _norm_key(cand)
    if len(key) < 5:
        return False
    for h in have:
        hk = _norm_key(h)
        if len(hk) < 4:
            continue
        if hk.startswith(key) or key.startswith(hk):
            return True
    return False


def resolve_font(name: str) -> str:
    """把逻辑字体名解析成当前系统真实存在的字体名。

    ★ 只在 macOS 上做替换，Windows 保持原样。
      原因：Windows 注册表里存的字体值是西文名（宋体 -> SimSun），
      若不加限制，「宋体」会被替换成「SimSun」。虽然二者等价，
      但 Windows 上「宋体」本来就能正常工作，没必要改动已验证的行为，
      改动只会引入不确定性。

    解析顺序（严格分两轮，不可合并）：
      第 1 轮 精确匹配：原名 -> 候选链依次精确匹配
      第 2 轮 保守模糊匹配（仅在精确全部失败时）：应付目录扫描只拿到文件名的情况

    * 非 macOS 平台 -> 原样返回
    * 探测失败 -> 原样返回（行为与改造前一致，不会更差）
    """
    if not name:
        return name
    if not is_macos():
        return name
    have = installed_fonts()
    if not have:
        return name
    reliable = fonts_from_reliable_source()

    # ---- 第 1 轮：精确匹配 ----
    # ★ 仅在字体名来源可靠时启用。目录扫描（dirscan）拿到的文件名可能与
    #   Windows 字体名撞名（Kaiti.ttc -> kaiti == KaiTi），
    #   精确匹配会因此选中一个 Mac 上不存在的名字。
    if reliable:
        if _norm(name) in have:
            return name
        for cand in FONT_ALIASES.get(name, []):
            if _norm(cand) in have:
                return cand

    # ---- 第 2 轮：保守模糊匹配（前缀关系 + 长度门槛）----
    # ★ 来源不可靠时的额外规则：先试「没有精确命中」的候选。
    #   理由：目录扫描给出的名字本身就是候选链里可能出现的东西
    #   （Kaiti.ttc -> kaiti，与 Windows 名 KaiTi 撞名）。
    #   若某个候选在集合里精确存在，很可能只是撞名，不是真的装了这个字体；
    #   反过来，只靠前缀匹配上的「Kaiti SC」才是真实的字体家族名。
    #   所以先挑「非精确命中」的，能拿到 Kaiti SC 而不是 KaiTi。
    if not reliable:
        for cand in FONT_ALIASES.get(name, []):
            if _norm(cand) not in have and _fuzzy_present(cand, have):
                return cand

    if _fuzzy_present(name, have):
        return name
    for cand in FONT_ALIASES.get(name, []):
        if _fuzzy_present(cand, have):
            return cand
    return name


def resolve_fonts_in_preset(preset: dict) -> tuple[dict, list[tuple[str, str, str]]]:
    """批量解析预设里的字体。返回（新预设, 替换记录）。

    替换记录为 (区域, 原字体, 实际字体)，用于在界面上告诉用户发生了什么。
    这不改动原预设对象，返回的是副本。
    """
    import copy
    out = copy.deepcopy(preset)
    changes: list[tuple[str, str, str]] = []
    for region, fmt in out.items():
        orig = getattr(fmt, "font_ea", None)
        if orig:
            res = resolve_font(orig)
            if res != orig:
                fmt.font_ea = res
                changes.append((region, orig, res))
        for attr in ("font_ascii",):
            orig2 = getattr(fmt, attr, None)
            if orig2:
                res2 = resolve_font(orig2)
                if res2 != orig2:
                    setattr(fmt, attr, res2)
    return out, changes


# 界面上 Windows 中文字体名 -> macOS 对应字体名。
# ★ 提到模块级是为了能被测试直接断言（原先是 platform_font_choices 内的局部变量）。
#   显式列出而不是查 FONT_ALIASES，因为这两件事目的不同：
#   FONT_ALIASES 是「写文档时该用哪个真实字体名」，
#   这里是「界面上该给用户看哪个选项」，且必须保证目标不含 Windows 字体名。
MAC_CHOICE_MAP: dict[str, str] = {
    "宋体": "Songti SC",
    "黑体": "Heiti SC",
    "楷体": "Kaiti SC",
    # 仿宋曾映射到 STFangsong，但它是否随 macOS 出厂无法确证。
    # 已确证存在的中文衬线字体里，Kaiti SC 最接近仿宋的观感，故用它兜底。
    "仿宋": "Kaiti SC",
    "仿宋_GB2312": "Kaiti SC",
    "微软雅黑": "PingFang SC",
    "思源宋体": "Songti SC",
    "思源黑体": "Heiti SC",
    "方正书宋": "Songti SC",
}


def platform_font_choices(base_choices: list[str]) -> list[str]:
    """按平台给出界面上展示的字体选项。

    macOS 上把不可用的 Windows 字体替换成对应的 Mac 字体，
    避免用户以为选了「宋体」就真的生效。
    """
    if not is_macos():
        return list(base_choices)
    seen, out = set(), []
    for c in base_choices:
        v = MAC_CHOICE_MAP.get(c, c)
        if v not in seen:
            seen.add(v)
            out.append(v)
    return out


def open_path(path: str) -> bool:
    """跨平台用系统默认程序打开文件。

    原实现只有 os.startfile，在 macOS 上会直接报错。
    """
    try:
        if sys.platform == "win32":
            os.startfile(path)              # type: ignore[attr-defined]
            return True
        if is_macos():
            subprocess.Popen(["open", path])
            return True
        subprocess.Popen(["xdg-open", path])
        return True
    except Exception:
        return False


if __name__ == "__main__":
    print("平台: %s   是否 macOS: %s" % (sys.platform, is_macos()))
    have = installed_fonts()
    print("探测到系统字体数: %d" % len(have))
    for probe in ("宋体", "黑体", "楷体", "Songti SC", "不存在的字体XYZ",
                  "Times New Roman"):
        print("  %-18s -> %s" % (probe, resolve_font(probe)))
