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
@functools.lru_cache(maxsize=1)
def installed_fonts() -> set[str]:
    """返回系统已安装字体名集合（小写归一化）。探测失败时返回空集合，
    此时 font_chain 会退回原字体名，行为等同改造前，不会更差。"""
    try:
        if sys.platform == "win32":
            return _fonts_windows()
        if is_macos():
            return _fonts_macos()
        return _fonts_fontconfig()
    except Exception:
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

    for item in data.get("SPFontsDataType", []) or []:
        if not isinstance(item, dict):
            continue
        _take(item)
        for fam in item.get("_items", []) or []:
            if not isinstance(fam, dict):
                continue
            _take(fam)
            for typ in fam.get("typefaces", []) or []:
                if isinstance(typ, dict):
                    _take(typ)
    return names


def _fonts_macos() -> set[str]:
    """macOS：优先用 system_profiler 拿字体家族名，失败再退回文件扫描。"""
    names: set[str] = set()
    try:
        r = subprocess.run(["system_profiler", "SPFontsDataType", "-json"],
                           capture_output=True, text=True, timeout=25)
        if r.returncode == 0 and r.stdout:
            names = parse_system_profiler_fonts(r.stdout)
    except Exception:
        pass
    if not names:
        # 退回：扫描字体目录里的文件名（不含 .ttc 内部家族名，但聊胜于无）
        for d in ("/System/Library/Fonts", "/Library/Fonts",
                  os.path.expanduser("~/Library/Fonts")):
            if not os.path.isdir(d):
                continue
            for root, _dirs, files in os.walk(d):
                for fn in files:
                    if fn.lower().endswith((".ttf", ".ttc", ".otf", ".dfont")):
                        names.add(_norm(os.path.splitext(fn)[0]))
    return names


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
def resolve_font(name: str) -> str:
    """把逻辑字体名解析成当前系统真实存在的字体名。

    ★ 只在 macOS 上做替换，Windows 保持原样。
      原因：Windows 注册表里存的字体值是西文名（宋体 -> SimSun），
      若不加限制，「宋体」会被替换成「SimSun」。虽然二者等价，
      但 Windows 上「宋体」本来就能正常工作，没必要改动已验证的行为，
      改动只会引入不确定性。

    * 非 macOS 平台 -> 原样返回
    * macOS：系统里有这个名字 -> 原样返回
    * macOS：没有 -> 按候选表找第一个存在的
    * 探测失败 -> 原样返回（行为与改造前一致，不会更差）
    """
    if not name:
        return name
    if not is_macos():
        return name
    have = installed_fonts()
    if not have:
        return name
    if _norm(name) in have:
        return name
    for cand in FONT_ALIASES.get(name, []):
        if _norm(cand) in have:
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
