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
改写命令行入口
=============
用法:
    python rewrite_cli.py <输入.docx> [输出.docx] [预设.json]
不带参数时使用内置默认预设。
"""
import json
import os
import sys

from rewriter import (
    rewrite, rewrite_report, DEFAULT_PRESET, preset_from_dict, preset_to_dict,
)

HERE = os.path.dirname(os.path.abspath(__file__))


def load_preset(path: str | None):
    if path and os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return preset_from_dict(json.load(f)), path
    # 默认找工具目录下的「我的预设.json」
    auto = os.path.join(HERE, "我的预设.json")
    if os.path.exists(auto):
        with open(auto, encoding="utf-8") as f:
            return preset_from_dict(json.load(f)), auto
    return DEFAULT_PRESET, "内置默认预设"


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("选项:")
        print("  --keep-numbering  保留原有编号（默认会重排编号：把自动编号转成文字编号）")
        print("  --no-superscript  不把正文引用标记 [1] 设为上标（默认会设）")
        print("  --no-note-marks   不把注释序号 ①② 设为上标（默认会设）")
        print("  --number-style X  编号体系：sysu(默认，一、/（一）/1、/（1）) / arabic")
        print("  --dry             只演算不写文件")
        sys.exit(1)
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith("--") \
        else os.path.splitext(src)[0] + "_已排版.docx"
    preset_path = None
    for a in sys.argv[2:]:
        if a.endswith(".json"):
            preset_path = a
    preset, preset_name = load_preset(preset_path)

    do_renumber = "--keep-numbering" not in sys.argv
    do_cite_sup = "--no-superscript" not in sys.argv
    do_note_sup = "--no-note-marks" not in sys.argv
    number_style = "sysu"
    if "--number-style" in sys.argv:
        number_style = sys.argv[sys.argv.index("--number-style") + 1]
    dry = "--dry" in sys.argv

    print("使用预设: %s" % preset_name)
    print("编号处理: %s" % ("开启（%s 体系，重排编号并去掉自动编号）" % number_style
                        if do_renumber else "关闭（保留原有编号）"))
    print("上标处理: %s" % (
        "引用标记 %s / 注释序号 %s" % (
            "开" if do_cite_sup else "关", "开" if do_note_sup else "关")))
    print()
    result = rewrite(src, out, preset=preset, dry_run=dry,
                     renumber_headings=do_renumber, number_style=number_style,
                     superscript_citations=do_cite_sup,
                     superscript_notes=do_note_sup)
    print(rewrite_report(result))

    # 顺手导出本次实际使用的预设，便于用户拿去界面上改
    dump = os.path.join(HERE, "本次使用的预设.json")
    with open(dump, "w", encoding="utf-8") as f:
        json.dump(preset_to_dict(preset), f, ensure_ascii=False, indent=2)
    print("\n本次使用的预设已导出: %s" % dump)


if __name__ == "__main__":
    main()
