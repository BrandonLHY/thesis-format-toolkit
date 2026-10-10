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
预设系统 + 文档改写器
====================
预设：每个"区域类型"对应一套字体字号等格式设定。这是界面上可调的东西。
改写：按预设，把一份 docx 的文字统一套用格式，输出一份**新文档**，原件不动。

保真原则
--------
不新建文档，而是打开原文档、就地改每段的格式、另存为新文件。
这样图片、页眉页脚、嵌入字体、分节、目录域、公式全部原样保留。
"""
from __future__ import annotations

import copy
import json
from dataclasses import dataclass, asdict, field
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

from region_detector import detect_regions, REGION_TYPES, REGION_LABEL

# ============================================================
# 字号对照（中文号数 <-> 磅值）
# ============================================================
CN_SIZE = {
    "初号": 42.0, "小初": 36.0, "一号": 26.0, "小一": 24.0, "二号": 22.0, "小二": 18.0,
    "三号": 16.0, "小三": 15.0, "四号": 14.0, "小四": 12.0, "五号": 10.5, "小五": 9.0,
}
SIZE_CHOICES = ["初号", "小初", "一号", "小一", "二号", "小二",
                "三号", "小三", "四号", "小四", "五号", "小五"]

FONT_CHOICES = ["宋体", "黑体", "楷体", "仿宋_GB2312", "仿宋", "微软雅黑",
                "思源宋体", "思源黑体", "方正书宋", "Times New Roman", "Arial"]

ALIGN_CHOICES = ["未指定", "左对齐", "居中", "右对齐", "两端对齐"]


# ============================================================
# 预设
# ============================================================
@dataclass
class RegionFormat:
    """一个区域类型的格式设定——界面上每行就是这个。"""
    font_ea: str = "宋体"          # 中文字体
    font_ascii: str = "Times New Roman"  # 西文字体
    size: str = "小四"             # 字号（中文号数）
    bold: bool = False
    align: str = "未指定"
    first_line_indent_chars: int = 0   # 首行缩进（字符）
    line_spacing: float = 1.5
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0

    @property
    def pt(self) -> float:
        return CN_SIZE.get(self.size, 12.0)


DEFAULT_PRESET: dict[str, RegionFormat] = {
    # 依据：中山大学本科毕业论文格式要求；规范未覆盖的部分沿用样例论文的实际做法
    "title":       RegionFormat("黑体", "Times New Roman", "三号", True, "居中", 0, 1.5, 0, 12),
    "cover_field": RegionFormat("黑体", "Times New Roman", "四号", True, "左对齐", 0, 1.5, 0, 0),
    "h1":          RegionFormat("黑体", "Times New Roman", "四号", False, "居中", 0, 1.5, 12, 12),
    "h2":          RegionFormat("黑体", "Times New Roman", "四号", False, "左对齐", 0, 1.5, 6, 6),
    "h3":          RegionFormat("宋体", "Times New Roman", "小四", True, "左对齐", 0, 1.5, 3, 3),
    "body":        RegionFormat("宋体", "Times New Roman", "小四", False, "两端对齐", 2, 1.5, 0, 0),
    "abstract":    RegionFormat("楷体", "Times New Roman", "五号", False, "两端对齐", 2, 1.5, 0, 0),
    "keywords":    RegionFormat("楷体", "Times New Roman", "五号", False, "左对齐", 0, 1.5, 6, 6),
    "caption":     RegionFormat("宋体", "Times New Roman", "五号", False, "居中", 0, 1.5, 3, 3),
    # 规范原文：「注释、参考文献标题 小五号黑体并加冒号」——冒号是规范要求的一部分
    "ref_head":    RegionFormat("黑体", "Times New Roman", "小五", False, "居中", 0, 1.5, 12, 6),
    "ref_item":    RegionFormat("宋体", "Times New Roman", "小五", False, "左对齐", 0, 1.5, 0, 0),
    "toc":         RegionFormat("宋体", "Times New Roman", "小四", False, "左对齐", 0, 1.5, 0, 0),
    "ack":         RegionFormat("宋体", "Times New Roman", "小四", False, "两端对齐", 2, 1.5, 0, 0),
    "quote":       RegionFormat("楷体", "Times New Roman", "五号", False, "两端对齐", 2, 1.25, 0, 0),
}

ALIGN_MAP = {
    "未指定": None,
    "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
    "居中": WD_ALIGN_PARAGRAPH.CENTER,
    "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
    "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def preset_to_dict(preset: dict[str, RegionFormat]) -> dict:
    return {k: asdict(v) for k, v in preset.items()}


def preset_from_dict(d: dict) -> dict[str, RegionFormat]:
    out = {}
    for k, v in d.items():
        out[k] = RegionFormat(**{kk: vv for kk, vv in v.items()
                                 if kk in RegionFormat.__dataclass_fields__})
    return out


# ============================================================
# 底层 XML 格式化
# ============================================================
def _strip_theme_fonts(rf) -> None:
    """删掉 rFonts 上的主题字体引用。

    为什么必须删：OOXML 里 w:eastAsiaTheme="minorEastAsia" 这类主题引用与
    w:eastAsia="宋体" 同时存在时，**主题引用优先**，具体字体名被忽略。
    实测成稿里 483 个 rFonts 同时带两套属性，导致字号改了、字体却不变
    （Word 把字体报成「+中文正文」而不是「宋体」）。
    """
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rf.get(qn(attr)) is not None:
            del rf.attrib[qn(attr)]


def _kill_para_theme_fonts(p) -> None:
    """清掉段落级 rPr 里的主题字体引用。

    段落级 rPr 会作为 run 的继承基。虽然 run 级直接格式通常优先，
    但段落级留着主题引用会让空 run 和域代码仍显示旧字体，因此一并清掉。
    """
    ppr = p._element.pPr
    if ppr is None:
        return
    prpr = ppr.find(qn("w:rPr"))
    if prpr is None:
        return
    for rf in prpr.findall(qn("w:rFonts")):
        _strip_theme_fonts(rf)


def _set_run_format(run, fmt: RegionFormat) -> None:
    """直接写 XML，不用 python-docx 的 font.name（它只设 ascii/hAnsi，不设 eastAsia）。"""
    rpr = run._element.get_or_add_rPr()

    # 字体：西文 ascii/hAnsi + 中文 eastAsia 必须分别设置
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = rpr.makeelement(qn("w:rFonts"), {})
        rpr.insert(0, rf)
    _strip_theme_fonts(rf)      # 关键：先清主题引用，否则具体字体名无效
    rf.set(qn("w:ascii"), fmt.font_ascii)
    rf.set(qn("w:hAnsi"), fmt.font_ascii)
    rf.set(qn("w:eastAsia"), fmt.font_ea)
    rf.set(qn("w:cs"), fmt.font_ascii)
    rf.set(qn("w:hint"), "eastAsia")   # 提示按东亚字体渲染中文

    # 字号：sz 是西文，szCs 是复杂文种，中文实际按 sz 生效
    half = str(int(round(fmt.pt * 2)))
    for tag in ("w:sz", "w:szCs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set(qn("w:val"), half)

    # 加粗/斜体
    for tag, on in (("w:b", fmt.bold), ("w:bCs", fmt.bold)):
        el = rpr.find(qn(tag))
        if on:
            if el is None:
                el = rpr.makeelement(qn(tag), {})
                rpr.append(el)
            el.set(qn("w:val"), "1")
        elif el is not None:
            rpr.remove(el)


def _set_para_format(p, fmt: RegionFormat) -> None:
    # 段落级 rPr 里的主题字体引用必须清掉，否则会盖住 run 级的具体字体
    _kill_para_theme_fonts(p)

    pf = p.paragraph_format

    align = ALIGN_MAP.get(fmt.align)
    if align is not None:
        pf.alignment = align

    pf.line_spacing = fmt.line_spacing
    pf.space_before = Pt(fmt.space_before_pt)
    pf.space_after = Pt(fmt.space_after_pt)

    # 缩进：用字符单位，这样换字号后仍保持「2字符」的语义
    ppr = p._element.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if fmt.first_line_indent_chars:
        if ind is None:
            ind = ppr.makeelement(qn("w:ind"), {})
            ppr.append(ind)
        ind.set(qn("w:firstLineChars"), str(int(fmt.first_line_indent_chars * 100)))
        # 同时给一个绝对值兜底，部分阅读器不认 Chars
        ind.set(qn("w:firstLine"), str(int(fmt.pt * 20 * fmt.first_line_indent_chars)))
    else:
        if ind is not None:
            for a in ("w:firstLineChars", "w:firstLine"):
                if ind.get(qn(a)) is not None:
                    del ind.attrib[qn(a)]


def _format_nested_paragraphs(doc, preset: dict[str, RegionFormat],
                              out_path: str, dry_run: bool) -> int:
    """处理被 w:sdt（内容控件）包裹的段落——典型是自动目录 TOC。

    两个必须绕开的 python-docx 限制：
    1. doc.paragraphs 只返回 body 直属段落，看不到 sdt 内部。
       实测一份文档 Word 报告 215 段，python-docx 只给 187 段，
       差的 28 段全是目录条目，因此从未被改写，仍保持极小字号（看不见）。
    2. p.runs 只返回段落直属的 run，**不包含 w:hyperlink 里的 run**。
       而目录条目恰恰是超链接，所以只遍历 p.runs 同样会全部漏掉。
       必须用 .//w:r 抓全部后代 run。

    返回被处理的段落数。
    """
    from docx.text.paragraph import Paragraph as _P
    from docx.text.run import Run as _R

    body = doc.element.body
    n = 0

    def format_all_runs(p_el) -> int:
        """给段落内所有 run（含超链接内的）套格式，返回处理的 run 数。"""
        cnt = 0
        for r_el in p_el.findall(".//" + qn("w:r")):
            # 跳过域代码本身（instrText），只改显示文本
            if r_el.find(qn("w:instrText")) is not None:
                continue
            try:
                run = _R(r_el, _P(p_el, doc))
                _set_run_format(run, fmt)     # 空 run 也格式化，避免残留主题字体
                cnt += 1
            except Exception:
                continue
        return cnt

    # ---- 第一遍：sdt（内容控件）内的目录 ----
    for child in list(body):
        if child.tag != qn("w:sdt"):
            continue
        instrs = [x.text or "" for x in child.findall(".//" + qn("w:instrText"))]
        if not any("TOC" in s for s in instrs):
            continue
        fmt = preset.get("toc")
        if fmt is None:
            continue
        for p_el in child.findall(".//" + qn("w:p")):
            try:
                p = _P(p_el, doc)
                format_all_runs(p_el)
                _set_para_format(p, fmt)
                n += 1
            except Exception:
                continue

    # ---- 第二遍：body 内其它非直属段落（更深的嵌套）----
    for p_el in body.iter(qn("w:p")):
        if p_el.getparent() is body:
            continue          # 直属段落已由主流程处理
        anc, in_table = p_el.getparent(), False
        while anc is not None and anc is not body:
            if anc.tag == qn("w:tbl"):
                in_table = True
                break
            anc = anc.getparent()
        if in_table:
            continue
        instrs = [x.text or "" for x in p_el.findall(".//" + qn("w:instrText"))]
        if not any("TOC" in s for s in instrs):
            continue
        fmt = preset.get("toc")
        if fmt is None:
            continue
        try:
            p = _P(p_el, doc)
            format_all_runs(p_el)
            _set_para_format(p, fmt)
            n += 1
        except Exception:
            continue

    return n


# ============================================================
# 改写主流程
# ============================================================
def rewrite(docx_path: str, out_path: str,
            preset: dict[str, RegionFormat] | None = None,
            dry_run: bool = False,
            renumber_headings: bool = False,
            number_style: str = "sysu",
            region_overrides: dict[int, str] | None = None,
            superscript_citations: bool = True,
            superscript_notes: bool = True) -> dict:
    """按预设改写文档格式，输出新文件。原件不动。

    renumber_headings: 是否把标题自动编号改成文字编号并重排。
    region_overrides: {段落索引: 区域key}，人工纠正识别结果。
    superscript_citations: 是否把正文 [1] 式引用标为上标（默认开）。
        规范未强制，但多数中文学术体例如此；不需要可关掉。
    superscript_notes: 是否把 ①② 式注释序号标为上标（规范明确要求）。
    """
    preset = preset or DEFAULT_PRESET

    # ---- 跨平台字体解析 ----
    # macOS 上没有「宋体/黑体/楷体」这几个名字，必须替换成系统真实存在的字体，
    # 否则 Word 回退成默认字体，用户看到的就是「格式没改成功」。
    # Windows 上不做任何替换（保持已验证的行为）。
    font_changes: list[tuple[str, str, str]] = []
    try:
        from platform_fonts import is_macos, resolve_fonts_in_preset
        if is_macos():
            preset, font_changes = resolve_fonts_in_preset(preset)
    except Exception:
        pass

    paras, summary = detect_regions(docx_path)
    doc = summary["doc"]

    # ---- 编号处理（必须在套格式之前，因为会改动文字）----
    num_result = None
    if renumber_headings:
        from numbering import renumber
        region_of = {p["idx"]: p["region"] for p in paras}
        if region_overrides:
            region_of.update(region_overrides)
        num_result = renumber(doc, paras, region_of, style=number_style, dry_run=dry_run)

    changed, skipped, by_region = [], [], {}
    for info in paras:
        if info["empty"]:
            continue
        region = info["region"]
        if region_overrides and info["idx"] in region_overrides:
            region = region_overrides[info["idx"]]
        if region in ("blank",):
            continue
        fmt = preset.get(region)
        if fmt is None:
            skipped.append((info["idx"], region, "该区域未在预设中定义"))
            continue

        p = doc.paragraphs[info["idx"]]

        # 编号处理后文字已变，重新读取
        if renumber_headings:
            info["text"] = p.text.strip()

        # 保护：含图片/公式的段落只改段落级格式，不动 run，避免破坏对象
        has_object = info["has_drawing"]
        try:
            if not has_object:
                # 空 run 也要格式化：否则它会显示为主题默认字体，
                # 出现「远东=+中文正文」这种混杂，字号也会报异常值。
                for run in p.runs:
                    _set_run_format(run, fmt)
            _set_para_format(p, fmt)
        except Exception as e:
            skipped.append((info["idx"], region, "改写失败: %r" % e))
            continue

        by_region[region] = by_region.get(region, 0) + 1
        changed.append({
            "idx": info["idx"], "region": region, "text": info["text"][:40],
            "before": {"pt": info["pt_mode"], "ea": info["ea_mode"], "bold": info["bold_mode"]},
            "after": {"font": fmt.font_ea, "size": fmt.size, "pt": fmt.pt, "bold": fmt.bold},
            "object_protected": has_object,
        })

    # ---- 上标：必须在套格式之后 ----
    # 因为切分 run 时会复制原 run 的 rPr，先套好格式再切，新 run 自然继承正确字体字号。
    sup_result = None
    if superscript_citations or superscript_notes:
        from superscript import apply_superscripts
        sup_result = apply_superscripts(
            paras, doc, region_overrides=region_overrides,
            do_citations=superscript_citations,
            do_note_marks=superscript_notes)

    if not dry_run:
        nested_count = _format_nested_paragraphs(doc, preset, out_path, dry_run)
        doc.save(out_path)
    else:
        nested_count = 0

    return {
        "source": docx_path,
        "output": None if dry_run else out_path,
        "changed_count": len(changed),
        "nested_count": nested_count,
        "numbering": num_result,
        "superscript": sup_result,
        "font_changes": font_changes,
        "by_region": by_region,
        "skipped": skipped,
        "changes": changed,
        "low_confidence": [
            {"idx": p["idx"], "region": p["region"], "conf": p["conf"],
             "text": p["text"][:40]}
            for p in summary["low_confidence"]
        ],
    }


def rewrite_report(result: dict) -> str:
    lines = []
    lines.append("改写完成：%d 段已套用新格式" % result["changed_count"])
    if result.get("nested_count"):
        lines.append("另有 %d 段在内容控件（目录域）内，也已处理" % result["nested_count"])
    if result["output"]:
        lines.append("输出文件：%s" % result["output"])
    if result.get("numbering"):
        from numbering import renumber_report
        lines.append("")
        lines.append(renumber_report(result["numbering"]))
    if result.get("superscript"):
        from superscript import superscript_report
        lines.append("")
        lines.append(superscript_report(result["superscript"]))
    if result.get("font_changes"):
        lines.append("")
        lines.append("字体适配（本机缺少原字体，已替换为可用字体）：")
        for region, old, new in result["font_changes"]:
            lines.append("  %-14s %s -> %s" % (region, old, new))
    lines.append("")
    lines.append("各区域改写段数：")
    for k, n in sorted(result["by_region"].items(), key=lambda kv: -kv[1]):
        lines.append("  %-14s %3d 段" % (REGION_LABEL.get(k, k), n))
    if result["skipped"]:
        lines.append("")
        lines.append("跳过 %d 段：" % len(result["skipped"]))
        for idx, region, why in result["skipped"][:10]:
            lines.append("  #%d [%s] %s" % (idx, region, why))
    if result["low_confidence"]:
        lines.append("")
        lines.append("以下段落区域判定置信度低，建议人工确认：")
        for p in result["low_confidence"]:
            lines.append("  #%-4d [%s %.2f] %s" % (
                p["idx"], REGION_LABEL.get(p["region"], p["region"]), p["conf"], p["text"]))
    return "\n".join(lines)


if __name__ == "__main__":
    import sys
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else "改写结果.docx"
    dry = "--dry" in sys.argv
    res = rewrite(src, out, dry_run=dry)
    print(rewrite_report(res))
    if "--json" in sys.argv:
        with open("改写明细.json", "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=2)
        print("\n已写出 改写明细.json")
