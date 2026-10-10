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
上标处理
========
把正文里的引用标注变成上标。

覆盖两类标记
------------
1. 参考文献引用标记：[1] [1,2] [1-3] ［1］
2. 注释序号：①②③…（中文学术规范常用「篇末注」，序号标在被注词条右上角）

关键约束（每条都是实测/规范推出来的）
------------------------------------
* 只处理正文类区域。标题里的 (1) 是层级编号，不是引用，不能上标。
* 参考文献条目区一律不动：[1] 是条目自己的序号，规范要求「左顶格」，
  上了标就错了。
* 段落开头的 [n] 若是条目式（后面跟作者名或「.」），判为编号不上标。
"""
from __future__ import annotations

import copy
import re

from docx.oxml.ns import qn

# 参考文献引用标记：[1] [1,2] [1-3] [1，2] ／ 全角版本
CITE_RE = re.compile(r"[\[［]\s*\d+(?:\s*[,，\-–—~～]\s*\d+)*\s*[\]］]")
# 注释序号：①..⑳ 及常见符号
NOTE_RE = re.compile(r"[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]")

# 参考文献序号不会很大。超过这个位数说明更像年份或数量，不是引用。
# 必要性：实测「统计区间为[2010-2015]年」会被误当引用上标。
MAX_REF_NUMBER_DIGITS = 3


def is_reference_marker(marker: str) -> bool:
    """判断方括号标记是不是参考文献引用，而不是年份区间之类。

    [3-5]  -> True（引用区间，数字很小）
    [2010-2015] -> False（年份区间）
    [1,2] -> True
    [100] -> True
    """
    nums = re.findall(r"\d+", marker)
    if not nums:
        return False
    return all(len(n) <= MAX_REF_NUMBER_DIGITS for n in nums)

# 这些区域允许上标
SUPERSCRIPT_REGIONS = {"body", "abstract", "keywords", "quote", "ack", "caption"}
# 这些区域绝对不动（序号是它们自身的编号）
NEVER_REGIONS = {"ref_item", "ref_head", "toc", "title", "h1", "h2", "h3",
                 "cover_field", "blank"}

# 段落开头形如「[1] 作者...」——这是条目序号，不是引用
_ENTRY_LIKE = re.compile(r"^\s*[\[［]\s*\d+\s*[\]］]\s*(?=[\u4e00-\u9fffA-Za-z（(])")


def find_markers(text: str) -> list[tuple[int, int, str]]:
    """返回 [(起点, 终点, 匹配内容)]，按位置排序。"""
    out = []
    for m in NOTE_RE.finditer(text):
        out.append((m.start(), m.end(), m.group(0)))
    for m in CITE_RE.finditer(text):
        if not is_reference_marker(m.group(0)):
            continue          # 年份区间等，不是引用
        out.append((m.start(), m.end(), m.group(0)))
    out.sort()
    # 去掉重叠（如全角括号里再嵌套的情况，理论上不会出现，保险起见）
    cleaned = []
    last_end = -1
    for s, e, t in out:
        if s >= last_end:
            cleaned.append((s, e, t))
            last_end = e
    return cleaned


def has_marker(text: str) -> bool:
    return bool(CITE_RE.search(text) or NOTE_RE.search(text))


def _set_superscript(run, on: bool = True) -> None:
    """给 run 设置/取消上标。"""
    rpr = run._element.get_or_add_rPr()
    va = rpr.find(qn("w:vertAlign"))
    if on:
        if va is None:
            va = rpr.makeelement(qn("w:vertAlign"), {})
            rpr.append(va)
        va.set(qn("w:val"), "superscript")
    elif va is not None:
        rpr.remove(va)


def _split_run_at_markers(run, markers: list[tuple[int, int]]) -> list:
    """把一个 run 按其内部标记位置切成多个 run，返回新 run 列表。

    为什么要切：一个 run 的文字是连续的，上标是 run 级属性。
    如果标记和正文挤在同一个 run 里（实测「…理论突破[1]。基于…」就是一整个 run），
    只能先把 run 拆开，才能只给标记上标。

    实现要点：不删原 run，而是在其后连续插入新 run，最后把原 run 清空。
    这样能天然继承原 run 的全部格式（字体、字号、书签引用等）。
    """
    from docx.text.run import Run

    r_el = run._element
    parent = r_el.getparent()
    text = run.text
    if not text:
        return []

    pieces = []
    pos = 0
    for s, e in markers:
        if s > pos:
            pieces.append((text[pos:s], False))
        pieces.append((text[s:e], True))
        pos = e
    if pos < len(text):
        pieces.append((text[pos:], False))

    if len(pieces) <= 1:
        return []

    rpr_src = r_el.find(qn("w:rPr"))
    new_runs = []
    last_el = r_el
    for piece_text, is_mark in pieces:
        new_el = copy.deepcopy(r_el)
        # 清掉所有文字节点，只留一个
        for t in new_el.findall(qn("w:t")):
            new_el.remove(t)
        t = new_el.makeelement(qn("w:t"), {})
        t.set(qn("xml:space"), "preserve")
        t.text = piece_text
        new_el.append(t)
        last_el.addnext(new_el)
        last_el = new_el
        new_runs.append((Run(new_el, run._parent), is_mark))

    # 原 run 置空，保持结构不变（有些文档的序号挂在原 run 上，删掉会出问题）
    for t in r_el.findall(qn("w:t")):
        r_el.remove(t)
    return new_runs


def apply_superscripts(paras: list[dict], doc,
                       region_overrides: dict[int, str] | None = None,
                       do_citations: bool = True,
                       do_note_marks: bool = True) -> dict:
    """给正文类区域的引用标记/注释序号加是上标。返回统计。

    do_citations: 是否处理 [1] 式参考文献引用（规范未强制，默认开）
    do_note_marks: 是否处理 ①② 式注释序号（规范明确要求右上方）
    """
    stats = {
        "paragraphs_touched": 0,
        "markers_superscripted": 0,
        "skipped_ref_items": 0,
        "skipped_entry_like": 0,
        "details": [],
    }

    for info in paras:
        if info["empty"]:
            continue
        idx = info["idx"]
        region = (region_overrides or {}).get(idx, info["region"])

        if region in NEVER_REGIONS:
            if region == "ref_item" and has_marker(info["text"]):
                stats["skipped_ref_items"] += 1
            continue
        if region not in SUPERSCRIPT_REGIONS:
            continue

        # 段落开头像条目编号的，不当引用
        if _ENTRY_LIKE.match(info["text"]):
            stats["skipped_entry_like"] += 1
            continue

        p = doc.paragraphs[idx]
        touched = 0

        # ★ 必须按「整段拼接文本」来匹配标记，不能逐 run 匹配。
        #   踩过的坑：Word/WPS 在编辑后经常把一个标记拆到多个 run，
        #   例如 "跨 run 的[" + "4" + "]标记。" 三个 run。
        #   逐 run 匹配时每个 run 里都看不到完整标记，于是全部漏掉（实测漏检）。
        #   正确做法：拼接全段文本定位标记，再把命中位置映射回各个 run。
        runs = [r for r in p.runs if r.text]
        if not runs:
            continue
        texts = [r.text for r in runs]
        full = "".join(texts)

        char_spans = []
        if do_note_marks:
            char_spans += [(m.start(), m.end()) for m in NOTE_RE.finditer(full)]
        if do_citations:
            for m in CITE_RE.finditer(full):
                if is_reference_marker(m.group(0)):
                    char_spans.append((m.start(), m.end()))
        if not char_spans:
            continue
        char_spans.sort()

        # 把字符区间映射到 (run 下标, run 内起点, run 内终点)
        run_spans = []
        pos = 0
        for ri, t in enumerate(texts):
            rs, re_ = pos, pos + len(t)
            for cs, ce in char_spans:
                a, b = max(cs, rs), min(ce, re_)
                if a < b:
                    run_spans.append((ri, a - rs, b - rs))
            pos = re_
        if not run_spans:
            continue

        # 同一个 run 内的多段命中合并，避免重复处理
        by_run: dict[int, list[tuple[int, int]]] = {}
        for ri, a, b in run_spans:
            by_run.setdefault(ri, []).append((a, b))

        for ri, spans in by_run.items():
            run = runs[ri]
            spans.sort()
            merged = []
            for s, e in spans:
                if merged and s <= merged[-1][1]:
                    merged[-1] = (merged[-1][0], max(merged[-1][1], e))
                else:
                    merged.append((s, e))

            # 整段 run 就是一个标记 -> 直接加属性，不用切分
            if len(merged) == 1 and merged[0] == (0, len(run.text)):
                _set_superscript(run, True)
                touched += 1
                continue

            new_runs = _split_run_at_markers(run, merged)
            if new_runs:
                for nr, is_mark in new_runs:
                    if is_mark:
                        _set_superscript(nr, True)
                        touched += 1
            else:
                # 切分失败（例如 run 里没有 w:t 节点），退化为整体上标，
                # 宁可多标一点也不漏标——多标会被用户看到并纠正，漏标不会。
                _set_superscript(run, True)
                touched += 1

        if touched:
            stats["paragraphs_touched"] += 1
            stats["markers_superscripted"] += touched
            stats["details"].append({
                "idx": idx, "region": region, "count": touched,
                "text": info["text"][:44],
            })

    return stats


def superscript_report(stats: dict) -> str:
    lines = ["上标处理"]
    lines.append("  加为的标记数: %d（涉及 %d 个段落）" % (
        stats["markers_superscripted"], stats["paragraphs_touched"]))
    if stats["skipped_ref_items"]:
        lines.append("  跳过参考文献条目: %d 段（其 [n] 是条目序号，规范要求左顶格）" %
                     stats["skipped_ref_items"])
    if stats["skipped_entry_like"]:
        lines.append("  跳过条目式开头: %d 段" % stats["skipped_entry_like"])
    for d in stats["details"][:20]:
        lines.append("     #%-4d [%s] %d 处  %s" % (
            d["idx"], d["region"], d["count"], d["text"]))
    return "\n".join(lines)
