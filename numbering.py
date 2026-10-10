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
标题编号处理器
==============
用途：把「自动编号」转成「文字编号」，按层级重新编号。

为什么需要
----------
WPS/Word 的自动编号（w:numPr）有两个常见病：
1. 多个一级标题共用同一个 numId，导致编号全局连续、不按章重置
   （实测：某章第二个二级标题显示成「（二）」，因为它继承了上一章的计数）
2. 只改字体字号完全碰不到它——序号不是文字，是渲染层加的

编号体系（中山大学本科毕业论文格式要求）
    一级：一、二、三……
    二级：（一）（二）（三）……
    三级：1、2、3……
    四级：（1）（2）（3）……
"""
from __future__ import annotations

import re

from docx.oxml.ns import qn

CN_NUM = "一二三四五六七八九十"

# 已存在的编号前缀，全部要剥掉后重排
PREFIX_PATTERNS = [
    re.compile(r"^[一二三四五六七八九十]+\s*[、\.]\s*"),
    re.compile(r"^[（(][一二三四五六七八九十]+[）)]\s*"),
    re.compile(r"^\d+\.\d+(?:\.\d+)*[\s　]*"),      # 1.1 / 1.1.1
    re.compile(r"^\d+\s*[、\.]\s*"),                  # 1、 / 1.
    re.compile(r"^[（(]\d+[）)]\s*"),                # （1）
]

# 这些前置/后置部分不参与编号
NO_NUMBER_TITLES = ("摘要", "abstract", "目录", "contents",
                    "参考文献", "注释", "致谢", "附录")


def cn_number(n: int) -> str:
    """1 -> 一, 11 -> 十一, 21 -> 二十一"""
    if n <= 0:
        return str(n)
    if n <= 10:
        return CN_NUM[n - 1]
    if n < 20:
        return "十" + CN_NUM[n - 11]
    tens, ones = divmod(n, 10)
    return CN_NUM[tens - 1] + "十" + (CN_NUM[ones - 1] if ones else "")


def strip_number_prefix(text: str) -> str:
    """剥掉已有的编号前缀，返回纯标题文字。"""
    t = text.lstrip()
    changed = True
    while changed:
        changed = False
        for rx in PREFIX_PATTERNS:
            m = rx.match(t)
            if m:
                t = t[m.end():]
                changed = True
                break
    return t.strip()


def make_prefix(level: int, n: int, style: str = "sysu") -> str:
    """按层级和序号生成前缀。"""
    if style == "sysu":
        if level == 1:
            return "%s、" % cn_number(n)
        if level == 2:
            return "（%s）" % cn_number(n)
        if level == 3:
            return "%d、" % n
        return "（%d）" % n
    if style == "arabic":
        return "%d. " % n
    return "%s、" % cn_number(n)


def _remove_auto_numbering(p) -> bool:
    """去掉段落的自动编号，返回是否真的去掉了。"""
    ppr = p._element.pPr
    if ppr is None:
        return False
    removed = False
    numpr = ppr.find(qn("w:numPr"))
    if numpr is not None:
        ppr.remove(numpr)
        removed = True
    # 行内 run 也可能带 numPr（少见但存在）
    for rpr in p._element.findall(".//" + qn("w:rPr")):
        n = rpr.find(qn("w:numPr"))
        if n is not None:
            rpr.remove(n)
            removed = True
    return removed


def renumber(doc, paras: list[dict], preset_region_of: dict[int, str],
             style: str = "sysu", dry_run: bool = False) -> dict:
    """
    按层级给标题重新编文字序号，并去掉自动编号。

    paras: region_detector.detect_regions 的段落列表
    preset_region_of: {段落索引: 区域key}，允许人工覆盖识别结果
    返回统计与明细。
    """
    LEVEL_OF = {"h1": 1, "h2": 2, "h3": 3, "h4": 4}
    counters = {1: 0, 2: 0, 3: 0, 4: 0}
    details = []
    n_stripped_numpr = 0
    n_renumbered = 0

    for info in paras:
        if info["empty"]:
            continue
        idx = info["idx"]
        region = preset_region_of.get(idx, info["region"])
        level = LEVEL_OF.get(region)
        if level is None:
            continue

        p = doc.paragraphs[idx]
        text = info["text"]

        # 前置/后置部分不编号
        norm = re.sub(r"\s+", "", text)
        if any(norm.startswith(k) or norm.startswith(k.upper())
               for k in NO_NUMBER_TITLES) and len(norm) <= 12:
            if _remove_auto_numbering(p) and not dry_run:
                n_stripped_numpr += 1
            details.append({"idx": idx, "level": level, "action": "skip(前置部分)",
                            "before": text[:30], "after": text[:30]})
            continue

        # 真正的章标题才重置下属计数器
        if level == 1:
            counters[1] += 1
            for k in (2, 3, 4):
                counters[k] = 0
        else:
            # 若父级从未编号过，仍从 1 开始，避免出现「（0）」
            counters[level] += 1
            for k in range(level + 1, 5):
                counters[k] = 0

        body = strip_number_prefix(text)
        if not body:
            continue
        new_text = make_prefix(level, counters[level], style) + body

        had_numpr = _remove_auto_numbering(p)
        if had_numpr:
            n_stripped_numpr += 1

        if not dry_run and new_text != text:
            # 注意：python-docx 每次访问 p.runs 都会新建一批 Run 包装对象，
            # 所以不能用 `r is not first` 判断——那个身份比较永远为真，
            # 会把 first 自己也放进 others，导致刚写入的文字立刻被清空。
            # 必须用底层 XML 元素来比较。
            el_of = lambda r: r._element
            first = None
            for run in p.runs:
                if run.text and run.text.strip():
                    first = run
                    break
            if first is not None:
                first_el = el_of(first)
                others = [r for r in p.runs
                          if el_of(r) is not first_el and r.text and r.text.strip()]
                first.text = new_text
                for r in others:
                    r.text = ""
                n_renumbered += 1

        details.append({
            "idx": idx, "level": level,
            "action": "renumber" if new_text != text else "unchanged",
            "before": text[:34], "after": new_text[:34],
            "had_auto_numbering": had_numpr,
        })

    return {
        "style": style,
        "stripped_auto_numbering": n_stripped_numpr,
        "renumbered": n_renumbered,
        "details": details,
    }


def renumber_report(result: dict) -> str:
    lines = ["编号处理（体系：%s）" % result["style"]]
    lines.append("  去掉自动编号的段落: %d" % result["stripped_auto_numbering"])
    lines.append("  重写文字编号的段落: %d" % result["renumbered"])
    changed = [d for d in result["details"] if d["action"] == "renumber"]
    if changed:
        lines.append("")
        lines.append("  编号变化明细：")
        for d in changed[:40]:
            mark = " [原为自动编号]" if d.get("had_auto_numbering") else ""
            lines.append("     #%-4d L%d  %s  ->  %s%s" % (
                d["idx"], d["level"], d["before"], d["after"], mark))
        if len(changed) > 40:
            lines.append("     ... 另有 %d 条" % (len(changed) - 40))
    return "\n".join(lines)
