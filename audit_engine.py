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
论文格式合规体检引擎
====================
输入: 一份 .docx + 一份格式规范(JSON)
输出: 逐条违规清单 + 结构缺陷清单 + 合规率统计

设计原则
--------
1. 只报告确定性的偏离。规范本身没写清的地方报为「规范空白」，不当违规。
2. 分类时区分「纯格式段落」与「混排正文段落」——后者首句加粗极常见，
   若只看首个 run 的加粗会把大量正文误判成标题。
3. 每条结论都带定位(段落号)与证据(实际值 vs 要求值)，可复核。
"""
from __future__ import annotations

import json
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Any

from docx import Document

# 标题判据与改写引擎共用同一份实现。
# ★ 为什么必须共用：原先体检引擎自己写了一份更窄的副本，只认 outlineLvl 与
#   「≥14pt + 黑体」，于是摘要区里的纯文字编号标题（一、引言）不被认作标题，
#   分区无法退出，标题与其后全部正文都被归入摘要内容——体检查出的问题全是错的。
#   （issue #1 报的是改写引擎的同款缺陷，体检引擎有同样的问题。）
from region_detector import classify_heading

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# ============================================================
# 字号换算：OOXML 用半点(half-point)，规范用中文号数/磅值
# ============================================================
CN_SIZE_MAP = {
    "初号": 42.0, "小初": 36.0, "一号": 26.0, "小一": 24.0, "二号": 22.0, "小二": 18.0,
    "三号": 16.0, "小三": 15.0, "四号": 14.0, "小四": 12.0, "五号": 10.5, "小五": 9.0,
}


def g(el, name):
    return el.get(W + name) if el is not None else None


def half_points_to_pt(v):
    try:
        return int(v) / 2.0
    except (TypeError, ValueError):
        return None


# ============================================================
# 段落格式解析
# ============================================================
def parse_run(r) -> dict[str, Any]:
    rpr = r._element.rPr
    rf = rpr.find(W + "rFonts") if rpr is not None else None
    sz = rpr.find(W + "sz") if rpr is not None else None
    return {
        "text": r.text,
        "ea": g(rf, "eastAsia"),
        "ascii": g(rf, "ascii"),
        "hAnsi": g(rf, "hAnsi"),
        "pt": half_points_to_pt(g(sz, "val")),
        "bold": (rpr is not None and rpr.find(W + "b") is not None),
        "italic": (rpr is not None and rpr.find(W + "i") is not None),
    }


def parse_para(p, idx: int) -> dict[str, Any]:
    pf = p._element.pPr
    ind = pf.find(W + "ind") if pf is not None else None
    sp = pf.find(W + "spacing") if pf is not None else None
    jc = pf.find(W + "jc") if pf is not None else None
    ol = pf.find(W + "outlineLvl") if pf is not None else None
    # 关键：OOXML 中 w:outlineLvl 的 w:val 默认值为 0。
    # 形如 <w:outlineLvl/> 的元素表示「大纲级别 0」，即一级标题；
    # 若把「无 val」当成 None 跳过，会漏掉全部这类标题（实测漏检 32 个）。
    ol_val = None
    if ol is not None:
        ol_val = ol.get(W + "val")
        if ol_val is None:
            ol_val = "0"
    runs = [parse_run(r) for r in p.runs]
    runs = [r for r in runs if r["text"] and r["text"].strip()]
    txt = p.text.strip()

    # 段落级主流格式：全段同格式才算「纯格式段落」
    def mode(key):
        c = Counter(r[key] for r in runs if r[key] is not None)
        return c.most_common(1)[0][0] if c else None

    pt_set = sorted({r["pt"] for r in runs if r["pt"] is not None})
    bold_ratio = (sum(1 for r in runs if r["bold"]) / len(runs)) if runs else 0.0

    # 归一化文本：用于关键词匹配（去掉"摘  要"这类排版空格）
    norm = re.sub(r"\s+", "", txt)
    # 中文/日文字符占比：标题通常全中文，英文摘要标题以拉丁字母为主
    cjk = len(re.findall(r"[\u3000-\u9fff\uff00-\uffef]", txt))
    cjk_ratio = (cjk / len(txt)) if txt else 0.0

    return {
        "idx": idx,
        "text": txt,
        "norm": norm,
        "len": len(txt),
        "cjk_ratio": round(cjk_ratio, 3),
        "empty": not txt,
        "style": p.style.name if p.style is not None else None,
        "ol": ol_val,
        "jc": g(jc, "val"),
        "left": g(ind, "left"), "leftChars": g(ind, "leftChars"),
        "firstLine": g(ind, "firstLine"), "firstLineChars": g(ind, "firstLineChars"),
        "line": g(sp, "line"), "lineRule": g(sp, "lineRule"),
        "runs": runs,
        "n_runs": len(runs),
        "pt_mode": mode("pt"),
        "ea_mode": mode("ea"),
        "bold_mode": (bold_ratio >= 0.5),
        "bold_ratio": round(bold_ratio, 3),
        "pure_format": (len(pt_set) <= 1 and (bold_ratio in (0.0, 1.0))),
        "has_drawing": p._element.findall(".//" + W + "drawing") != [],
        "has_field": p._element.findall(".//" + W + "instrText") != [],
        "has_break": p._element.findall(".//" + W + "br") != [],
    }


# ============================================================
# 分区分区状态机
# ============================================================
ZONE_KEYWORDS = {
    "abstract": ["摘要", "abstract"],
    "reference": ["参考文献", "注释"],
    "ack": ["致谢", "附录"],
    "toc": ["目录"],
}

FIELD_LABELS = ("课程名称", "学院名称", "专业名称", "小组成员", "任课教师", "完成时间",
                "学位申请人", "导　　　师", "导 师", "专　　业", "学　　院", "姓　　名",
                "学　　号", "论文题目", "指导教师")

NUMBERING = [
    (1, re.compile(r"^[一二三四五六七八九十]+[、\.]")),
    (2, re.compile(r"^[（(][一二三四五六七八九十]+[）)]")),
    (3, re.compile(r"^\d+[、]")),
    (3, re.compile(r"^\d+\.[\s　]")),
    (2, re.compile(r"^\d+\.\d+[\s　]")),
    (4, re.compile(r"^[（(]\d+[）)]")),
]

REF_NUM = re.compile(r"^[\[［]\d+[\]］]")


def numbering_level(txt: str):
    for lvl, rx in NUMBERING:
        if rx.search(txt):
            return lvl, rx.search(txt).group(0)
    return None, None


def classify(paras: list[dict], spec: dict) -> None:
    """就地写入 cat / conf / why / zone 字段。"""
    zone = "cover"
    seen_abstract = seen_reference = False
    ref_buffer: list[dict] = []

    # ---- 预扫描：封面边界 ----
    # 封面识别必须保守，否则会吞掉正文里出现「学院：」字样的段落。
    # 实测教训：宽松规则曾把一份 106 段文档压缩到只剩 3 段参与判定。
    # 规则：找出字段标签簇，再向前后小幅扩展「刊物式」排版块，并设硬上限。
    COVER_SCAN_LIMIT = 40      # 只在文档前 40 个非空段内寻找
    COVER_MIN_LABELS = 2       # 至少 2 个不同字段标签
    COVER_EXTEND_MAX = 14      # 扩展硬上限，防止吞掉正文
    label_idx = []
    seen_labels: set[str] = set()
    scan = 0
    for p in paras:
        if p["empty"]:
            continue
        scan += 1
        if scan > COVER_SCAN_LIMIT:
            break
        if any(lbl in p["text"] for lbl in FIELD_LABELS):
            seen_labels.add(next(lbl for lbl in FIELD_LABELS if lbl in p["text"]))
            label_idx.append(p["idx"])

    cover_end = -1
    if len(seen_labels) >= COVER_MIN_LABELS and label_idx:
        first_label, last_label = min(label_idx), max(label_idx)

        def ornament(pp):
            """刊物式排版块：大字号 / 居中 / 极短无句读。"""
            if pp["len"] == 0:
                return True
            if pp["pt_mode"] and pp["pt_mode"] >= 20.0:
                return True
            if pp["jc"] == "center":
                return True
            return pp["len"] <= 20 and not re.search(r"[。；！？，,]$", pp["text"])

        by_idx = {p["idx"]: p for p in paras}
        start = first_label
        steps = 0
        while start - 1 in by_idx and steps < COVER_EXTEND_MAX:
            prev = by_idx[start - 1]
            if prev["empty"]:
                start -= 1
                steps += 1
                continue
            if any(lbl in prev["text"] for lbl in FIELD_LABELS) or ornament(prev):
                start -= 1
                steps += 1
                continue
            break
        end = last_label
        steps = 0
        while end + 1 in by_idx and steps < COVER_EXTEND_MAX:
            nxt = by_idx[end + 1]
            if nxt["empty"]:
                end += 1
                steps += 1
                continue
            if ornament(nxt) and nxt["len"] <= 20:
                end += 1
                steps += 1
                continue
            break
        cover_start, cover_end = start, end
    else:
        cover_start = cover_end = -1

    for i, p in enumerate(paras):
        t = p["text"]
        p["zone"] = zone
        if p["empty"]:
            p["cat"], p["conf"], p["why"] = "空段落", 1.0, "无文本"
            continue

        # ---------- 封面 ----------
        if cover_start <= p["idx"] <= cover_end:
            p["cat"] = "封面"
            p["conf"] = 0.9
            p["why"] = "位于封面字段区(%d-%d)" % (cover_start, cover_end)
            zone = "cover"
            continue

        # ---------- 分区切换 ----------
        # ---------- 前置部分（不参与编号体系，单独归类别，避免污染结构检查） ----------
        if re.match(r"^(中文)?摘要$", p["norm"]) or p["norm"] in ("Abstract", "ABSTRACT"):
            zone = "abstract"
            seen_abstract = True
            p["cat"], p["conf"], p["why"] = "中文摘要标题", 0.95, "归一化文本=摘要"
            continue
        if p["norm"] in ("目录", "contents", "Contents", "CONTENTS"):
            zone = "toc"
            p["cat"], p["conf"], p["why"] = "目录标题", 0.95, "归一化文本=目录"
            continue
        if re.match(r"^(参考文献|注释)[:：]?$", p["norm"]) and p["len"] <= 12:
            zone = "reference"
            seen_reference = True
            p["cat"], p["conf"], p["why"] = "注释参考文献标题", 0.95, "归一化文本=参考文献/注释"
            continue
        if p["len"] <= 8 and re.match(r"^(致谢|附录[A-ZＡ-Ｚ]?)$", p["norm"]):
            zone = "ack"
            p["cat"], p["conf"], p["why"] = "致谢附录标题", 0.9, "归一化文本=致谢/附录"
            continue

        # ---------- 分区退出：前置部分必须能被标题打断 ----------
        # 缺少这段退出逻辑会导致摘要区一路吞掉全文（实测把 88 段正文误判为摘要内容）。
        #
        # ★ 这里复用 region_detector 的 classify_heading，不再自己写一份判据。
        #   踩过的坑（issue #1）：原先这里是一份独立且更窄的副本
        #   （只看 outlineLvl 与「≥14pt + 黑体」），于是摘要区里遇到
        #   「一、引言」这类只有文字编号的标题时不认，标题和其后**全部正文**
        #   都被归入摘要内容，体检查出的问题也就全是错的。
        #   改写引擎与体检引擎必须共用同一份标题判据，否则两者会各自漂移。
        #   （first_body_idx 传 -1：本函数没有「文档首段即标题」这条信号，
        #     而返回 title 的情况本来也不作为退出信号。）
        lvl_for_exit, num_for_exit = numbering_level(t)
        _h = classify_heading(p, lvl_for_exit, num_for_exit, p["len"] <= 50,
                              t, p["norm"], -1)
        if zone in ("abstract", "toc") and _h is not None and _h[0] != "title":
            zone = "body"

        # ---------- 图表题注 ----------
        if re.match(r"^(图|表|Fig|Table)\s*[\d\.]+", t):
            p["cat"], p["conf"], p["why"] = "图表题注", 0.85, "以图/表+编号开头"
            continue

        # ---------- 目录区 ----------
        if zone == "toc":
            p["cat"], p["conf"], p["why"] = "目录其他内容", 0.7, "处于目录区"
            continue

        # ---------- 参考文献区 ----------
        if zone == "reference":
            p["cat"], p["conf"], p["why"] = "注释参考文献内容", 0.8, "处于参考文献区"
            continue

        # ---------- 摘要区 ----------
        if zone == "abstract":
            if re.match(r"^(关键词|关键字|Key\s*words?)\s*[:：]", t, re.I):
                p["cat"], p["conf"], p["why"] = "中文关键词段落", 0.85, "文本=关键词"
                continue
            if re.match(r"^[A-Za-z\s,\.\-]{20,}", t) and len(re.findall(r"[A-Za-z]", t)) > p["len"] * 0.5:
                p["cat"], p["conf"], p["why"] = "英文摘要内容", 0.7, "以拉丁字母为主体"
                continue
            p["cat"], p["conf"], p["why"] = "中文摘要内容", 0.75, "处于摘要区"
            continue

        # ---------- 保留纯格式的纯格式判断 ----------
        ea, pt, bold = p["ea_mode"], p["pt_mode"], p["bold_mode"]

        # 标题必须是短行。这是排除「混排长正文」最有效的单条规则：
        # 样本文档里 #82/#84/#130/#131 落在 outlineLvl=2 区段内，但文本 200+ 字，
        # 只看大纲级别会把它们整段吃掉。
        HEAD_MAX_LEN = 50
        short = p["len"] <= HEAD_MAX_LEN

        # 标题候选 1：显式大纲级别 —— 最强的结构信号。
        # 注意 outlineLvl=9 是 Word 表示「不进入大纲」的约定值（正文），必须排除；
        # 0/1/2/3 才是真正的标题级别。此信号优先于字号，否则会漏掉所有
        # 12pt 的次级标题（实测漏检 32 个）。
        if (p["ol"] is not None and p["ol"] not in ("9",) and short
                and p["cjk_ratio"] >= 0.5):
            p["cat"], p["conf"], p["why"] = "各部分标题", 0.9, "outlineLvl=%s" % p["ol"]
            p["numbering_level"] = numbering_level(t)[0]
            p["outline_level"] = int(p["ol"])
            continue
        # 标题候选 2：字号 >= 14pt 且（黑体 或 居中）
        if (pt is not None and pt >= 14.0 and short
                and (ea == "黑体" or p["jc"] == "center") and p["cjk_ratio"] >= 0.5):
            p["cat"], p["conf"], p["why"] = "各部分标题", 0.85, "%.1fpt+%s" % (pt, ea or p["jc"])
            p["numbering_level"] = numbering_level(t)[0]
            continue
        # 标题候选 3：全段统一加粗、短文本、纯格式段落、且带规范编号模式
        lvl, num = numbering_level(t)
        if (p["pure_format"] and bold and p["len"] <= 30 and lvl is not None
                and not t.endswith(("。", "；", "！", "？", ".", "，", "："))):
            p["cat"], p["conf"], p["why"] = "各部分标题", 0.8, "纯格式加粗短文本+编号%s" % num
            p["numbering_level"] = lvl
            continue
        # 标题候选 4：全段统一加粗、纯格式、极短、无编号（如「研究背景」）
        if (p["pure_format"] and bold and p["len"] <= 20
                and not re.search(r"[。；！？，,]$", t)):
            p["cat"], p["conf"], p["why"] = "各部分标题", 0.55, "纯格式加粗极短文本(无编号)"
            p["numbering_level"] = lvl
            continue

        # ---------- 其余归正文 ----------
        p["cat"], p["conf"], p["why"] = "正文", 0.8, "默认正文"
        p["is_bold_prefix"] = (0.0 < p["bold_ratio"] < 1.0)


# ============================================================
# 合规判定
# ============================================================
@dataclass
class Violation:
    idx: int
    cat: str
    text: str
    kind: str
    actual: str
    expected: str
    severity: str = "error"
    note: str = ""


def check_para(p: dict, spec: dict) -> list[Violation]:
    cat = p["cat"]
    els = spec["elements"]
    if cat not in els:
        return []
    want = els[cat]
    v: list[Violation] = []
    snippet = p["text"][:40]

    def add(kind, actual, expected, note=""):
        v.append(Violation(p["idx"], cat, snippet, kind, actual, expected, "error", note))

    # 字号
    if p["pt_mode"] is not None and "pt" in want:
        if abs(p["pt_mode"] - want["pt"]) > 0.01:
            add("字号", "%.1fpt" % p["pt_mode"], "%.1fpt" % want["pt"],
                "%s" % want.get("raw", ""))
    # 混排段落字号不一致
    pts = sorted({r["pt"] for r in p["runs"] if r["pt"] is not None})
    if len(pts) > 1:
        v.append(Violation(p["idx"], cat, snippet, "段内字号不一致",
                           "/".join("%.1f" % x for x in pts), "全段统一", "warn",
                           "同一段落内出现多个字号，通常来自复制粘贴"))
    # 中文字体
    if p["ea_mode"] is not None and "font_ea" in want and p["ea_mode"] != want["font_ea"]:
        add("中文字体", p["ea_mode"], want["font_ea"], want.get("raw", ""))
    # 加粗
    if "bold" in want and p["pure_format"] and p["bold_mode"] != want["bold"]:
        add("加粗", "是" if p["bold_mode"] else "否", "是" if want["bold"] else "否", "")
    # 缩进（仅正文/内容类要求首行缩进2字符）
    if cat in ("正文", "致谢附录内容", "中文摘要内容"):
        if p["firstLineChars"] not in ("200", None):
            if p["firstLineChars"] != "200":
                add("首行缩进", "%s字符" % (int(p["firstLineChars"]) / 100), "2字符", "首行缩进2字符")
        elif p["firstLineChars"] is None and p["firstLine"] not in ("0", None):
            add("首行缩进", "未按字符设置(%s)" % p["firstLine"], "2字符(=200)", "建议改用字符单位以适应字号变化")
    if cat == "中文关键词段落":
        if p["firstLineChars"] not in (None, "0"):
            add("首行缩进", "%s字符" % (int(p["firstLineChars"]) / 100), "顶格0字符", "关键词应顶格")
        if p["text"].rstrip().endswith(("；", ";", "。", "，")):
            add("关键词标点", "末词后带标点", "末词后不加标点", "")
    return v


def check_reference_content(p: dict) -> list[Violation]:
    v = []
    if not REF_NUM.match(p["text"]):
        v.append(Violation(p["idx"], "注释参考文献内容", p["text"][:40], "缺方括号序号",
                           p["text"][:16], "[序号] 左顶格", "error",
                           "规范：参考文献序号左顶格，用数字加方括号表示"))
    if not p["text"].rstrip().endswith("."):
        v.append(Violation(p["idx"], "注释参考文献内容", p["text"][:40], "条目未以句点结束",
                           p["text"].rstrip()[-8:], "以『.』结束", "warn", ""))
    return v


# ============================================================
# 结构缺陷
# ============================================================
CN_ORDER = "一二三四五六七八九十"


def structural_checks(paras: list[dict], spec: dict) -> list[dict]:
    defects = []
    heads = [p for p in paras if p["cat"] == "各部分标题"]

    # 1. 一级标题编号连续性与缺失
    lvl1 = [p for p in heads if numbering_level(p["text"])[0] == 1]
    seq = [(p["idx"], numbering_level(p["text"])[1], p["text"][:26]) for p in lvl1]

    # 找出「字号>=14 且黑体居中」这类明显的一级标题但无编号者
    unnumbered_l1 = [
        p for p in heads
        if numbering_level(p["text"])[0] is None
        and p["pt_mode"] and p["pt_mode"] >= 14.0
    ]
    if seq:
        nums = [s[1][0] for s in seq if s[1] and s[1][0] in CN_ORDER]
        missing = [c for c in CN_ORDER[:len(nums)] if c not in nums]
        if missing:
            defects.append({
                "kind": "一级标题编号跳缺", "severity": "error",
                "detail": "已出现 %s，但缺少 %s" % (
                    "、".join(nums), "、".join(missing)),
                "evidence": "; ".join("#%d %s" % (a, b) for a, _, b in seq),
                "suggestion": "编号缺失有两种可能：作者漏编，或这些标题本属更低层级。需人工确认层级归属。",
            })
    if unnumbered_l1:
        defects.append({
            "kind": "一级标题缺编号", "severity": "error",
            "detail": "%d 个看起来是一级标题的段落完全没有编号" % len(unnumbered_l1),
            "evidence": "; ".join("#%d %s" % (p["idx"], p["text"][:22]) for p in unnumbered_l1[:8]),
            "suggestion": "需人工确认这些标题应编号为『一、二……』的哪一级。",
        })

    # 2. 二级括号中文编号重复序列
    lvl2 = [p for p in heads if numbering_level(p["text"])[0] == 2]
    groups = []
    cur = []
    for p in lvl2:
        n = numbering_level(p["text"])[1]
        if n and n[1] == "一" and cur:
            groups.append(cur)
            cur = []
        cur.append(p)
    if cur:
        groups.append(cur)
    if len(groups) > 1:
        defects.append({
            "kind": "二级编号序列重复", "severity": "error",
            "detail": "『（一）』式编号出现 %d 轮独立序列，编号重复" % len(groups),
            "evidence": " | ".join("轮%d: %s" % (i + 1, ", ".join(x["text"][:12] for x in gp[:3]))
                                   for i, gp in enumerate(groups)),
            "suggestion": "每章内的二级标题应各自从『（一）』重新开始；若跨章连续则编号需接续。需按章节确认。",
        })

    # 3. 参考文献无序号
    refs = [p for p in paras if p["cat"] == "注释参考文献内容"]
    if refs:
        no_num = [p for p in refs if not REF_NUM.match(p["text"])]
        if len(no_num) == len(refs):
            defects.append({
                "kind": "参考文献完全缺序号", "severity": "error",
                "detail": "全部 %d 条参考文献均无 [1] 式序号" % len(refs),
                "evidence": "; ".join("#%d %s" % (p["idx"], p["text"][:20]) for p in refs[:4]),
                "suggestion": "规范明确要求序号左顶格、数字加方括号。可直接自动补号并核对条目顺序。",
            })

    # 4. 标题带标点
    bad_punct = [p for p in heads if re.search(r"[。；，、]$", p["text"])]
    if bad_punct:
        defects.append({
            "kind": "标题末尾带标点", "severity": "warn",
            "detail": "%d 个标题以标点结尾" % len(bad_punct),
            "evidence": "; ".join("#%d %s" % (p["idx"], p["text"][:22]) for p in bad_punct[:6]),
            "suggestion": "规范：标题不使用标点符号。可自动去除。",
        })

    # 5. 混排段落（正文首句加粗）
    bold_prefix = [p for p in paras if p.get("is_bold_prefix")]
    if bold_prefix:
        defects.append({
            "kind": "正文含加粗片段", "severity": "warn",
            "detail": "%d 个正文段落内部混有加粗文字" % len(bold_prefix),
            "evidence": "; ".join("#%d %s" % (p["idx"], p["text"][:20]) for p in bold_prefix[:6]),
            "suggestion": "规范正文为小四号宋体，未授权加粗。若是有意的强调需人工确认，否则应统一取消。",
        })

    # 6. 断行符做换行
    brk = [p for p in paras if p["has_break"]]
    if brk:
        defects.append({
            "kind": "段内使用手动换行", "severity": "warn",
            "detail": "%d 个段落用 Shift+Enter 换行而非分段" % len(brk),
            "evidence": "; ".join("#%d %s" % (p["idx"], p["text"][:20]) for p in brk[:6]),
            "suggestion": "影响段落级排版与目录提取，建议转为独立段落。",
        })

    # 7. 规范要求的组成部分缺失
    # 仅在文档看起来确实是一篇论文时才检查参考文献，避免对规范文件、
    # 课程作业等非论文文档误报（实测规范文件自身会被误报）。
    body_like = sum(1 for p in paras if p["cat"] == "正文" and p["len"] > 80)
    have_ref = any(p["cat"] == "注释参考文献内容" for p in paras)
    if not have_ref and body_like >= 5:
        defects.append({
            "kind": "缺少参考文献区", "severity": "warn",
            "detail": "检出 %d 个长正文段落，但未见『参考文献』区" % body_like,
            "evidence": "全文未见『参考文献』标题或方括号序号条目",
            "suggestion": "若这是论文，规范要求包含参考文献；若是课程作业或说明书，可忽略此项。",
        })
    return defects


# ============================================================
# 版式检查
# ============================================================
def check_layout(doc, spec: dict) -> list[dict]:
    pg = spec["page"]
    out = []
    for si, s in enumerate(doc.sections):
        def mm(v):
            return round(v.mm, 1) if v is not None else None
        checks = [
            ("纸张宽度", mm(s.page_width), pg["width_mm"], pg["tolerance_mm"]),
            ("纸张高度", mm(s.page_height), pg["height_mm"], pg["tolerance_mm"]),
            ("上边距", mm(s.top_margin), pg["margin_top_mm"], pg["tolerance_mm"]),
            ("下边距", mm(s.bottom_margin), pg["margin_bottom_mm"], pg["tolerance_mm"]),
            ("左边距", mm(s.left_margin), pg["margin_left_mm"], pg["tolerance_mm"]),
            ("右边距", mm(s.right_margin), pg["margin_right_mm"], pg["tolerance_mm"]),
        ]
        for name, act, exp, tol in checks:
            if act is None:
                continue
            if abs(act - exp) > tol:
                out.append({"sec": si + 1, "item": name, "actual": "%.1fmm" % act,
                            "expected": "%.1fmm" % exp, "delta": round(act - exp, 1)})
    return out


def check_line_spacing(paras: list[dict], spec: dict) -> list[dict]:
    allowed = {str(x["line"]) for x in spec["line_spacing"]["allowed"]}
    bad = defaultdict(list)
    for p in paras:
        if p["cat"] in ("正文", "各部分标题", "注释参考文献内容", "致谢附录内容", "中文摘要内容"):
            if p["line"] is not None and p["line"] not in allowed:
                bad[p["line"]].append(p["idx"])
    return [{"line_value": k, "count": len(v), "paras": v[:10],
             "expected": "1.5倍(360) 或 1.25倍(300)"} for k, v in bad.items()]


# ============================================================
# 主流程
# ============================================================
def audit(docx_path: str, spec: dict) -> dict:
    doc = Document(docx_path)
    paras = [parse_para(p, i) for i, p in enumerate(doc.paragraphs)]
    classify(paras, spec)

    violations = []
    for p in paras:
        if p["empty"]:
            continue
        if p["cat"] == "注释参考文献内容":
            violations += check_reference_content(p)
        violations += check_para(p, spec)

    # 合规率：每类中被判定为完全合规的段落占比
    per_cat = defaultdict(lambda: {"total": 0, "clean": 0, "low_conf": 0})
    viol_by_idx = defaultdict(list)
    for v in violations:
        viol_by_idx[v.idx].append(v)
    for p in paras:
        if p["empty"] or p["cat"] not in spec["elements"]:
            continue
        c = per_cat[p["cat"]]
        c["total"] += 1
        if p["conf"] < 0.7:
            c["low_conf"] += 1
        if not [x for x in viol_by_idx[p["idx"]] if x.severity == "error"]:
            c["clean"] += 1

    return {
        "file": docx_path,
        "spec_name": spec["spec_name"],
        "paragraphs_total": len(doc.paragraphs),
        "paragraphs_nonempty": sum(1 for p in paras if not p["empty"]),
        "per_category": {
            k: {**v, "rate": round(v["clean"] / v["total"] * 100, 1) if v["total"] else None}
            for k, v in sorted(per_cat.items(), key=lambda kv: -kv[1]["total"])
        },
        "violations": violations,
        "structural_defects": structural_checks(paras, spec),
        "layout_issues": check_layout(doc, spec),
        "line_spacing_issues": check_line_spacing(paras, spec),
        "paragraphs": paras,
        "spec": spec,
    }


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        print("用法: python audit_engine.py <docx> <spec.json> [--json out.json]")
        sys.exit(1)
    docx_path, spec_path = sys.argv[1], sys.argv[2]
    with open(spec_path, encoding="utf-8") as f:
        spec = json.load(f)
    result = audit(docx_path, spec)

    print("=" * 78)
    print("合规体检: %s" % docx_path.split("\\")[-1])
    print("规范: %s" % result["spec_name"])
    print("=" * 78)
    print("非空段落 %d" % result["paragraphs_nonempty"])
    print("\n--- 各区合规率 ---")
    for cat, v in result["per_category"].items():
        print("  %-16s %3d 段  合规 %5s%%  低置信度 %d 段" % (
            cat, v["total"], v["rate"], v["low_conf"]))

    print("\n--- 结构缺陷 (%d) ---" % len(result["structural_defects"]))
    for d in result["structural_defects"]:
        print("  [%s] %s" % (d["severity"].upper(), d["kind"]))
        print("       %s" % d["detail"])
        if d.get("evidence"):
            print("       证据: %s" % d["evidence"][:150])
        print("       建议: %s" % d["suggestion"])

    print("\n--- 版式问题 (%d) ---" % len(result["layout_issues"]))
    for x in result["layout_issues"]:
        print("  第%d节 %s: %s (应 %s, 差 %+.1fmm)" % (
            x["sec"], x["item"], x["actual"], x["expected"], x["delta"]))

    print("\n--- 行距问题 (%d 种) ---" % len(result["line_spacing_issues"]))
    for x in result["line_spacing_issues"]:
        print("  line=%s 的段落 %d 个 (应 %s)" % (x["line_value"], x["count"], x["expected"]))

    print("\n--- 违规明细：按类型汇总 ---")
    by_kind = Counter(v.kind for v in result["violations"])
    for k, c in by_kind.most_common():
        print("  %-18s %d" % (k, c))
    print("\n  违规总数: %d" % len(result["violations"]))

    # 分类自检：异常长的"标题"说明分类器可能把长正文吃掉了
    long_heads = [p for p in result["paragraphs"]
                  if p["cat"] == "各部分标题" and p["len"] > 50]
    print("\n--- 分类器自检 ---")
    print("  被判为标题且长度 > 50 字符的段落: %d 个 (应为 0)" % len(long_heads))
    for p in long_heads[:5]:
        print("     #%d (%d字) %s" % (p["idx"], p["len"], p["text"][:40]))
    low = [p for p in result["paragraphs"]
           if not p["empty"] and p["cat"] in spec["elements"] and p["conf"] < 0.7]
    print("  低置信度分类: %d 个 (需人工复核)" % len(low))
    for p in low[:8]:
        print("     #%-4d [%s conf=%.2f] %s" % (p["idx"], p["cat"], p["conf"], p["text"][:40]))

    if "--json" in sys.argv:
        out = sys.argv[sys.argv.index("--json") + 1]
        slim = {k: v for k, v in result.items() if k != "paragraphs"}
        slim["violations"] = [v.__dict__ for v in result["violations"]]
        with open(out, "w", encoding="utf-8") as f:
            json.dump(slim, f, ensure_ascii=False, indent=2)
        print("\n已写出 JSON: %s" % out)


if __name__ == "__main__":
    main()
