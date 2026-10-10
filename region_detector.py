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
区域识别器
==========
把文档切分成"格式区"——哪些段是大标题、哪些是小标题、哪些是正文、哪些是参考文献。

这是改写的根基：先知道每段是什么，才能对它套用对应的字体字号。

与 audit_engine 的区别
----------------------
audit_engine 判断的是"这段是否符合某规范"；
本模块只回答"这段是什么区域"，不涉及合规判断，供改写器使用。
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Any

from docx import Document

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# ============================================================
# 区域类型定义（用户界面上的可调项，就是这些）
# ============================================================
REGION_TYPES = [
    ("title",        "论文/文档标题",   "封面或首页的大标题"),
    ("cover_field",  "封面信息",        "课程名称、姓名、学号等字段行"),
    ("h1",           "一级标题",        "一、引言 / 1. 绪论"),
    ("h2",           "二级标题",        "（一）基本情况 / 1.1 概述"),
    ("h3",           "三级标题",        "1、要点 / 1.1.1 细节"),
    ("body",         "正文",            "段落主体文字"),
    ("abstract",     "摘要内容",        "摘要正文"),
    ("keywords",     "关键词",          "关键词段落"),
    ("caption",      "图表题注",        "图1：xxx / 表1 xxx"),
    ("ref_head",     "参考文献标题",    "参考文献 / 注释"),
    ("ref_item",     "参考文献条目",    "每条文献著录"),
    ("toc",          "目录",            "目录标题与条目"),
    ("ack",          "致谢/附录",       "致谢与附录内容"),
    ("quote",        "引用/注释文字",   "较长的引文或注释"),
]
REGION_LABEL = {k: v for k, v, _ in REGION_TYPES}
REGION_DESC = {k: d for k, _, d in REGION_TYPES}

FIELD_LABELS = ("课程名称", "学院名称", "专业名称", "小组成员", "任课教师", "完成时间",
                "学位申请人", "导　　　师", "导 师", "专　　业", "学　　院", "姓　　名",
                "学　　号", "论文题目", "指导教师", "学生姓名", "班级")

# 标题编号模式 -> 层级。规范体系：一、二 / （一）（二） / 1、2 / （1）（2）
NUMBERING = [
    (1, re.compile(r"^[一二三四五六七八九十]+\s*[、\.]")),
    (2, re.compile(r"^[（(][一二三四五六七八九十]+[）)]")),
    (2, re.compile(r"^\d+\.\d+(?!\.\d)[\s　、]")),
    (3, re.compile(r"^\d+[、]")),
    (3, re.compile(r"^\d+\.[\s　]")),
    (4, re.compile(r"^[（(]\d+[）)]")),
]
REF_NUM = re.compile(r"^[\[［]\d+[\]］]")

# 摘要 / 关键词的常见标注形式：
#   摘要 / 【摘要】 / ［摘要］ / [摘要] / 摘  要 / 中文摘要 / Abstract
#   关键词 / 【关键词】 / 关键字 / Key words / 关键词：xxx
# 作者写标注的样式五花八门，识别必须容忍这些括号与空白。
# 实测教训：一份文档写的是「【摘要】：正式规则退场…」，旧规则只认裸的「摘要」，
# 结果摘要和关键词都没被识别出来，导致它们该用的楷体一次都没生效。
_BR_L = r"[\[［【（(〔]?"
_BR_R = r"[\]］】）)〕]?"
# 注意：这里**不能**用 $ 结尾。真实文本是「【摘要】：正式规则退场…」，
# 标注后紧跟冒号和正文，用 $ 会导致整条规则匹配失败（实测踩过这个坑）。
# 用 re.match 天然只锚定行首，所以标注匹配成功后剩下的正文自然被忽略。
ABSTRACT_HEAD = re.compile(
    r"^%s\s*(中文|英文)?\s*(摘要|Abstract)\s*%s\s*[:：]?" % (_BR_L, _BR_R), re.I)
# 英文摘要标题与内容同段（「Abstract: This paper...」）——标注后紧跟冒号与正文。
# 中文摘要已在 ABSTRACT_HEAD 里覆盖；这里单列便于判定「同段」情形。
ABSTRACT_INLINE_EN = re.compile(r"^\s*Abstract\s*[:：]", re.I)
KEYWORD_HEAD = re.compile(
    r"^%s\s*(关键词|关键字|Key\s*words?|Keywords?)\s*%s\s*[:：]?"
    % (_BR_L, _BR_R), re.I)

# 常见的无编号结构标题。这类词本身就是结构信号，不依赖任何格式。
# 纯文本稿（没有任何字体字号信息）里，这是识别它们的主要依据。
STRUCT_HEADINGS = {
    "引言", "导言", "绪论", "导论", "前言", "序言",
    "结语", "结论", "余论", "小结", "总结",
    "参考文献", "致谢", "附录", "注释",
}


def numbering_level(txt: str):
    for lvl, rx in NUMBERING:
        m = rx.search(txt)
        if m:
            return lvl, m.group(0)
    return None, None


# ============================================================
# 段落解析
# ============================================================
def _parse_run(r):
    rpr = r._element.rPr
    rf = rpr.find(W + "rFonts") if rpr is not None else None
    sz = rpr.find(W + "sz") if rpr is not None else None
    szcs = rpr.find(W + "szCs") if rpr is not None else None
    val = sz.get(W + "val") if sz is not None else (szcs.get(W + "val") if szcs is not None else None)
    return {
        "text": r.text,
        "ea": (rf.get(W + "eastAsia") if rf is not None else None),
        "ascii": (rf.get(W + "ascii") if rf is not None else None),
        "pt": (int(val) / 2.0) if val else None,
        "bold": (rpr is not None and rpr.find(W + "b") is not None),
        "italic": (rpr is not None and rpr.find(W + "i") is not None),
    }


def parse_paragraph(p, idx: int) -> dict[str, Any]:
    pf = p._element.pPr
    ind = pf.find(W + "ind") if pf is not None else None
    sp = pf.find(W + "spacing") if pf is not None else None
    jc = pf.find(W + "jc") if pf is not None else None
    ol = pf.find(W + "outlineLvl") if pf is not None else None
    ol_val = None
    if ol is not None:
        ol_val = ol.get(W + "val")
        if ol_val is None:
            ol_val = "0"   # OOXML 默认值：大纲级别 0

    runs = [x for x in (_parse_run(r) for r in p.runs) if x["text"] and x["text"].strip()]
    txt = p.text.strip()
    norm = re.sub(r"\s+", "", txt)
    cjk = len(re.findall(r"[\u3000-\u9fff\uff00-\uffef]", txt))

    def mode(key):
        c = Counter(x[key] for x in runs if x[key] is not None)
        return c.most_common(1)[0][0] if c else None

    bold_ratio = (sum(1 for x in runs if x["bold"]) / len(runs)) if runs else 0.0
    pt_mode = mode("pt")

    return {
        "idx": idx,
        "text": txt,
        "norm": norm,
        "len": len(txt),
        "empty": not txt,
        "style": p.style.name if p.style is not None else None,
        "ol": ol_val,
        "jc": (jc.get(W + "val") if jc is not None else None),
        "firstLineChars": (ind.get(W + "firstLineChars") if ind is not None else None),
        "firstLine": (ind.get(W + "firstLine") if ind is not None else None),
        "left": (ind.get(W + "left") if ind is not None else None),
        "line": (sp.get(W + "line") if sp is not None else None),
        "runs": runs,
        "n_runs": len(runs),
        "pt_mode": pt_mode,
        "ea_mode": mode("ea"),
        "bold_mode": bold_ratio >= 0.5,
        "bold_ratio": round(bold_ratio, 3),
        "pure_format": (len({x["pt"] for x in runs if x["pt"]}) <= 1) and bold_ratio in (0.0, 1.0),
        "cjk_ratio": round(cjk / len(txt), 3) if txt else 0.0,
        "has_drawing": bool(p._element.findall(".//" + W + "drawing")),
        "has_break": any(b.get(W + "type") == "page" for b in p._element.findall(".//" + W + "br")),
        # 输出字段
        "region": None, "conf": 0.0, "why": "",
    }


# ============================================================
# 区域判定
# ============================================================
HEAD_MAX_LEN = 50


def detect_regions(docx_path: str) -> tuple[list[dict], dict]:
    """返回 (段落列表, 汇总)。段落的 region/conf/why 字段被就地填入。"""
    doc = Document(docx_path)
    paras = [parse_paragraph(p, i) for i, p in enumerate(doc.paragraphs)]
    _classify(paras)
    summary = Counter(p["region"] for p in paras if not p["empty"])
    low = [p for p in paras if not p["empty"] and p["conf"] < 0.7]
    return paras, {
        "counts": dict(summary),
        "total_nonempty": sum(1 for p in paras if not p["empty"]),
        "low_confidence": low,
        "doc": doc,
    }


def _classify(paras: list[dict]) -> None:
    by_idx = {p["idx"]: p for p in paras}

    # 首个非空段落索引：纯文本稿没有格式可依，题目只能靠位置判定
    first_body_idx = next((p["idx"] for p in paras if not p["empty"]), -1)

    # ---- 封面区（保守识别，避免吞掉正文）----
    label_idx, seen = [], set()
    scan = 0
    for p in paras:
        if p["empty"]:
            continue
        scan += 1
        if scan > 40:
            break
        for lbl in FIELD_LABELS:
            if lbl in p["text"]:
                seen.add(lbl)
                label_idx.append(p["idx"])
                break

    cover_start = cover_end = -1
    if len(seen) >= 2 and label_idx:
        def ornament(pp):
            """刊物式排版块：大字号 / 居中 / 极短无句读。"""
            if pp["len"] == 0:
                return True
            if pp["pt_mode"] and pp["pt_mode"] >= 20.0:
                return True
            if pp["jc"] == "center":
                return True
            return pp["len"] <= 20 and not re.search(r"[。；！？，,]$", pp["text"])

        def protected(pp):
            """绝不能吞进封面的段落：章节标题、摘要、关键词等。
            必要性：若原稿全文格式统一（例如整篇 5.5pt 黑体），ornament()
            会对几乎所有段落返回真，封面区会一路蔓延，把摘要和关键词吃掉。"""
            n = pp["norm"]
            if ABSTRACT_HEAD.match(n):
                return True
            if KEYWORD_HEAD.match(n):
                return True
            if n in ("目录", "contents", "Contents", "CONTENTS"):
                return True
            if re.match(r"^(参考文献|注释)[:：]?$", n):
                return True
            if re.match(r"^(致谢|附录[A-ZＡ-Ｚ]?)$", n):
                return True
            # 带规范编号的标题（一、/（一）/1.）——封面不会有编号标题
            if numbering_level(pp["text"])[0] is not None and pp["len"] <= 30:
                return True
            # 显式大纲级别（非 9）——封面不参与大纲
            if pp["ol"] is not None and pp["ol"] != "9" and pp["len"] <= HEAD_MAX_LEN:
                return True
            return False

        start, steps = min(label_idx), 0
        while start - 1 in by_idx and steps < 14:
            prev = by_idx[start - 1]
            if protected(prev):
                break
            if prev["empty"] or any(l in prev["text"] for l in FIELD_LABELS) or ornament(prev):
                start -= 1
                steps += 1
            else:
                break
        end, steps = max(label_idx), 0
        while end + 1 in by_idx and steps < 14:
            nxt = by_idx[end + 1]
            if protected(nxt):
                break
            if nxt["empty"] or (ornament(nxt) and nxt["len"] <= 20):
                end += 1
                steps += 1
            else:
                break
        cover_start, cover_end = start, end

    zone = "body"
    for p in paras:
        if p["empty"]:
            p["region"], p["conf"], p["why"] = "blank", 1.0, "空段落"
            continue
        t, norm = p["text"], p["norm"]

        # ---- 封面 ----
        if cover_start <= p["idx"] <= cover_end:
            if any(l in t for l in FIELD_LABELS):
                p["region"], p["conf"], p["why"] = "cover_field", 0.95, "封面字段行"
            elif p["pt_mode"] and p["pt_mode"] >= 20.0:
                p["region"], p["conf"], p["why"] = "title", 0.85, "封面大字号标题"
            else:
                p["region"], p["conf"], p["why"] = "cover_field", 0.7, "封面区内文字"
            zone = "cover"
            continue

        # ---- 分区切换（先判前置部分，顺序重要）----
        if ABSTRACT_HEAD.match(norm):
            # 标注与正文同段时（如「【摘要】：xxx」），整段按摘要内容处理。
            # 英文同理：「Abstract: This paper…」标注后紧跟冒号和正文。
            inline = ABSTRACT_INLINE_EN.match(p["text"])
            if p["len"] > 12 or inline:
                zone = "abstract"
                p["region"], p["conf"], p["why"] = "abstract", 0.85, "摘要标注与内容同段"
                continue
            zone = "abstract"
            p["region"], p["conf"], p["why"] = "h1", 0.9, "摘要标题"
            continue
        if KEYWORD_HEAD.match(norm):
            # 同理：「【关键词】xx；yy」整段是关键词内容
            zone = "abstract"
            p["region"], p["conf"], p["why"] = "keywords", 0.9, "关键词"
            continue
        if norm in ("目录", "contents", "Contents", "CONTENTS"):
            zone = "toc"
            p["region"], p["conf"], p["why"] = "toc", 0.95, "目录标题"
            continue
        if re.match(r"^(参考文献|注释)[:：]?$", norm) and p["len"] <= 12:
            zone = "reference"
            p["region"], p["conf"], p["why"] = "ref_head", 0.95, "参考文献标题"
            continue
        if p["len"] <= 8 and re.match(r"^(致谢|附录[A-ZＡ-Ｚ]?)$", norm):
            zone = "ack"
            p["region"], p["conf"], p["why"] = "ack", 0.9, "致谢/附录标题"
            continue

        # ---- 前置区退出：遇到大标题即回正文 ----
        def is_major(pp):
            if pp["len"] > HEAD_MAX_LEN:
                return False
            if pp["ol"] == "0":
                return True
            return bool(pp["pt_mode"] and pp["pt_mode"] >= 14.0 and pp["ea_mode"] == "黑体")

        if zone in ("abstract", "toc") and is_major(p):
            zone = "body"

        # ---- 图表题注 ----
        if re.match(r"^(图|表|Fig|Table)\s*[\d\.]+", t):
            p["region"], p["conf"], p["why"] = "caption", 0.85, "图/表+编号开头"
            continue

        # ---- 各分区内 ----
        if zone == "toc":
            p["region"], p["conf"], p["why"] = "toc", 0.8, "目录区"
            continue
        if zone == "reference":
            p["region"], p["conf"], p["why"] = "ref_item", 0.85, "参考文献区"
            continue
        if zone == "abstract":
            if KEYWORD_HEAD.match(norm):
                p["region"], p["conf"], p["why"] = "keywords", 0.9, "关键词"
                continue
            p["region"], p["conf"], p["why"] = "abstract", 0.8, "摘要正文"
            continue
        if zone == "ack":
            p["region"], p["conf"], p["why"] = "ack", 0.75, "致谢/附录区"
            continue

        # ---- 标题判定 ----
        short = p["len"] <= HEAD_MAX_LEN
        lvl, num = numbering_level(t)

        # ============ A 组：结构信号（最可靠）============
        # 信号1：大纲级别（最强，但需配合长度）
        if p["ol"] is not None and p["ol"] != "9" and short and p["cjk_ratio"] >= 0.5:
            r = {0: "h1", 1: "h2", 2: "h3", 3: "h3"}.get(int(p["ol"]), "h3")
            p["region"], p["conf"], p["why"] = r, 0.9, "大纲级别 %s" % p["ol"]
            continue

        # 信号2：编号模式 —— 必须排在字号判据之前
        # ★ 踩过的坑：原先把「字号≥14pt + 黑体」放在编号之前，
        #   结果带「（一）」前缀的二级标题被字号规则抢先判成了一级标题。
        #   编号是作者显式标注的层级，比字号更明确，必须优先。
        if lvl is not None and short and p["cjk_ratio"] >= 0.4 \
                and not t.endswith(("。", "；", "！", "？", ".", "，", "：", "”")):
            r = {1: "h1", 2: "h2", 3: "h3", 4: "h3"}.get(lvl, "h3")
            p["region"], p["conf"], p["why"] = r, 0.85, "编号写法 %s" % num
            continue

        # 信号3：大字号 + 居中/黑体（无编号时才作为判据）
        if p["pt_mode"] and p["pt_mode"] >= 14.0 and short and p["cjk_ratio"] >= 0.5 \
                and (p["ea_mode"] == "黑体" or p["jc"] == "center"):
            p["region"], p["conf"], p["why"] = "h1", 0.8, "%.1fpt+%s" % (p["pt_mode"], p["ea_mode"] or p["jc"])
            continue
        # 信号4：全段统一加粗的极短行（无编号，如「研究背景」）
        if p["pure_format"] and p["bold_mode"] and p["len"] <= 20 \
                and not re.search(r"[。；！？，,]$", t):
            p["region"], p["conf"], p["why"] = "h2", 0.55, "全段加粗短行(无编号)"
            continue

        # ============ B 组：纯文字信号 ============
        # 必须独立成立。实测教训：一份文档 41 段完全没有格式信息
        # （无字号、无字体、无加粗、无大纲级别），A 组信号全部失效，
        # 结果所有标题和正文长得一模一样，一个标题都识别不出来。
        # 这种「纯文本稿」必须靠编号写法与文章结构来判。
        is_head_like = (
            short
            and not t.endswith(("。", "；", "！", "？", ".", "，", "：", "”"))
            and p["cjk_ratio"] >= 0.4
        )

        # 信号5：规范编号模式（一、/（一）/1、/（1））——不要求任何格式
        if is_head_like and lvl is not None and p["len"] <= 40:
            r = {1: "h1", 2: "h2", 3: "h3", 4: "h3"}.get(lvl, "h3")
            p["region"], p["conf"], p["why"] = r, 0.75, "编号写法 %s（无格式辅助）" % num
            continue
        # 信号6：结构词（引言/结语/结论/参考文献/致谢…）——不要求任何格式
        if p["len"] <= 12 and norm in STRUCT_HEADINGS:
            p["region"], p["conf"], p["why"] = "h1", 0.75, "结构词「%s」" % norm
            continue
        # 信号7：文档首段且像标题 => 论文题目
        if p["idx"] == first_body_idx and p["len"] <= 80 and lvl is None \
                and not t.endswith(("。", "；", "！", "？")):
            p["region"], p["conf"], p["why"] = "title", 0.7, "文档首段，判定为标题"
            continue

        # ---- 默认正文 ----
        p["region"], p["conf"], p["why"] = "body", 0.8, "默认正文"


def region_report(paras: list[dict], summary: dict) -> str:
    lines = []
    lines.append("区域识别结果：共 %d 个非空段落" % summary["total_nonempty"])
    lines.append("")
    for key, label, desc in REGION_TYPES:
        n = summary["counts"].get(key, 0)
        if n:
            lines.append("  %-14s %3d 段   %s" % (label, n, desc))
    lines.append("")
    lines.append("需人工复核（置信度 <0.7）：%d 段" % len(summary["low_confidence"]))
    for p in summary["low_confidence"][:15]:
        lines.append("    #%-4d [%s %.2f] %s" % (
            p["idx"], REGION_LABEL.get(p["region"], p["region"]), p["conf"], p["text"][:38]))
    return "\n".join(lines)


if __name__ == "__main__":
    import sys, json
    paras, summary = detect_regions(sys.argv[1])
    print(region_report(paras, summary))
    if "--dump" in sys.argv:
        print("\n=== 逐段明细 ===")
        for p in paras:
            if p["empty"]:
                continue
            print("#%-4d %-12s %.2f  %s" % (
                p["idx"], p["region"], p["conf"], p["text"][:50]))
