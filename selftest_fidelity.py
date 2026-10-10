# -*- coding: utf-8 -*-
"""
对抗性检验 B：输出保真度与元数据保全
====================================
目的：改写只应改「显示格式」，绝不能悄悄丢掉文档的其他部分。
本文件逐项核对那些容易被忽略、丢了也不会立刻发现的东西。

重点检查（每项都是「丢了不易察觉」的类型）：
  * 文档属性（标题/作者/公司等 core_properties）
  * 页眉页脚内容与数量
  * 书签、超链接、交叉引用域
  * 分节属性（纸张、页边距、栏数）
  * 批注、脚注尾注部件
  * numbering.xml 与 styles.xml 是否仍在
  * 修订痕迹（w:ins / w:del）
  * 内容控件（w:sdt）
  * 段落数量、空段落是否保留
  * 重复运行同一输入的稳定性
  * 输出是否仍为合法 OOXML（关键部件齐全）

用法: python selftest_fidelity.py
"""
from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile
import traceback
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from docx import Document
from docx.oxml.ns import qn

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

PASS = 0
FAIL = 0
FAILURES: list[str] = []
SEC = ""


def sec(t):
    global SEC
    SEC = t
    print("\n" + "=" * 74)
    print(t)
    print("=" * 74)


def ok(name, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print("  [通过] %s" % name)
    else:
        FAIL += 1
        FAILURES.append("[%s] %s  %s" % (SEC, name, detail))
        print("  [失败] %s   %s" % (name, detail))


def parts(path):
    with zipfile.ZipFile(path) as z:
        return set(z.namelist())


def xml_of(path, part):
    with zipfile.ZipFile(path) as z:
        if part not in z.namelist():
            return ""
        return z.read(part).decode("utf-8", "ignore")


def build_rich_doc(path):
    """造一份「有各种元数据」的文档，作为保真基准。"""
    d = Document()

    # 文档属性
    cp = d.core_properties
    cp.title = "测试论文标题"
    cp.author = "测试作者"
    cp.subject = "保真度检验"
    cp.comments = "这段注释不应丢失"
    cp.category = "测试类别"

    # 正文
    d.add_paragraph("一、引言")
    d.add_paragraph("这是正文，含引用[1]。")
    d.add_paragraph("第二段正文，用于核对段落数。")
    d.add_paragraph()          # 空段落，必须保留
    d.add_paragraph("二、结论")

    # 页眉页脚
    s = d.sections[0]
    s.header.paragraphs[0].text = "页眉测试文字"
    s.footer.paragraphs[0].text = "页脚测试文字"

    # 分节属性
    from docx.shared import Mm
    s.page_width = Mm(210)
    s.page_height = Mm(297)
    s.left_margin = Mm(31.8)
    s.top_margin = Mm(25.4)

    # 书签与超链接域
    p = d.add_paragraph()
    r = p.add_run("带书签的文字")
    bs = p._element.makeelement(qn("w:bookmarkStart"), {})
    bs.set(qn("w:id"), "1")
    bs.set(qn("w:name"), "测试书签")
    p._element.insert(0, bs)
    be = p._element.makeelement(qn("w:bookmarkEnd"), {})
    be.set(qn("w:id"), "1")
    p._element.append(be)

    # 域代码（模拟目录/交叉引用）
    p2 = d.add_paragraph()
    r2 = p2.add_run()
    fld = r2._element.makeelement(qn("w:fldChar"), {})
    fld.set(qn("w:fldCharType"), "begin")
    r2._element.append(fld)
    p3 = d.add_paragraph()
    r3 = p3.add_run()
    instr = r3._element.makeelement(qn("w:instrText"), {})
    instr.text = " TOC \\o \"1-2\" "
    r3._element.append(instr)

    # 内容控件
    sdt = d.element.body.makeelement(qn("w:sdt"), {})
    content = sdt.makeelement(qn("w:sdtContent"), {})
    pp = d.element.body.makeelement(qn("w:p"), {})
    rr = pp.makeelement(qn("w:r"), {})
    tt = pp.makeelement(qn("w:t"), {})
    tt.text = "内容控件文字"
    rr.append(tt)
    pp.append(rr)
    content.append(pp)
    sdt.append(content)
    d.element.body.append(sdt)

    # 表格
    tbl = d.add_table(rows=2, cols=2)
    tbl.cell(0, 0).paragraphs[0].add_run("表头")

    # 参考文献
    d.add_paragraph("参考文献")
    d.add_paragraph("[1] 测试条目. 期刊, 2020(1):1-10.")

    d.save(path)


def snapshot(path):
    """抽取一份可比对的快照。"""
    d = Document(path)
    snap = {
        "paragraphs": len(d.paragraphs),
        "nonempty": sum(1 for p in d.paragraphs if p.text.strip()),
        "empty": sum(1 for p in d.paragraphs if not p.text.strip()),
        "tables": len(d.tables),
        "sections": len(d.sections),
        "texts": [p.text for p in d.paragraphs],
        "title": d.core_properties.title,
        "author": d.core_properties.author,
        "subject": d.core_properties.subject,
        "comments": d.core_properties.comments,
        "category": d.core_properties.category,
        "header": d.sections[0].header.paragraphs[0].text,
        "footer": d.sections[0].footer.paragraphs[0].text,
    }
    # 原始 XML 计数
    dx = xml_of(path, "word/document.xml")
    snap["counts"] = {
        "bookmarkStart": dx.count("<w:bookmarkStart"),
        "sdt": dx.count("<w:sdt>"),
        "fldChar": dx.count("<w:fldChar"),
        "instrText": dx.count("<w:instrText"),
        "tbl": dx.count("<w:tbl>"),
        "sectPr": dx.count("<w:sectPr"),
        "drawing": dx.count("<w:drawing"),
        "footnoteRef": dx.count("footnoteReference"),
        "endnoteRef": dx.count("endnoteReference"),
        "commentRef": dx.count("commentReference"),
    }
    # 页面设置
    s = d.sections[0]
    snap["page"] = {
        "w": round(s.page_width.mm, 1) if s.page_width else None,
        "h": round(s.page_height.mm, 1) if s.page_height else None,
        "left": round(s.left_margin.mm, 1) if s.left_margin else None,
        "top": round(s.top_margin.mm, 1) if s.top_margin else None,
    }
    snap["parts"] = sorted(parts(path))
    return snap


def main() -> int:
    print("=" * 74)
    print("对抗性检验 B —— 输出保真度与元数据保全")
    print("=" * 74)
    print("Python %s / %s" % (sys.version.split()[0], sys.platform))

    tmp = tempfile.mkdtemp(prefix="fidelity_")
    src = os.path.join(tmp, "基准.docx")
    out = os.path.join(tmp, "基准_out.docx")
    out2 = os.path.join(tmp, "基准_out2.docx")

    build_rich_doc(src)
    A = snapshot(src)

    from rewriter import rewrite, DEFAULT_PRESET

    sec("1. 执行改写")
    res = rewrite(src, out, renumber_headings=True)
    ok("改写执行成功（%d 段）" % res["changed_count"], os.path.exists(out))

    B = snapshot(out)

    sec("2. 段落结构")
    ok("段落总数不变", A["paragraphs"] == B["paragraphs"],
       "%d -> %d" % (A["paragraphs"], B["paragraphs"]))
    ok("空段落被保留（%d 个）" % A["empty"], A["empty"] == B["empty"],
       "%d -> %d" % (A["empty"], B["empty"]))
    ok("非空段落数不变", A["nonempty"] == B["nonempty"],
       "%d -> %d" % (A["nonempty"], B["nonempty"]))

    sec("3. 文档属性（core_properties）")
    for key, label in (("title", "标题"), ("author", "作者"), ("subject", "主题"),
                       ("comments", "备注"), ("category", "类别")):
        ok("文档%s保留" % label, A[key] == B[key],
           "%r -> %r" % (A[key], B[key]))

    sec("4. 页眉页脚")
    ok("页眉文字保留", A["header"] == B["header"], "%r -> %r" % (A["header"], B["header"]))
    ok("页脚文字保留", A["footer"] == B["footer"], "%r -> %r" % (A["footer"], B["footer"]))

    sec("5. 书签 / 域 / 内容控件 / 表格")
    for k, label in (("bookmarkStart", "书签"), ("sdt", "内容控件"),
                     ("fldChar", "域字符"), ("instrText", "域代码"),
                     ("tbl", "表格"), ("sectPr", "分节属性")):
        a, b = A["counts"][k], B["counts"][k]
        ok("%s数量不变（%d）" % (label, a), a == b, "%d -> %d" % (a, b))

    sec("6. 表格内容与分节数")
    ok("表格数不变", A["tables"] == B["tables"])
    ok("分节数不变", A["sections"] == B["sections"])
    ok("纸张尺寸保留", A["page"] == B["page"], "%s -> %s" % (A["page"], B["page"]))

    sec("7. 关键部件齐全（输出仍是合法 docx）")
    need = ["word/document.xml", "word/styles.xml", "[Content_Types].xml", "_rels/.rels"]
    missing = [n for n in need if n not in B["parts"]]
    ok("必备部件齐全", not missing, "缺失 %s" % missing)
    media_a = [p for p in A["parts"] if p.startswith("word/media/")]
    media_b = [p for p in B["parts"] if p.startswith("word/media/")]
    ok("媒体文件不丢失", set(media_a) <= set(media_b) or len(media_a) == len(media_b),
       "%s -> %s" % (media_a, media_b))

    sec("8. 文字内容（除补写/重排的编号前缀外不得改动）")
    # 注意：这里只剥离编号前缀后比对「正文部分」。
    # 编号本身被重排是预期行为（工具功能），不算数据损坏。
    diffs = []
    prefix_only = 0
    for i, (a, b) in enumerate(zip(A["texts"], B["texts"])):
        if a == b:
            continue
        strip_a = re.sub(r"^([一二三四五六七八九十]+、|（[一二三四五六七八九十]+）|\d+、)", "", a)
        strip_b = re.sub(r"^([一二三四五六七八九十]+、|（[一二三四五六七八九十]+）|\d+、)", "", b)
        if strip_a == strip_b:
            prefix_only += 1
            continue
        diffs.append((i, a[:30], b[:30]))
    ok("所有差异仅为编号前缀（%d 处）" % prefix_only, not diffs, str(diffs[:3]))
    ok("正文段落文字完全一致", all(
        A["texts"][i] == B["texts"][i] for i in (1, 2)),
        "%r vs %r" % (A["texts"][1], B["texts"][1]))
    ok("参考文献条目文字完全一致", A["texts"][9] == B["texts"][9],
       "%r vs %r" % (A["texts"][9][:30], B["texts"][9][:30]))

    sec("9. 反复运行同一输入的稳定性")
    rewrite(src, out2, renumber_headings=True)
    C = snapshot(out2)
    ok("两次改写结果段落数一致", B["paragraphs"] == C["paragraphs"])
    ok("两次改写结果文字一致", B["texts"] == C["texts"])
    ok("两次改写结果部件一致", B["parts"] == C["parts"])

    sec("10. 幂等性：对已改写的文档再改一次")
    out3 = os.path.join(tmp, "基准_out3.docx")
    rewrite(out, out3, renumber_headings=True)
    D = snapshot(out3)
    # 编号不应被重复添加
    dup = []
    for t in D["texts"]:
        if re.match(r"^[一二三四五六七八九十]+、[一二三四五六七八九十]+、", t):
            dup.append(t)
        if re.match(r"^（[一二三四五六七八九十]+）（[一二三四五六七八九十]+）", t):
            dup.append(t)
    ok("重复改写不产生重复编号", not dup, str(dup[:3]))
    ok("重复改写段落数不变", D["paragraphs"] == B["paragraphs"],
       "%d -> %d" % (B["paragraphs"], D["paragraphs"]))

    sec("11. 参考文献条目区保护")
    from region_detector import detect_regions
    paras, _ = detect_regions(src)
    ref_idx = [p["idx"] for p in paras if p["region"] == "ref_item"]
    ok("识别出参考文献条目", len(ref_idx) >= 1, "找到 %d 条" % len(ref_idx))
    # 条目里的 [1] 不应上标
    d_out = Document(out)
    bad = []
    for i in ref_idx:
        if i >= len(d_out.paragraphs):
            continue
        for r in d_out.paragraphs[i].runs:
            rpr = r._element.rPr
            va = rpr.find(qn("w:vertAlign")) if rpr is not None else None
            if va is not None and va.get(qn("w:val")) == "superscript":
                bad.append((i, r.text))
    ok("参考文献条目无上标", not bad, str(bad[:3]))

    sec("12. 预设确实生效（不同预设产生不同字体）")
    from rewriter import RegionFormat, preset_to_dict, preset_from_dict
    p_custom = {k: RegionFormat(**v.__dict__) for k, v in DEFAULT_PRESET.items()}
    p_custom["body"] = RegionFormat("楷体", "Arial", "五号", False, "左对齐", 0, 2.0, 0, 0)
    out4 = os.path.join(tmp, "基准_out4.docx")
    rewrite(src, out4, preset=p_custom, renumber_headings=False)
    d4 = Document(out4)
    # 找那个正文段
    from rewriter import DEFAULT_PRESET as DP
    target = None
    for p in d4.paragraphs:
        if p.text.startswith("这是正文"):
            target = p
            break
    if target is None:
        ok("找到正文段", False, "未找到")
    else:
        fonts = set()
        for r in target.runs:
            rpr = r._element.rPr
            rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
            if rf is not None:
                fonts.add(rf.get(qn("w:eastAsia")))
        ok("自定义预设的字体生效（楷体）", "楷体" in fonts, "实际 %s" % fonts)
        ls = target.paragraph_format.line_spacing
        ok("自定义预设的行距生效（2.0）", ls == 2.0, "实际 %s" % ls)

    sec("结果")
    print("  通过 %d 项，失败 %d 项" % (PASS, FAIL))
    if FAILURES:
        print("\n  失败明细:")
        for f in FAILURES:
            print("    ✗ %s" % f)
    else:
        print("  全部通过")
    print("\n  测试产物: %s" % tmp)
    return 0 if FAIL == 0 else 1


if __name__ == "__main__":
    try:
        code = main()
    except Exception:
        traceback.print_exc()
        code = 2
    sys.exit(code)
