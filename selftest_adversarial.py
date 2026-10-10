# -*- coding: utf-8 -*-
"""
对抗性检验 A：OOXML 边角情况与健壮性
====================================
目的：用畸形/异常的文档结构去撞程序，看它会崩、会静默出错、还是会给出可解释的结果。

与既有自检的区别
----------------
现有自检用的是「正常论文 + 常见变体」。本文件专门构造**不正常的输入**：
空文档、只有空白、无 run、超长标题、表格段落、文本框、重复编号、
嵌套内容控件、异常数字、损坏的样式引用、含图片段落、修订痕迹等。

判定标准：
  * 不抛未捕获异常（能给出可解释结果或明确跳过）
  * 不产生损坏的输出文件（输出必须能被重新打开）
  * 不静默改错（该跳过的要跳过，该保留的要保留）

用法: python selftest_adversarial.py
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
import traceback
from copy import deepcopy


# 保证中文输出在任何终端都不崩。
# Windows 上 stdout 默认编码可能是 cp1252，直接 print 中文会抛
# UnicodeEncodeError 并让脚本崩掉（GitHub Actions 的 windows runner 实测踩过）。
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
import zipfile

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


def safe(label, fn, *a, **kw):
    """执行一个操作，捕获异常。返回 (成功, 结果或异常)"""
    try:
        return True, fn(*a, **kw)
    except Exception as e:
        print("  [异常] %s -> %s: %s" % (label, type(e).__name__, str(e)[:120]))
        return False, e


def add_run(p, text, ea=None, pt=None, bold=False):
    r = p.add_run(text)
    rpr = r._element.get_or_add_rPr()
    if ea:
        rf = rpr.makeelement(qn("w:rFonts"), {})
        rf.set(qn("w:eastAsia"), ea)
        rpr.insert(0, rf)
    if pt:
        sz = rpr.makeelement(qn("w:sz"), {})
        sz.set(qn("w:val"), str(int(pt * 2)))
        rpr.append(sz)
    if bold:
        rpr.append(rpr.makeelement(qn("w:b"), {}))
    return r


def can_open(path):
    """输出文件是否能被重新打开且不抛异常。"""
    try:
        d = Document(path)
        _ = [p.text for p in d.paragraphs]
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
        return bad is None, d
    except Exception as e:
        return False, e


def run_pipeline(src, out, **kw):
    from rewriter import rewrite
    return rewrite(src, out, **kw)


# ============================================================
# 构造各种畸形输入
# ============================================================
def make_empty(path):
    Document().save(path)


def make_only_blanks(path):
    d = Document()
    for _ in range(5):
        d.add_paragraph()
    d.add_paragraph("   ")
    d.add_paragraph("\t")
    d.save(path)


def make_no_runs(path):
    """段落存在但没有任何 run（只有 pPr）。"""
    d = Document()
    d.add_paragraph("正常段落")
    p = d.add_paragraph()
    # 手工造一个空段落元素：清掉所有 run
    for r in list(p._element.findall(qn("w:r"))):
        p._element.remove(r)
    d.save(path)


def make_long_heading(path):
    d = Document()
    long_text = "一、" + "这是一段被误当成标题的超长文字" * 12
    p = d.add_paragraph()
    add_run(p, long_text, ea="黑体", pt=16)
    d.add_paragraph("正常正文。")
    d.save(path)


def make_duplicate_numbers(path):
    d = Document()
    for t in ("一、引言", "一、另一个一", "二、正文", "二、重复的二", "三、结论"):
        p = d.add_paragraph()
        add_run(p, t, ea="黑体", pt=14)
        d.add_paragraph("正文内容。")
    d.save(path)


def make_table_paragraphs(path):
    d = Document()
    d.add_paragraph("一、表格测试")
    tbl = d.add_table(rows=3, cols=2)
    tbl.cell(0, 0).paragraphs[0].add_run("表头一")
    tbl.cell(1, 0).paragraphs[0].add_run("数据 一、看起来像标题")
    tbl.cell(2, 1).paragraphs[0].add_run("[1] 看起来像引用")
    d.add_paragraph("表格后的正文。")
    d.save(path)


def make_with_image(path):
    """含图片的段落不应被破坏。"""
    d = Document()
    d.add_paragraph("一、含图文档")
    from docx.shared import Inches
    # 造一个最小 PNG
    png = bytes.fromhex(
        "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4"
        "890000000a49444154789c63000100000500010d0a2db40000000049454e44ae"
        "426082")
    tmp_png = os.path.join(os.path.dirname(path), "_t.png")
    with open(tmp_png, "wb") as f:
        f.write(png)
    p = d.add_paragraph()
    p.add_run("图前文字")
    try:
        p.add_run().add_picture(tmp_png, width=Inches(0.5))
    except Exception:
        pass
    d.add_paragraph("图后文字。")
    d.save(path)


def make_numbered_lists(path):
    """正文里的真正项目符号列表，不应被判成标题。"""
    d = Document()
    d.add_paragraph("一、列表测试")
    for t in ("第一项内容比较长，说明这确实是正文而不是标题",
              "第二项内容同样很长，也不该被当成标题"):
        d.add_paragraph(t, style="List Bullet")
    d.save(path)


def make_huge_font(path):
    """异常字号（超小/超大）不应导致崩溃。"""
    d = Document()
    d.add_paragraph("一、异常字号")
    p = d.add_paragraph()
    add_run(p, "极小字", ea="宋体", pt=1)
    p2 = d.add_paragraph()
    add_run(p2, "极大字", ea="宋体", pt=300)
    d.save(path)


def make_broken_style_ref(path):
    """引用一个不存在的样式 id。"""
    d = Document()
    d.add_paragraph("一、坏样式引用")
    p = d.add_paragraph()
    ppr = p._element.get_or_add_pPr()
    ps = ppr.makeelement(qn("w:pStyle"), {})
    ps.set(qn("w:val"), "不存在的样式ID")
    ppr.insert(0, ps)
    p.add_run("这段引用了不存在的样式。")
    d.save(path)


def make_nested_sdt(path):
    """嵌套的内容控件，里面有段落。"""
    d = Document()
    d.add_paragraph("一、内容控件测试")
    sdt = d.element.body.makeelement(qn("w:sdt"), {})
    content = sdt.makeelement(qn("w:sdtContent"), {})
    p = d.element.body.makeelement(qn("w:p"), {})
    r = p.makeelement(qn("w:r"), {})
    t = p.makeelement(qn("w:t"), {})
    t.text = "内容控件里的段落文字"
    r.append(t)
    p.append(r)
    content.append(p)
    sdt.append(content)
    d.element.body.append(sdt)
    d.add_paragraph("控件后的正文。")
    d.save(path)


def make_superscript_edge(path):
    """上标边角：连续标记、跨 run 标记、只有标记的段落。"""
    d = Document()
    d.add_paragraph("一、上标边角")
    d.add_paragraph("连续标记[1][2][3]应该都上标。")
    p = d.add_paragraph()
    p.add_run("跨 run 的[")
    p.add_run("4")
    p.add_run("]标记。")
    d.add_paragraph("[5]")
    d.add_paragraph("[6]只出现在参考文献之前，这里是正文。")
    d.add_paragraph("参考文献")
    d.add_paragraph("[1] 条目一.")
    d.save(path)


def make_all_in_one_paragraph(path):
    """整篇塞进一个段落（极端情况）。"""
    d = Document()
    d.add_paragraph("一、标题和正文挤在一起。这里是正文内容。二、又一个标题。更多正文。")
    d.save(path)


def make_mixed_script(path):
    """中英混排、含 emoji 与特殊符号。"""
    d = Document()
    d.add_paragraph("一、Mixed Script 测试")
    d.add_paragraph("English text with [1] citation and 中文混排，还有 emoji 🙂 与符号 §¶†。")
    d.add_paragraph("参考文献")
    d.add_paragraph("[1] Smith J. A study on mixed scripts[J]. Journal, 2020(1):1-10.")
    d.save(path)


CASES = [
    ("空文档", make_empty),
    ("仅空白段落", make_only_blanks),
    ("段落无run", make_no_runs),
    ("超长标题", make_long_heading),
    ("重复编号", make_duplicate_numbers),
    ("表格内段落", make_table_paragraphs),
    ("含图片", make_with_image),
    ("项目符号列表", make_numbered_lists),
    ("异常字号", make_huge_font),
    ("坏样式引用", make_broken_style_ref),
    ("嵌套内容控件", make_nested_sdt),
    ("上标边角", make_superscript_edge),
    ("全文单段", make_all_in_one_paragraph),
    ("中英混排", make_mixed_script),
]


def main() -> int:
    print("=" * 74)
    print("对抗性检验 A —— OOXML 边角情况与健壮性")
    print("=" * 74)
    print("Python %s / %s" % (sys.version.split()[0], sys.platform))

    tmp = tempfile.mkdtemp(prefix="adversarial_")
    from region_detector import detect_regions

    sec("1. 逐个畸形输入跑「识别 + 改写」")
    for name, builder in CASES:
        src = os.path.join(tmp, "%s.docx" % name)
        out = os.path.join(tmp, "%s_out.docx" % name)
        # 造文档
        good, err = safe("造文档:%s" % name, builder, src)
        if not good or not os.path.exists(src):
            ok("%s 构造成功" % name, False, repr(err)[:80])
            continue

        # 识别
        okr, res = safe("识别:%s" % name, detect_regions, src)
        if not okr:
            ok("%s 识别不抛异常" % name, False, repr(res)[:100])
            continue
        paras, summary = res
        n_regions = len([k for k, v in summary["counts"].items() if v])

        # 改写
        okw, res2 = safe("改写:%s" % name, run_pipeline, src, out,
                         renumber_headings=True)
        if not okw:
            ok("%s 改写不抛异常" % name, False, repr(res2)[:100])
            continue

        # 输出可打开
        can, doc_or_err = can_open(out)
        ok("%-8s 输出可打开（%d 段，%d 类区域）" % (name, summary["total_nonempty"], n_regions),
           can, repr(doc_or_err)[:100])

        # 原文件未被改动
        can_src, _ = can_open(src)
        ok("%-8s 原文件仍可打开" % name, can_src)

    sec("2. 空文档不应产生损坏输出")
    src = os.path.join(tmp, "空文档.docx")
    out = os.path.join(tmp, "空文档_out.docx")
    if os.path.exists(out):
        can, d = can_open(out)
        ok("空文档输出可打开", can, repr(d)[:80])
        if can:
            ok("空文档输出段落数合理", len(d.paragraphs) <= 2,
               "实际 %d 段" % len(d.paragraphs))

    sec("3. 表格内段落不应被当标题改写")
    src = os.path.join(tmp, "表格内段落.docx")
    out = os.path.join(tmp, "表格内段落_out.docx")
    if os.path.exists(out):
        d = Document(out)
        ok("表格数量保留", len(d.tables) == 1, "实际 %d" % len(d.tables))
        if d.tables:
            cell_text = d.tables[0].cell(1, 0).paragraphs[0].text
            ok("表格内文字未被改动", "数据" in cell_text, repr(cell_text[:30]))

    sec("4. 含图片文档的图片不应丢失")
    src = os.path.join(tmp, "含图片.docx")
    out = os.path.join(tmp, "含图片_out.docx")
    if os.path.exists(out):
        try:
            with zipfile.ZipFile(src) as za, zipfile.ZipFile(out) as zb:
                ma = [n for n in za.namelist() if n.startswith("word/media/")]
                mb = [n for n in zb.namelist() if n.startswith("word/media/")]
                same = set(ma) == set(mb)
                ok("图片文件全部保留（%d 个）" % len(ma), same,
                   "原 %s / 新 %s" % (ma, mb))
        except Exception as e:
            ok("图片检查", False, repr(e)[:80])

    sec("5. 上标边角情况")
    src = os.path.join(tmp, "上标边角.docx")
    out = os.path.join(tmp, "上标边角_out.docx")
    if os.path.exists(out):
        d = Document(out)

        def sup_of(i):
            p = d.paragraphs[i]
            out_ = []
            for r in p.runs:
                rpr = r._element.rPr
                va = rpr.find(qn("w:vertAlign")) if rpr is not None else None
                if va is not None and va.get(qn("w:val")) == "superscript":
                    out_.append(r.text)
            return out_

        s1 = "".join(sup_of(1))
        # 注意：相邻标记可能被合并进同一个 run（切分后自然相邻），
        # 所以用子串判断，不要用逐元素精确计数。
        has_all = all(m in s1 for m in ("[1]", "[2]", "[3]"))
        ok("连续标记 [1][2][3] 都上标", has_all, "上标内容 %r" % s1)
        s2 = "".join(sup_of(2))
        ok("跨 run 的 [4] 能上标", "[4]" in s2, "上标内容 %r" % s2)
        s3 = "".join(sup_of(3))
        ok("只有标记的段落 [5] 上标", "[5]" in s3, "上标内容 %r" % s3)
        s_last = "".join(sup_of(len(d.paragraphs) - 1))
        ok("参考文献条目 [1] 不上标", s_last == "", "上标内容 %r" % s_last)

    sec("6. 输出文件结构完整性（zip 无损坏）")
    bad_zips = []
    for fn in os.listdir(tmp):
        if fn.endswith("_out.docx"):
            p = os.path.join(tmp, fn)
            try:
                with zipfile.ZipFile(p) as z:
                    if z.testzip() is not None:
                        bad_zips.append(fn)
            except Exception:
                bad_zips.append(fn)
    n_out = len([f for f in os.listdir(tmp) if f.endswith("_out.docx")])
    ok("%d 个输出 zip 结构完好" % n_out, not bad_zips, str(bad_zips))

    sec("7. 没有任何输出把原文件覆盖掉")
    overwritten = []
    for name, _ in CASES:
        src = os.path.join(tmp, "%s.docx" % name)
        out = os.path.join(tmp, "%s_out.docx" % name)
        if os.path.exists(src) and os.path.exists(out):
            if os.path.samefile(src, out):
                overwritten.append(name)
    ok("输入输出不是同一文件", not overwritten, str(overwritten))

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
