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
合规体检 HTML 报告生成器
========================
把 audit_engine 的结果渲染成一份可离线打开、可打印的体检报告。
用法: python report_html.py <docx> <spec.json> [out.html]
"""
from __future__ import annotations

import html
import json
import sys
from collections import Counter, defaultdict
from datetime import datetime

from audit_engine import audit

CSS = """
*{box-sizing:border-box}
body{margin:0;padding:0;background:#f5f6f8;
 font:15px/1.75 -apple-system,"Segoe UI","Microsoft YaHei",sans-serif;color:#1c1e21}
.wrap{max-width:1080px;margin:0 auto;padding:32px 28px 80px}
h1{font-size:26px;margin:0 0 6px;letter-spacing:-.3px}
h2{font-size:19px;margin:38px 0 14px;padding-bottom:8px;border-bottom:2px solid #e4e6eb}
h3{font-size:16px;margin:24px 0 10px;color:#3a3d42}
.meta{color:#8a8f98;font-size:13px;margin-bottom:26px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(158px,1fr));gap:12px;margin:18px 0}
.card{background:#fff;border:1px solid #e4e6eb;border-radius:10px;padding:15px 16px}
.card .k{font-size:12px;color:#8a8f98;margin-bottom:4px}
.card .v{font-size:23px;font-weight:650;letter-spacing:-.5px}
.card.bad .v{color:#d93025}.card.warn .v{color:#e37400}.card.ok .v{color:#137333}
table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #e4e6eb;
 border-radius:10px;overflow:hidden;font-size:14px}
th,td{padding:9px 12px;text-align:left;border-bottom:1px solid #eef0f3;vertical-align:top}
th{background:#fafbfc;font-weight:600;font-size:13px;color:#5a5f66;white-space:nowrap}
tr:last-child td{border-bottom:none}
.pill{display:inline-block;padding:1px 9px;border-radius:20px;font-size:12px;font-weight:600}
.p-red{background:#fce8e6;color:#c5221f}.p-orange{background:#fef3e2;color:#b06000}
.p-green{background:#e6f4ea;color:#137333}.p-gray{background:#eef0f3;color:#5a5f66}
.p-blue{background:#e8f0fe;color:#1a56c4}
.bar{height:7px;background:#eef0f3;border-radius:4px;overflow:hidden;min-width:80px}
.bar>i{display:block;height:100%;border-radius:4px}
.defect{background:#fff;border:1px solid #e4e6eb;border-left:4px solid #d93025;
 border-radius:8px;padding:14px 16px;margin:12px 0}
.defect.warn{border-left-color:#e37400}
.defect .t{font-weight:650;margin-bottom:6px}
.defect .d{font-size:14px;color:#3a3d42;margin-bottom:8px}
.defect .ev{font:12.5px/1.6 ui-monospace,Consolas,monospace;background:#f7f8fa;
 padding:8px 10px;border-radius:6px;color:#5a5f66;word-break:break-all;margin-bottom:8px}
.defect .sug{font-size:13.5px;color:#1a56c4}
code{font:13px ui-monospace,Consolas,monospace;background:#f0f1f4;padding:1px 5px;border-radius:4px}
.foot{margin-top:44px;padding-top:18px;border-top:1px solid #e4e6eb;color:#8a8f98;font-size:12.5px}
.note{background:#e8f0fe;border-radius:8px;padding:13px 16px;font-size:13.5px;color:#174ea6;margin:14px 0}
.tabs{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap}
.legend{font-size:12.5px;color:#8a8f98;margin-top:8px}
"""


def esc(s):
    return html.escape(str(s))


def rate_class(r):
    if r is None:
        return "p-gray", "#8a8f98"
    if r >= 90:
        return "p-green", "#137333"
    if r >= 60:
        return "p-orange", "#e37400"
    return "p-red", "#d93025"


def build(result: dict, out_path: str) -> str:
    spec = result["spec"]
    paras = {p["idx"]: p for p in result["paragraphs"]}
    viols = result["violations"]
    defects = result["structural_defects"]

    # 汇总
    n_err = sum(1 for d in defects if d["severity"] == "error")
    n_warn = sum(1 for d in defects if d["severity"] != "error")
    total_checked = sum(v["total"] for v in result["per_category"].values())
    total_clean = sum(v["clean"] for v in result["per_category"].values())
    overall = (total_clean / total_checked * 100) if total_checked else 0.0

    # 违规按段落归并
    by_para = defaultdict(list)
    for v in viols:
        by_para[v.idx].append(v)

    h = []
    A = h.append
    A("<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'>")
    A("<meta name='viewport' content='width=device-width,initial-scale=1'>")
    A("<title>论文格式合规体检报告</title><style>%s</style></head><body><div class='wrap'>" % CSS)

    A("<h1>论文格式合规体检报告</h1>")
    A("<div class='meta'>受检文件：%s ｜ 适用规范：%s ｜ 生成时间：%s</div>" % (
        esc(result["file"].split("\\")[-1]), esc(spec["spec_name"]),
        datetime.now().strftime("%Y-%m-%d %H:%M")))

    # 总览卡片
    A("<div class='cards'>")
    A("<div class='card %s'><div class='k'>总体合规率</div><div class='v'>%.1f%%</div></div>" % (
        "bad" if overall < 60 else ("warn" if overall < 90 else "ok"), overall))
    A("<div class='card'><div class='k'>受检段落</div><div class='v'>%d</div></div>" % total_checked)
    A("<div class='card bad'><div class='k'>结构缺陷</div><div class='v'>%d</div></div>" % n_err)
    A("<div class='card warn'><div class='k'>格式偏离条目</div><div class='v'>%d</div></div>" % len(viols))
    A("<div class='card'><div class='k'>全文段落</div><div class='v'>%d</div></div>" % result["paragraphs_nonempty"])
    A("</div>")

    if n_err:
        A("<div class='note'><b>结论：</b>检出 %d 项结构缺陷。格式层面可以自动修正，"
          "但结构缺陷（标题编号缺失/重复、参考文献无序号）涉及语义判断，"
          "建议先确认层级归属再动手。</div>" % n_err)

    # ---------- 各区合规率 ----------
    A("<h2>一、各区合规率</h2>")
    A("<table><tr><th>区块</th><th>段数</th><th>合规率</th><th></th><th>规范要求</th></tr>")
    for cat, v in result["per_category"].items():
        cls, color = rate_class(v["rate"])
        raw = spec["elements"].get(cat, {}).get("raw", "—")
        A("<tr><td>%s</td><td>%d</td><td><span class='pill %s'>%s%%</span></td>"
          "<td><div class='bar'><i style='width:%s%%;background:%s'></i></div></td>"
          "<td style='color:#5a5f66'>%s</td></tr>" % (
              esc(cat), v["total"], cls, v["rate"], v["rate"], color, esc(raw)))
    A("</table>")
    A("<div class='legend'>合规率 = 该区块中未检出任何格式偏离的段落占比。"
      "「低置信度」段落已单独列出，不计入可信结论。</div>")

    # ---------- 结构缺陷 ----------
    A("<h2>二、结构缺陷（优先级最高）</h2>")
    if not defects:
        A("<p>未检出结构缺陷。</p>")
    for d in defects:
        cls = "" if d["severity"] == "error" else " warn"
        A("<div class='defect%s'>" % cls)
        A("<div class='t'>%s <span class='pill %s'>%s</span></div>" % (
            esc(d["kind"]), "p-red" if d["severity"] == "error" else "p-orange",
            "严重" if d["severity"] == "error" else "提示"))
        A("<div class='d'>%s</div>" % esc(d["detail"]))
        if d.get("evidence"):
            A("<div class='ev'>%s</div>" % esc(d["evidence"]))
        A("<div class='sug'>建议：%s</div>" % esc(d["suggestion"]))
        A("</div>")

    # ---------- 版式与行距 ----------
    A("<h2>三、版式与行距</h2>")
    if result["layout_issues"]:
        A("<table><tr><th>节</th><th>项目</th><th>实际</th><th>要求</th><th>偏差</th></tr>")
        for x in result["layout_issues"]:
            A("<tr><td>第%d节</td><td>%s</td><td>%s</td><td>%s</td><td>%+.1fmm</td></tr>" % (
                x["sec"], esc(x["item"]), esc(x["actual"]), esc(x["expected"]), x["delta"]))
        A("</table>")
    else:
        pg = spec["page"]
        A("<p>纸张与页边距<b>全部合规</b>：A4 %s×%smm，上下 %.1fmm，左右 %.1fmm。</p>" % (
            pg["width_mm"], pg["height_mm"], pg["margin_top_mm"], pg["margin_left_mm"]))
    if result["line_spacing_issues"]:
        A("<h3>行距偏离</h3><table><tr><th>行距值</th><th>段数</th><th>要求</th></tr>")
        for x in result["line_spacing_issues"]:
            A("<tr><td>%s</td><td>%d</td><td>%s</td></tr>" % (
                esc(x["line_value"]), x["count"], esc(x["expected"])))
        A("</table>")
    else:
        A("<p>行距<b>合规</b>：全部受检段落均为规范许可的 1.5 倍或 1.25 倍。</p>")

    # ---------- 违规明细 ----------
    A("<h2>四、格式偏离明细</h2>")
    kinds = Counter(v.kind for v in viols)
    A("<div class='tabs'>")
    for k, c in kinds.most_common():
        A("<span class='pill p-blue'>%s %d</span>" % (esc(k), c))
    A("</div>")

    A("<table><tr><th>#</th><th>位置</th><th>区块</th><th>问题</th><th>实际</th><th>要求</th></tr>")
    shown = 0
    for idx in sorted(by_para):
        p = paras.get(idx, {})
        for v in by_para[idx]:
            if shown >= 200:
                break
            A("<tr><td>%d</td><td style='max-width:230px'>%s</td><td>%s</td>"
              "<td><b>%s</b></td><td><code>%s</code></td><td><code>%s</code></td></tr>" % (
                  idx, esc(v.text[:34]), esc(v.cat), esc(v.kind),
                  esc(v.actual), esc(v.expected)))
            shown += 1
        if shown >= 200:
            break
    A("</table>")
    if len(viols) > shown:
        A("<div class='legend'>仅展示前 %d 条，共 %d 条。完整清单见 JSON 输出。</div>" % (shown, len(viols)))

    # ---------- 低置信度 ----------
    low = [p for p in result["paragraphs"]
           if not p["empty"] and p["cat"] in spec["elements"] and p["conf"] < 0.7]
    A("<h2>五、需人工复核的判定</h2>")
    if not low:
        A("<p>无低置信度判定，全部段落分类可信。</p>")
    else:
        A("<table><tr><th>#</th><th>判定</th><th>置信度</th><th>文本</th></tr>")
        for p in low:
            A("<tr><td>%d</td><td>%s</td><td>%.2f</td><td>%s</td></tr>" % (
                p["idx"], esc(p["cat"]), p["conf"], esc(p["text"][:50])))
        A("</table>")

    # ---------- 规范空白 ----------
    A("<h2>六、规范未覆盖、本工具不判定的事项</h2>")
    A("<ul>")
    for x in spec.get("out_of_scope", []):
        A("<li>%s</li>" % esc(x))
    A("</ul>")
    if spec.get("ambiguous_sizes"):
        A("<div class='note'><b>待确认：</b>")
        for k, v in spec["ambiguous_sizes"].items():
            A("字号「%s」在 Word 中存在 %s 两种映射。%s" % (
                esc(k), esc(v["candidates_pt"]), esc(v["reason"])))
        A("</div>")

    A("<div class='foot'>本报告由论文格式体检工具生成，逐条结论均附段落定位与实际值，可复核。<br>"
      "判定只覆盖规范中有明确数值/格式条款的部分；涉及分页渲染、图片位置、扉页与推荐表一致性的要求"
      "不在此列。</div>")
    A("</div></body></html>")

    doc_html = "".join(h)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(doc_html)
    return out_path


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    docx_path, spec_path = sys.argv[1], sys.argv[2]
    out = sys.argv[3] if len(sys.argv) > 3 else "体检报告.html"
    with open(spec_path, encoding="utf-8") as f:
        spec = json.load(f)
    result = audit(docx_path, spec)
    path = build(result, out)
    print("已生成报告: %s" % path)
    total = sum(v["total"] for v in result["per_category"].values())
    clean = sum(v["clean"] for v in result["per_category"].values())
    print("总体合规率: %.1f%%  结构缺陷: %d  格式偏离: %d" % (
        clean / total * 100 if total else 0,
        sum(1 for d in result["structural_defects"] if d["severity"] == "error"),
        len(result["violations"])))


if __name__ == "__main__":
    main()
