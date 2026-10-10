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
论文格式改写器 —— 桌面版
========================
一个窗口搞定：选文件 → 选格式 → 生成新文档。

设计原则
--------
1. 永不修改原文件，只生成新文件。
2. 识别结果先给用户看，再动手（识别有出错的可能）。
3. 预设全部可视化调节，不需要懂命令行。
"""
from __future__ import annotations

import os
import queue
import re
import sys
import threading
import traceback
import tkinter as tk
from dataclasses import asdict
from tkinter import filedialog, messagebox, ttk

# ---- 让打包后的 exe 也能 import 同目录的模块 ----
APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)


def user_dir() -> str:
    """用户可写目录（exe 或脚本所在目录）。

    单文件打包后 __file__ 指向临时解压目录，程序退出就被删除，
    因此用户改的格式设定必须存到 exe 自己所在的目录，否则每次运行都丢失。
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return APP_DIR

from region_detector import REGION_TYPES, detect_regions          # noqa: E402
from rewriter import (                                            # noqa: E402
    ALIGN_CHOICES, CN_SIZE, DEFAULT_PRESET, FONT_CHOICES, SIZE_CHOICES,
    RegionFormat, preset_from_dict, preset_to_dict, rewrite,
)

APP_TITLE = "论文格式改写器"
VERSION = "1.1"
ASCII_FONTS = ["Times New Roman", "Arial", "Calibri", "Courier New"]
PRESET_FILENAME = "预设.json"

# 字体选项按平台给出：macOS 上没有「宋体」这些名字，
# 界面直接显示 Mac 对应字体（Songti SC 等），避免用户选了不存在的字体。
try:
    from platform_fonts import open_path as _open_path, platform_font_choices
    EA_FONT_CHOICES = platform_font_choices(FONT_CHOICES)
except Exception:
    _open_path = None
    EA_FONT_CHOICES = list(FONT_CHOICES)

# 界面上每行可调的列：(字段名, 显示名, 宽度)
COLUMNS = [
    ("font_ea", "中文字体", 120),
    ("size", "字号", 78),
    ("bold", "加粗", 52),
    ("align", "对齐", 92),
    ("first_line_indent_chars", "首行缩进", 72),
    ("line_spacing", "行距", 62),
    ("space_before_pt", "段前", 54),
    ("space_after_pt", "段后", 54),
]

SIZE_ORDER = SIZE_CHOICES


def preset_path() -> str:
    return os.path.join(user_dir(), PRESET_FILENAME)


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.doc_path: str | None = None
        self.paras: list[dict] = []
        self.summary: dict = {}
        self.preset: dict[str, RegionFormat] = self._load_preset()
        self.queue: queue.Queue = queue.Queue()
        self._build_ui()
        self.root.after(120, self._drain_queue)

    # ========================================================
    # 预设读写
    # ========================================================
    def _load_preset(self) -> dict[str, RegionFormat]:
        p = preset_path()
        if os.path.exists(p):
            try:
                import json
                with open(p, encoding="utf-8") as f:
                    loaded = preset_from_dict(json.load(f))
                # 补齐缺失的区域
                for k, v in DEFAULT_PRESET.items():
                    loaded.setdefault(k, v)
                return loaded
            except Exception:
                pass
        return {k: RegionFormat(**asdict(v)) for k, v in DEFAULT_PRESET.items()}

    def _save_preset(self, silent: bool = False):
        import json
        try:
            with open(preset_path(), "w", encoding="utf-8") as f:
                json.dump(preset_to_dict(self.preset), f, ensure_ascii=False, indent=2)
            if not silent:
                self.set_status("预设已保存到 %s" % PRESET_FILENAME)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "保存预设失败：%r" % e)

    # ========================================================
    # 界面
    # ========================================================
    def _build_ui(self):
        r = self.root
        r.title("%s v%s" % (APP_TITLE, VERSION))
        r.geometry("1080x760")
        r.minsize(940, 660)
        try:
            r.iconbitmap(default="")
        except Exception:
            pass

        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure("Big.TButton", font=("Microsoft YaHei", 11, "bold"), padding=(18, 10))
        style.configure("H.TLabel", font=("Microsoft YaHei", 10, "bold"))
        style.configure("Hint.TLabel", foreground="#6b7280")

        # ---------- 顶部：选文件 ----------
        top = ttk.LabelFrame(r, text=" 第一步：选择要改写的论文 ", padding=12)
        top.pack(fill="x", padx=14, pady=(12, 8))

        self.var_path = tk.StringVar(value="尚未选择文件")
        ttk.Label(top, textvariable=self.var_path, style="Hint.TLabel",
                  wraplength=780, justify="left").pack(side="left", fill="x", expand=True)
        ttk.Button(top, text="选择文件…", command=self.pick_file).pack(side="right")
        ttk.Button(top, text="先看看识别结果", command=self.analyze).pack(side="right", padx=(0, 8))

        # ---------- 中部：左右分栏 ----------
        mid = ttk.Frame(r)
        mid.pack(fill="both", expand=True, padx=14, pady=4)

        left = ttk.LabelFrame(mid, text=" 识别结果（识别有出错的可能，请过一眼） ", padding=10)
        left.pack(side="left", fill="both", expand=False)
        self.txt_result = tk.Text(left, width=40, height=22, wrap="word",
                                  font=("Microsoft YaHei", 9), relief="flat",
                                  background="#fbfbfd")
        sb1 = ttk.Scrollbar(left, orient="vertical", command=self.txt_result.yview)
        self.txt_result.configure(yscrollcommand=sb1.set)
        self.txt_result.pack(side="left", fill="both", expand=True)
        sb1.pack(side="right", fill="y")

        right = ttk.LabelFrame(mid, text=" 第二步：每类内容用什么格式（可直接改） ", padding=10)
        right.pack(side="left", fill="both", expand=True, padx=(10, 0))

        # 表头
        head = ttk.Frame(right)
        head.pack(fill="x")
        ttk.Label(head, text="内容类型", width=16, style="H.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 6))
        for ci, (_, label, w) in enumerate(COLUMNS, start=1):
            ttk.Label(head, text=label, width=max(5, w // 9), style="H.TLabel").grid(
                row=0, column=ci, sticky="w", padx=2)

        # 可滚动区域
        canvas = tk.Canvas(right, highlightthickness=0)
        vsb = ttk.Scrollbar(right, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))

        self.vars: dict[str, dict] = {}
        for ri, (key, label, _desc) in enumerate(REGION_TYPES, start=1):
            fmt = self.preset[key]
            ttk.Label(inner, text=label, width=16).grid(row=ri, column=0, sticky="w", padx=(0, 6), pady=3)
            v = {}
            for ci, (fname, _lbl, _w) in enumerate(COLUMNS, start=1):
                if fname == "bold":
                    var = tk.BooleanVar(value=bool(getattr(fmt, fname)))
                    ttk.Checkbutton(inner, variable=var,
                                    command=lambda k=key: self._on_change(k)).grid(
                        row=ri, column=ci, sticky="w", padx=2)
                elif fname == "font_ea":
                    var = tk.StringVar(value=getattr(fmt, fname))
                    cb = ttk.Combobox(inner, textvariable=var, values=EA_FONT_CHOICES,
                                      width=13, state="readonly")
                    cb.grid(row=ri, column=ci, sticky="w", padx=2)
                    cb.bind("<<ComboboxSelected>>", lambda e, k=key: self._on_change(k))
                elif fname == "size":
                    var = tk.StringVar(value=getattr(fmt, fname))
                    cb = ttk.Combobox(inner, textvariable=var, values=SIZE_CHOICES,
                                      width=6, state="readonly")
                    cb.grid(row=ri, column=ci, sticky="w", padx=2)
                    cb.bind("<<ComboboxSelected>>", lambda e, k=key: self._on_change(k))
                elif fname == "align":
                    var = tk.StringVar(value=getattr(fmt, fname))
                    cb = ttk.Combobox(inner, textvariable=var, values=ALIGN_CHOICES,
                                      width=8, state="readonly")
                    cb.grid(row=ri, column=ci, sticky="w", padx=2)
                    cb.bind("<<ComboboxSelected>>", lambda e, k=key: self._on_change(k))
                else:
                    var = tk.StringVar(value=str(getattr(fmt, fname)))
                    ent = ttk.Entry(inner, textvariable=var, width=6)
                    ent.grid(row=ri, column=ci, sticky="w", padx=2)
                    ent.bind("<FocusOut>", lambda e, k=key: self._on_change(k))
                v[fname] = var
            self.vars[key] = v

        # ---------- 底部：动作 ----------
        bottom = ttk.Frame(r)
        bottom.pack(fill="x", padx=14, pady=(8, 4))

        self.var_renumber = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom, text="重排标题序号",
                        variable=self.var_renumber).pack(side="left")
        self.var_sup_cite = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom, text="引用 [1] 上标",
                        variable=self.var_sup_cite).pack(side="left", padx=(12, 0))
        self.var_sup_note = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom, text="注释 ① 上标",
                        variable=self.var_sup_note).pack(side="left", padx=(12, 0))

        ttk.Button(bottom, text="恢复默认格式", command=self.reset_preset).pack(side="right")
        ttk.Button(bottom, text="保存我的格式", command=self._save_preset).pack(side="right", padx=8)
        self.btn_run = ttk.Button(bottom, text="生成新文档", style="Big.TButton", command=self.run)
        self.btn_run.pack(side="right", padx=(0, 8))

        # ---------- 状态栏 ----------
        self.var_status = tk.StringVar(value="请先选择一份 .docx 论文")
        bar = ttk.Frame(r)
        bar.pack(fill="x", padx=14, pady=(0, 10))
        ttk.Label(bar, textvariable=self.var_status, style="Hint.TLabel").pack(side="left")
        self.pb = ttk.Progressbar(bar, mode="indeterminate", length=160)

    # ========================================================
    # 交互
    # ========================================================
    def set_status(self, text: str):
        self.var_status.set(text)
        self.root.update_idletasks()

    def _on_change(self, key: str):
        """把界面上的值写回 preset。"""
        fmt = self.preset[key]
        for fname, var in self.vars[key].items():
            val = var.get()
            if fname == "bold":
                fmt.bold = bool(val)
            elif fname in ("first_line_indent_chars", "space_before_pt", "space_after_pt"):
                try:
                    setattr(fmt, fname, int(float(val)))
                except (TypeError, ValueError):
                    setattr(fmt, fname, 0)
            elif fname == "line_spacing":
                try:
                    setattr(fmt, fname, float(val))
                except (TypeError, ValueError):
                    setattr(fmt, fname, 1.5)
            else:
                setattr(fmt, fname, val)
        self._save_preset(silent=True)

    def reset_preset(self):
        if not messagebox.askyesno(APP_TITLE, "恢复全部格式设定为默认值？"):
            return
        self.preset = {k: RegionFormat(**asdict(v)) for k, v in DEFAULT_PRESET.items()}
        for key, vs in self.vars.items():
            f = self.preset[key]
            for fname, var in vs.items():
                var.set(getattr(f, fname))
        self._save_preset()
        self.set_status("已恢复默认格式")

    def pick_file(self):
        p = filedialog.askopenfilename(
            title="选择要改写的论文",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if not p:
            return
        if p.lower().endswith(".doc"):
            messagebox.showwarning(
                APP_TITLE,
                "这是旧版 .doc 格式，程序读不了。\n\n"
                "请先用 Word 或 WPS 打开它，另存为 .docx 后再试。")
            return
        self.doc_path = p
        self.var_path.set(p)
        self.set_status("已选择文件，点「先看看识别结果」可以预览")
        self.analyze()

    def analyze(self):
        if not self.doc_path:
            messagebox.showinfo(APP_TITLE, "请先选择一份 .docx 文件。")
            return
        self.set_status("正在识别文档结构…")
        self.pb.pack(side="right")
        self.pb.start(12)
        threading.Thread(target=self._analyze_worker, daemon=True).start()

    def _analyze_worker(self):
        try:
            paras, summary = detect_regions(self.doc_path)
            self.queue.put(("analyzed", paras, summary))
        except Exception:
            self.queue.put(("error", traceback.format_exc(), None))

    def run(self):
        if not self.doc_path:
            messagebox.showinfo(APP_TITLE, "请先选择一份 .docx 文件。")
            return
        base = os.path.splitext(os.path.basename(self.doc_path))[0]
        out = os.path.join(user_dir(), "%s_已排版.docx" % base)
        out = filedialog.asksaveasfilename(
            title="新文档保存到哪", initialdir=user_dir(),
            initialfile="%s_已排版.docx" % base,
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")])
        if not out:
            return
        self.set_status("正在改写…")
        self.btn_run.state(["disabled"])
        self.pb.pack(side="right")
        self.pb.start(12)
        threading.Thread(target=self._run_worker, args=(out,), daemon=True).start()

    def _run_worker(self, out: str):
        try:
            for k in self.preset:
                self.queue.put(("sync", k, None))
            result = rewrite(self.doc_path, out, preset=self.preset,
                             renumber_headings=self.var_renumber.get(),
                             superscript_citations=self.var_sup_cite.get(),
                             superscript_notes=self.var_sup_note.get())
            self.queue.put(("done", result, out))
        except Exception:
            self.queue.put(("error", traceback.format_exc(), None))

    def _drain_queue(self):
        try:
            while True:
                kind, a, b = self.queue.get_nowait()
                if kind == "analyzed":
                    self.paras, self.summary = a, b
                    self._render_result()
                    self.pb.stop()
                    self.pb.pack_forget()
                    self.set_status("识别完成，请核对左侧结果后点「生成新文档」")
                elif kind == "done":
                    self.pb.stop()
                    self.pb.pack_forget()
                    self.btn_run.state(["!disabled"])
                    self.set_status("完成：%s" % os.path.basename(b))
                    self._render_run_result(a)
                    if messagebox.askyesno(APP_TITLE, "改写完成！\n\n%s\n\n现在打开吗？" % b):
                        opened = False
                        if _open_path is not None:
                            opened = _open_path(b)
                        if not opened:
                            try:
                                os.startfile(b)      # Windows 兜底
                            except Exception:
                                messagebox.showinfo(APP_TITLE, "请手动打开：\n%s" % b)
                elif kind == "error":
                    self.pb.stop()
                    self.pb.pack_forget()
                    self.btn_run.state(["!disabled"])
                    self.set_status("出错")
                    messagebox.showerror(APP_TITLE, "出错了：\n\n%s" % a[-1200:])
        except queue.Empty:
            pass
        self.root.after(120, self._drain_queue)

    # ========================================================
    # 渲染
    # ========================================================
    def _render_result(self):
        t = self.txt_result
        t.configure(state="normal")
        t.delete("1.0", "end")
        counts = self.summary.get("counts", {})
        total = self.summary.get("total_nonempty", 0)
        t.insert("end", "共 %d 个非空段落\n\n" % total)
        for key, label, _desc in REGION_TYPES:
            n = counts.get(key, 0)
            if n:
                t.insert("end", "  %-12s %3d 段\n" % (label, n))
        low = self.summary.get("low_confidence", [])
        t.insert("end", "\n" + "-" * 34 + "\n")
        if low:
            t.insert("end", "以下 %d 段不太确定，请重点核对：\n\n" % len(low))
            for p in low[:12]:
                t.insert("end", "  #%d %s\n     %s\n" % (
                    p["idx"], p["region"], p["text"][:34]))
        else:
            t.insert("end", "识别置信度均正常。\n")
        t.configure(state="disabled")

    def _render_run_result(self, result: dict):
        t = self.txt_result
        t.configure(state="normal")
        t.insert("end", "\n" + "=" * 34 + "\n完成\n")
        t.insert("end", "  改写段落: %d\n" % result.get("changed_count", 0))
        if result.get("nested_count"):
            t.insert("end", "  目录域内: %d 段\n" % result["nested_count"])
        num = result.get("numbering")
        if num:
            t.insert("end", "  去掉自动编号: %d 处\n" % num["stripped_auto_numbering"])
            t.insert("end", "  重排文字编号: %d 处\n" % num["renumbered"])
        t.see("end")
        t.configure(state="disabled")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
