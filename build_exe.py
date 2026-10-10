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
打包脚本
========
产出两种分发形态：

  单文件版  dist/论文格式改写器.exe          一个 exe，最方便
  文件夹版  dist/论文格式改写器_文件夹版/     启动最快、最稳（不依赖临时目录）

为什么要有文件夹版：单文件 exe 每次启动都要把自己解压到系统临时目录，
若对方机器临时目录受限（策略、磁盘满、杀软拦截），会直接报
「Could not create temporary directory」打不开。文件夹版不经过这一步。

用法:
    python build_exe.py            # 两种都打
    python build_exe.py onedir     # 只打文件夹版
    python build_exe.py onefile    # 只打单文件版
"""
import os
import shutil
import subprocess
import sys


# 保证中文输出在任何终端都不崩。
# Windows 上 stdout 默认编码可能是 cp1252，直接 print 中文会抛
# UnicodeEncodeError 并让脚本崩掉（GitHub Actions 的 windows runner 实测踩过）。
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
HERE = os.path.dirname(os.path.abspath(__file__))
ENTRY = os.path.join(HERE, "main.py")
DIST = os.path.join(HERE, "dist")
BUILD = os.path.join(HERE, "build")


def docx_template_add_data() -> list[str]:
    """显式把 python-docx 的模板目录交给 PyInstaller。

    为什么不能只靠 PyInstaller 自动收集：实测 --collect-all docx 与
    --collect-data docx 都**没有**把 docx/templates 打进文件夹版，
    结果 _internal\\docx\\templates 完全不存在，Document() 一调用就报
    PackageNotFoundError。显式 --add-data 才可靠。
    """
    try:
        import docx
    except ImportError:
        return []
    tpl = os.path.join(os.path.dirname(os.path.abspath(docx.__file__)), "templates")
    if not os.path.isdir(tpl):
        return []
    print("python-docx 模板目录: %s" % tpl)
    sep = ";" if os.name == "nt" else ":"
    return ["--add-data", "%s%s%s" % (tpl, sep, os.path.join("docx", "templates"))]


HIDDEN = [
    "--hidden-import", "docx",
    "--hidden-import", "docx.oxml",
    "--hidden-import", "docx.oxml.ns",
    "--hidden-import", "docx.opc",
    "--hidden-import", "docx.text.paragraph",
    "--hidden-import", "docx.text.run",
    "--hidden-import", "lxml",
    "--hidden-import", "lxml.etree",
    "--hidden-import", "lxml._elementpath",
    "--collect-all", "docx",
] + docx_template_add_data()


def build(name: str, mode: str, console: bool = False) -> str | None:
    sep = ";" if os.name == "nt" else ":"
    args = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--%s" % mode,
        "--console" if console else "--windowed",
        "--name", name,
        "--distpath", DIST,
        "--workpath", BUILD,
        "--specpath", HERE,
    ] + HIDDEN
    if mode == "onefile":
        # 关键：单文件 exe 默认解压到系统临时目录，若临时目录不可写就会报
        # 「Could not create temporary directory」而无法启动（本机实测踩到）。
        # 改成解压到 exe 自己所在的目录，不依赖系统临时目录。
        args += ["--runtime-tmpdir", "_MEI_temp"]
    specs = os.path.join(HERE, "specs")
    if os.path.exists(specs):
        args += ["--add-data", "%s%s%s" % (specs, sep, "specs")]
    args.append(ENTRY)

    print("\n" + "=" * 70)
    print("打包 [%s]  模式=%s" % (name, mode))
    print("=" * 70)
    rc = subprocess.call(args, cwd=HERE)
    if rc != 0:
        print("!! 打包失败，退出码 %d" % rc)
        return None

    target = os.path.join(DIST, "%s.exe" % name) if mode == "onefile" \
        else os.path.join(DIST, name)
    if not os.path.exists(target):
        print("!! 未找到产物: %s" % target)
        return None
    if mode == "onefile":
        print("完成: %s  (%.1f MB)" % (target, os.path.getsize(target) / 1024 / 1024))
    else:
        total = sum(os.path.getsize(os.path.join(dp, f))
                    for dp, _, fs in os.walk(target) for f in fs)
        print("完成: %s  (整个文件夹 %.1f MB)" % (target, total / 1024 / 1024))
    return target


def write_readme(folder: str):
    """往文件夹版里放一份给最终用户的说明。"""
    txt = """论文格式改写器 —— 使用说明
================================

【怎么用】
1. 双击「论文格式改写器.exe」打开程序
2. 点左上「选择文件…」挑一份 .docx 论文（旧版 .doc 不行，需先另存为 .docx）
3. 左侧会显示识别结果，核对有没有认错的
4. 右侧可以调每一类内容的字体字号（改完会自动记住）
5. 点右下「生成新文档」，选保存位置

【重要】
· 程序不会修改你的原文件，只会生成一个新文件
· 标题序号默认会重排（把 WPS 自动编号改成文字编号）。不想改就把那个勾去掉
· 生成后打开新文档，目录里的页码需要手动更新一次（点一下目录按 F9）

【发给别人】
· 把整个文件夹一起发过去，不要只发 exe
· 对方电脑需要 64 位 Windows 7 或更高版本
· 首次打开如果被 Windows 拦截，点「更多信息」→「仍要运行」

【许可证】
本程序是自由软件，以 GNU 通用公共许可证第 3 版（或更新版本）发布。
完整许可证文本见同目录的 LICENSE 文件；源码见：
https://github.com/BrandonLHY/thesis-format-toolkit
这意味着你可以自由使用、修改、再分发，但分发衍生作品时须同样以 GPL 开源。

【出问题怎么办】
· 打不开：在文件夹里按住 Shift 右键 →「在此处打开 PowerShell 窗口」，
  输入  .\\论文格式改写器.exe --selftest  回车，然后把生成的「自检报告.txt」发回来
"""
    with open(os.path.join(folder, "使用说明.txt"), "w", encoding="utf-8") as f:
        f.write(txt)
    print("已写入使用说明.txt")


def copy_license(folder: str) -> bool:
    """把 LICENSE 复制进发布文件夹。

    GPL-3.0 义务：对外分发二进制时应随附许可证文本。
    虽然也可以在文档里给出网络获取地址，但把文本直接放进包里最稳妥。
    """
    src = os.path.join(HERE, "LICENSE")
    if not os.path.exists(src):
        print("!! 未找到 LICENSE，跳过（GPL 要求随附许可证文本，请手动补上）")
        return False
    shutil.copy2(src, os.path.join(folder, "LICENSE"))
    print("已复制 LICENSE（GPL-3.0 全文）")
    return True


def main():
    modes = sys.argv[1:] or ["onedir", "onefile"]
    made = {}

    if "onedir" in modes:
        p = build("论文格式改写器", "onedir")
        if p:
            write_readme(p)
            copy_license(p)
            made["文件夹版"] = p
    if "onefile" in modes:
        p = build("论文格式改写器_单文件版", "onefile")
        if p:
            print("提示：单文件版无法随附 LICENSE 文件，")
            print("      请在 Release 说明里注明许可证与源码地址（GPL 义务）")
            made["单文件版"] = p

    print("\n" + "=" * 70)
    print("打包结果")
    print("=" * 70)
    for k, v in made.items():
        print("  %-8s %s" % (k, v))
    if not made:
        return 1
    print("\n分发建议：优先发「文件夹版」，最稳；单文件版适合随手发。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
