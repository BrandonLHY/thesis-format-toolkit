# macOS 真机验收清单

面向**手上有 Mac** 的人。这里测的是 CI **测不到**的东西。

CI（GitHub Actions 的 `macos-latest`）已经覆盖了：文档改写逻辑、字体解析、
文件读写。它测不到的是：真机上的字体探测行为、界面外观、
生成文档在 Mac 版 Word / Pages 里的实际表现、`.app` 打包。

按优先级排，**第 1 步最重要**。

---

## 准备

```bash
# 1. 确认有 Python 3.10+
python3 --version
# 若提示找不到命令，装 Xcode 命令行工具即可获得 python3：
#   xcode-select --install

# 2. 取代码并装依赖
git clone https://github.com/BrandonLHY/thesis-format-toolkit.git
cd thesis-format-toolkit
pip3 install -r requirements.txt
```

> 如果 `pip3 install` 报权限错，加 `--user`：
> `pip3 install --user -r requirements.txt`

---

## 第 1 步 ★ 字体探测诊断（最重要，CI 替代不了）

在 CI runner 上，`system_profiler` 返回的是**结构合法但字体列表为空**的 JSON，
程序因此退回「扫描字体目录」这条降级路径。

**真机是什么行为，我们不知道。** 这决定了 macOS 用户能不能拿到正确的字体替换。

```bash
python3 selftest_macos_live.py
```

它会打印：`system_profiler` 的返回码、输出长度、JSON 结构、
解析出多少个字体名、`installed_fonts()` 返回什么、核心字体命中情况。

**请把完整输出发回来。** 也可以先跑这一条快速看一眼：

```bash
system_profiler SPFontsDataType -json | head -c 500
```

| 看到什么 | 含义 |
|---|---|
| `SPFontsDataType` 里有 `_items`，且里面有很多字体条目 | ✅ 真机正常，字体替换会生效。降级路径只是 CI 的特殊情况 |
| `_items` 是空的 `[]` | ⚠️ 真机也退化，**需要改进方案**——请务必反馈 |
| 命令不存在 / 报错 | ⚠️ 环境特殊，也请反馈 |

---

## 第 2 步 跑全部自检

```bash
python3 run_all_selftests.py
```

预期：六套全通过。如果哪一套失败，把失败明细发回来。

```bash
# 单独跑，看得更清楚
python3 selftest_macos_sim.py      # 模拟两种 macOS 环境
python3 selftest_fontparse.py      # 字体解析（含录制回放）
```

---

## 第 3 步 界面外观（CI 完全测不到）

```bash
python3 main.py
```

**逐项检查：**

- [ ] 窗口能正常打开，标题栏显示「论文格式改写器 v1.2.1」
- [ ] **14 行内容类型是否都看得见**（表格可能需要滚动，看看滚动是否正常）
- [ ] 右侧表格的列是否够用——「中文字体 / 字号 / 加粗 / 对齐 / 首行缩进 / 行距」
      有没有被窗口边缘截掉
- [ ] 中文是否正常显示（没有方框、乱码）
- [ ] **字体下拉框里显示的是不是 Mac 字体名**（应该是 Songti SC / Heiti SC /
      Kaiti SC 这类，而不是「宋体 / 黑体」——这说明平台适配生效了）
- [ ] Retina 屏上文字是否清晰（tkinter 不跟随 HiDPI 时可能发虚）
- [ ] 点「选择文件…」能打开文件选择器吗
- [ ] 界面文字有没有被截断（Mac 上字体度量与 Windows 不同）

**如果布局有问题，请截图。** 这是最容易在 Mac 上出问题的地方。

---

## 第 4 步 端到端实测（用真实论文）

```bash
# 用你自己的一篇论文（脱敏的也行）
python3 main.py
```

1. 选择一份 `.docx` 论文
2. 核对左侧识别结果对不对（每类内容各有多少段）
3. 点「生成新文档」
4. **在 Word for Mac 或 Pages 里打开生成的文件**，检查：

- [ ] 标题字体是否真的是黑体风格（Heiti SC）
- [ ] 正文字体是否真的是宋体风格（Songti SC）
- [ ] 字号、行距、首行缩进是否符合预期
- [ ] 上标是否正常（引用 `[1]`、注释 `①②`）
- [ ] 标题编号是否按「一、／（一）／1、」重排
- [ ] **原文件是否完全没被改动**（对比修改时间）

**这一步是最终判据。** 前几步都是中间产物，只有这里能回答
「Mac 上的用户实际拿到的东西对不对」。

命令行方式（便于批量测试）：

```bash
python3 rewrite_cli.py 你的论文.docx 输出.docx
```

---

## 第 5 步 打包成 .app（可选，进阶）

这一步只能在 Mac 上做（PyInstaller 不能跨平台打包）。

```bash
pip3 install pyinstaller
python3 build_exe.py
```

> 注意：`build_exe.py` 目前是为 Windows 写的（产物名是 `.exe`、
> 使用了 Windows 特有的参数）。在 Mac 上**大概率需要改脚本**。
> 如果你走到这一步，把报错发回来，我针对 macOS 改一版。

---

## 回报模板

把下面这段填好发回来，我就能判断该改什么：

```text
【Mac 环境】
macOS 版本：
芯片：Intel / Apple Silicon
Python 版本（python3 --version）：

【第 1 步：字体探测】
system_profiler 返回码：
输出长度（字节）：
SPFontsDataType 里 _items 是空的还是有很多条目：
解析出多少个字体名：
installed_fonts() 返回多少个：
（完整输出贴在这里）

【第 2 步：自检】
run_all_selftests.py 结果：
失败的套件（如有）及明细：

【第 3 步：界面】
能打开吗：
14 行是否都可见：
字体下拉框里是 Mac 字体名还是「宋体/黑体」：
布局有没有问题（截图）：

【第 4 步：端到端】
生成的文档在 Word/Pages 里字体对吗：
上标、编号正常吗：
原文件是否未改动：
```

---

## 已知会遇到的坑（提前告知）

| 现象 | 原因 | 处理 |
|---|---|---|
| `python3` 找不到 | 没装 Xcode 命令行工具 | `xcode-select --install` |
| `import tkinter` 失败 | 系统自带 Python 可能不带 Tk | 从 [python.org](https://www.python.org/downloads/) 装官方版；或用 `brew install python-tk` |
| `pip3 install` 权限错 | 系统 Python 受保护 | 加 `--user`，或用虚拟环境 `python3 -m venv .venv` |
| 界面字体发虚 | tkinter 在 Retina 上的已知问题 | 记录现象即可，属显示层面 |
| 首次运行被拦截 | Gatekeeper | 从源码运行不会有这个问题；打包成 .app 后才有 |
| 生成的文档里字体没变 | 见第 1 步——可能真机也走降级路径 | 反馈诊断输出 |

---

## 关键概念：CI 通过 ≠ 真机没问题

这个项目的 macOS 支持**已经过了 CI**，但那只能证明：

- 文档改写逻辑在 macOS 上能跑
- 字体解析函数对**录制的**输入处理正确
- 代码没有平台相关的崩溃

它**证明不了**：

- 真机上 `system_profiler` 的实际行为（CI runner 是受限环境，可能不代表真机）
- 界面在 macOS 上长什么样、布局是否合理
- 生成的文件在 Mac 版 Word 里字体是否真的生效

**这正是需要真机验收的原因。**
