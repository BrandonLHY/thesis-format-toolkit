# 论文格式改写器

[![自检](https://github.com/BrandonLHY/thesis-format-toolkit/actions/workflows/selftest.yml/badge.svg)](https://github.com/BrandonLHY/thesis-format-toolkit/actions/workflows/selftest.yml)
[![许可证: GPL-3.0](https://img.shields.io/badge/%E8%AE%B8%E5%8F%AF%E8%AF%81-GPL--3.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/downloads/)

把 Word 论文按格式规范**重新排版**：自动识别标题层级，套用字体字号，输出新文档。

**不修改原文件** —— 只读取，结果写到新文件。学生最怕把论文改坏，所以这条是底线。

- **本地运行** —— 不联网、不上传任何内容
- **格式可自定义** —— 字体字号在界面上改，不用动代码
- **规范可替换** —— 换学校只需新增一个 JSON，程序不用改
- **六套自检** —— 共 175 项，每次提交自动运行

---

## 解决什么问题

毕业论文有严格的格式要求：一级标题四号黑体、正文小四号宋体、行距 1.5 倍、
参考文献小五号……几十条。手工调一遍要几小时，而且**字体字号必须逐个段落设**，
漏掉一处就要重来。

更麻烦的是标题层级。Word 里看起来一样的「一、引言」和「（一）基本情况」，
在文档内部可能是完全不同的东西——有的挂了命名样式，有的是直接设置格式，
有的是 WPS 自动编号，还有的**只是纯文字**。

这个工具做两件事：

1. **识别**：判断每个段落属于哪类内容（大标题 / 一至三级标题 / 正文 / 摘要 /
   关键词 / 图表题注 / 参考文献…）
2. **改写**：按你设定的格式，把每类内容统一改过来

---

## 快速开始

### 方式一：用打包好的程序（Windows）

1. 下载 Release 里的压缩包，解压到**有写权限的目录**
2. 双击 `论文格式改写器.exe`
3. 选论文 → 核对识别结果 → 调整格式 → 生成新文档

> 放在 `C:\Program Files` 下会因权限问题无法保存设定。
> 首次运行可能被 SmartScreen 提示「未知发布者」，点「更多信息」→「仍要运行」。

### 方式二：从源码运行（Windows / macOS / Linux）

```bash
pip install -r requirements.txt
python main.py
```

macOS 上可以双击 `启动.command`，它会自动检查并安装依赖。
详细说明见 [docs/Mac使用说明.md](docs/Mac使用说明.md)。

### 方式三：命令行

```bash
# 改写（默认会重排标题编号、处理上标）
python rewrite_cli.py 论文.docx 输出.docx

# 保留原有标题编号
python rewrite_cli.py 论文.docx 输出.docx --keep-numbering

# 不做上标处理
python rewrite_cli.py 论文.docx 输出.docx --no-superscript --no-note-marks
```

---

## 它会改什么

| 项目 | 说明 |
|---|---|
| 字体字号 | 中文/西文字体、字号、加粗，按内容类型分别设置 |
| 段落格式 | 对齐、首行缩进（字符单位）、行距、段前段后 |
| 标题编号 | 把 WPS/Word 的自动编号转成文字编号，按章重置（可选） |
| 上标 | 正文引用 `[1]` 与注释序号 `①②` 设为上标；参考文献条目里的 `[n]` 不动 |
| 跨平台字体 | 在 macOS 上自动把「宋体」替换成系统存在的 Songti SC 等 |

**不改动文字内容。** 唯一的例外是你主动开启「重排标题序号」时，
它会在标题前补写编号（如「二、相关概念和理论」）。

---

## 适用边界（请先读）

这个工具不是万能的，写清楚免得你白费时间。

**表现好的情况：**

- 论文有**编号**（一、／（一）／1、）或作者用过 Word 的标题样式
- 结构相对清晰，章节层次分明

**表现会退化的情况：**

- **完全没有编号的纯文本稿**：如果一段文字既没有编号、也没有加粗、
  字号也和其他段落一样，程序只能靠语义猜测，会出错
- 层级逻辑本身混乱的稿子（例如「研究背景」既可能是一级标题也可能是二级标题）

**实测数据：** 以 Word 自身的大纲级别（`w:outlineLvl`）为基准，
在一份 138 段的真实论文上，区域识别**精确率 86%、召回率 97%**。

**因此界面上会显示识别结果让你核对**，而不是直接闷头改完给你。
识别有疑问的段落会标出来。

**不覆盖的内容：**

- 分页渲染类要求（图题不拆页、表格跨页续表标注）
- 公式的居中与编号位置
- 目录页码自动刷新（需在 Word 里按 F9）
- 旧版 `.doc` 格式（需先另存为 `.docx`）

---

## 自定义格式

界面底部可以调整每个内容类型的字体、字号、加粗、对齐、缩进、行距。
改完自动保存到程序目录下的 `预设.json`，下次打开还在。

换学校/换规范，只需新增一个 JSON 文件到 `specs/`，结构参见
[specs/sysu_undergraduate_thesis.json](specs/sysu_undergraduate_thesis.json)。

```jsonc
{
  "elements": {
    "正文": {
      "font_ea": "宋体",           // 中文字体
      "font_ascii": "Times New Roman",
      "pt": 12.0,                  // 字号（磅），小四 = 12
      "bold": false,
      "raw": "小四号宋体",          // 规范原文，报告里展示
      "extra_rules": ["首行缩进2字符"]
    }
  },
  "page": { "width_mm": 210.0, "margin_top_mm": 25.4, "tolerance_mm": 1.0 },
  "line_spacing": { "allowed": [{ "label": "1.5倍", "line": 360 }] }
}
```

---

## 附带：格式合规体检工具

仓库里还有一个**只读**的检查工具，按规范逐段打分，输出带定位的 HTML 报告。

```bash
python audit_engine.py 论文.docx specs/sysu_undergraduate_thesis.json
python report_html.py 论文.docx specs/sysu_undergraduate_thesis.json 报告.html
```

它检出的是「哪里不符合规范」，不做修改。适合投稿前自查。

---

## 自检

**六套自检，共 175 项**，每次 push 由 CI 自动运行（Windows + Ubuntu × Python 3.10/3.12）。

```bash
python run_all_selftests.py     # 一键跑全部并汇总
```

单独运行：

```bash
python 验收检验.py                # 37 项：模拟全新用户走完整流程
python selftest_adversarial.py   # 39 项：畸形输入的健壮性
python selftest_fidelity.py      # 34 项：输出保真度与元数据保全
python selftest_synthetic.py     # 36 项：假想论文端到端
python selftest_platform.py      # 18 项：跨平台字体适配
python main.py --selftest        # 11 项：依赖、GUI、端到端
```

### 各套覆盖什么

| 套件 | 重点 |
|---|---|
| **验收检验** | 依赖 → 造论文 → 识别 → 改写 → 核对字体/编号/上标/文字完整性/原文件安全/GUI |
| **对抗性 A** | 空文档、仅空白、段落无 run、超长标题、重复编号、表格内段落、含图片、项目符号列表、异常字号、坏样式引用、嵌套内容控件、**跨 run 上标**、全文单段、中英混排 |
| **对抗性 B** | 文档属性、页眉页脚、书签、域代码、内容控件、分节属性、纸张尺寸、媒体文件、空段落保留、幂等性、预设生效性 |
| **假想论文** | 纯文本段与有格式段、编号标题与无编号结构词、参考文献条目序号、全角括号、多引用区间、注释混排、年份区间、摘要标注同段 |
| **跨平台字体** | 字体名不含逗号、Windows 不替换、模拟 macOS 替换、探测失败降级、界面选项 |
| **程序自检** | 打包后依赖完整、GUI 可构建、上标规则、端到端改写 |

### 自检实际发现过的缺陷

这些不是设想出来的用例，每一条都对应一次真实修复：

| 缺陷 | 怎么被发现的 |
|---|---|
| `（一）` 二级标题被判成一级 | 验收检验（「黑体14pt + （一）前缀」这个组合） |
| **跨 run 的引用标记漏上标** | 对抗性检验 A |
| 主题字体压制具体字体名 | 真实文档 + Word COM 读取实际字体 |
| 字体回退链在 Word 中无效 | 专门写实验验证假设，结果推翻了原方案 |
| `outlineLvl` 无 `val` 时整类标题漏检 | 分类器自检（「长度>50 的标题应为 0 个」） |
| run 对象身份比较导致标题文字被清空 | 逐层插桩追踪 |
| 分区无退出条件，吞掉 88 段正文 | 阅读输出时发现数字异常 |
| 目录区在内容控件内，从未被改写 | 对比 Word 报告的段落数与程序看到的段落数 |
| 打包缺 `docx/templates`，新建文档报错 | 打包后跑自检 |
| 单文件 exe 报临时目录错误，无法启动 | 实际运行 exe |

---

## 开发过程中踩过的坑

这些是实测发现、且会导致错误结论的问题，都已修复并有自检覆盖：

| 坑 | 现象 | 处理 |
|---|---|---|
| **主题字体压制** | 写了具体字体名却无效，Word 报告为「+中文正文」 | `w:eastAsiaTheme` 与 `w:eastAsia` 同时存在时主题优先，必须先删主题引用 |
| **字体回退链无效** | 写入 `"宋体,Songti SC"` 以为会依次尝试 | Word 把整串当成**一个字体名**。必须运行时探测系统字体，只写存在的那个 |
| **`outlineLvl` 默认值** | 漏检全部某类标题 | `<w:outlineLvl/>` 无 `w:val` 时 OOXML 默认是 **0**（一级标题），当成 None 就漏了 |
| **run 对象身份比较** | 刚写入的标题文字立刻被清空 | python-docx 每次访问 `p.runs` 都新建包装对象，`r is not first` 永远为真，必须比较 XML 元素 |
| **分区没有退出条件** | 88 段正文被误判为「摘要内容」 | 前置区必须能被大标题打断 |
| **封面识别过宽** | 106 段的文档只剩 3 段参与判定 | 限定范围 + 保护名单 |
| **目录区看不到** | 目录 28 段从未被改写 | 目录在 `w:sdt` 内容控件内，`doc.paragraphs` 不返回 |
| **超链接里的 run 看不到** | 目录条目全部漏改 | `p.runs` 不含 `w:hyperlink` 内的 run |
| **参考文献序号误上标** | 条目 `[1]` 被当成正文引用 | 按区域区分：正文区上标，条目区不动 |
| **年份区间误上标** | `[2010-2015]` 被当引用 | 用数字位数区分（引用序号不超过 3 位） |
| **跨 run 的引用标记漏上标** | `"…的["` + `"4"` + `"]标记"` 三个 run，标记被拆开，逐 run 匹配全都看不到 | 改为按段落拼接文本定位标记，再映射回各 run |
| **判据优先级错误** | `（一）` 二级标题因字体是黑体14pt 被判成一级标题 | 编号模式必须优先于字号判据——编号是作者显式标注的层级 |
| **打包缺模板文件** | 打包后新建文档报 `PackageNotFoundError` | PyInstaller 的 `--collect-all` 没带上 `docx/templates`，必须显式 `--add-data` |
| **单文件 exe 打不开** | 报 `Could not create temporary directory` | 用 `--runtime-tmpdir` 让它解压到自身目录 |

---

## 项目结构

```
.
├── main.py                  程序入口（--selftest 自检）
├── app.py                   图形界面
├── rewrite_cli.py           命令行入口
├── rewriter.py              改写引擎 + 默认预设
├── region_detector.py       区域识别
├── numbering.py             标题编号处理
├── superscript.py           上标处理
├── platform_fonts.py        跨平台字体适配
├── audit_engine.py          合规体检引擎
├── report_html.py           体检报告渲染
├── build_exe.py             Windows 打包
├── selftest_*.py            自检
├── specs/                   格式规范（可扩展）
├── docs/                    使用说明
└── .github/workflows/       CI
```

---

## 开发说明

本项目在开发过程中使用了 **AI 编程助手（DeepSeek Harness）** 辅助实现。

需求定义、产品形态设计、真实文档测试、问题定位与最终验收由作者完成。
代码由**六套自检共 175 项**覆盖，并在多份真实论文上验证——
上表列出的每一项已知缺陷，都来自真实使用中暴露的问题。

---

## 许可证

**GNU General Public License v3.0 or later** —— 见 [LICENSE](LICENSE)。

这意味着：你可以自由使用、修改、分发本程序，但**分发衍生作品时必须同样以 GPL 开源**。

依赖库许可证：`python-docx`（MIT）、`lxml`（BSD），均与 GPL 兼容。

---

## 关于格式规范的来源

`specs/sysu_undergraduate_thesis.json` 中的字号、页边距等**数值**来自
《中山大学本科生毕业论文格式要求》。规则的具体数值属于客观事实，
此处以机器可读的形式重新表达，便于程序使用。

其他学校/刊物只需按同样结构新增 JSON，**程序不需要修改**。
欢迎提交你所在院校的规范文件。

---

## 贡献

欢迎 issue 和 PR。提交前请先跑一遍全部自检：

```bash
python run_all_selftests.py
```

如果你有**识别错误的真实案例**，那是特别有价值的反馈——
请附上（脱敏后的）文档结构描述和你期望的识别结果。
