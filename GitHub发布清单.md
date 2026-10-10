# GitHub 发布检查清单

许可证：**GPL-3.0**

> 用法：从上往下逐项做完再 push。带 ⚠️ 的是**不做完不能推**的红线项。

---

## 一、红线：个人信息与隐私 ⚠️

这些文件一旦推上去就永久留在 Git 历史里，**必须在第一次 commit 之前排除**。

- [ ] ⚠️ **删除所有论文样本 .docx**（这些是别人的论文原文，含真实姓名）
  - `改写结果_案例报告.docx`（19.8 MB）
  - `需要改写案例报告_已排版.docx`
  - `样品1_已排版_修复版.docx`
  - `新建DOCX文档2_已排版.docx`
  - `_上标测试.docx`、`_回归测试.docx` 等带下划线的临时产物
- [ ] ⚠️ **删除体检报告样张 HTML**（含真实论文的文件名与统计数据）
  - `体检报告_案例报告.html`
  - `打开报告.bat`（指向上面那个文件，一并删或改）
- [ ] ⚠️ **清掉代码里的硬编码个人路径**
  - `一键改写.bat` 第 11 行含微信文件路径（暴露 wxid 账号）
  - `体检.bat` 第 13 行同上
  - 改法：默认值删掉，改成"把 docx 拖到本文件上"，或要求传参数
- [ ] ⚠️ **确认没有其他个人信息**
  - 扫描命令（在项目目录执行）：
    ```powershell
    Select-String -Path *.py,*.md,*.bat,*.json,*.html -Pattern "Administrator|C:\\Users|Desktop|微信|xwechat|wxid"
    ```
  - 预期结果：**零命中**
- [ ] **《预设编辑器.html》里清理示例文案**
  - 现在有 `cover_field:"课程名称：中国政府与政治"` 这类真实课程名，改成通用示例

---

## 二、必须新增的仓库文件 ⚠️

- [ ] ⚠️ **`.gitignore`**（没有它，`git add .` 会把你刚删的东西又加回来）
  必含内容见下方"附：.gitignore 模板"
- [ ] ⚠️ **`LICENSE`**（GPL-3.0 全文）
  - 获取方式：<https://www.gnu.org/licenses/gpl-3.0.txt>
  - 或让 GitHub 在新建仓库时自动生成
- [ ] **每个源码文件加版权头**（GPL 的常规做法）
  ```
  # 论文格式改写器 —— Word 论文格式批量规范化工具
  # Copyright (C) 2026 <你的名字>
  #
  # 本程序是自由软件：你可以按自由软件基金会发布的 GNU 通用公共许可证
  # （第 3 版或更新版本）条款重新发布和/或修改它。
  # ...
  ```
- [ ] **`README.md`** 至少包含：它是什么、解决什么问题、怎么用、系统要求、
  已知限制、许可证
- [ ] **AI 辅助开发的说明**（不强制，但建议）
  - 在 README 里用一小段说明使用了 AI 编程助手，同时写清你负责的部分
  - 参考写法（已加入 README「开发说明」一节）：
    > 本项目在开发过程中使用了 AI 编程助手（DeepSeek Harness）辅助实现。
    > 需求定义、产品形态设计、真实文档测试、问题定位与最终验收由作者完成。
  - 三个理由：诚实；成本为零；**它其实提升可信度**——把测试与验证摆出来，
    比声称"纯手写"更有说服力
  - 不要：在 README 顶部大字写「本工具由 AI 生成」（会埋没你自己的贡献）；
    也不要在 LICENSE 里写 AI 相关说明（许可证文件只放法律条款）
- [ ] **`requirements.txt`**（方便别人装依赖）
  ```
  python-docx>=1.1.0
  lxml>=5.0.0
  ```

---

## 三、GPL-3.0 的额外义务（选了 GPL 就必须做）⚠️

GPL 是"传染性"许可证，对外**分发二进制**（你的 exe）时有额外要求：

- [ ] ⚠️ **发 exe 时必须同时提供完整源码**
  - 最简单做法：README 里写明"二进制版本的源码即本仓库对应 tag"
  - 每次发 Release 时，在 Release 说明里附上对应 commit 的链接
- [ ] ⚠️ **随 exe 附上许可证文本**
  - 把 `LICENSE`（GPL-3.0 全文）一起放进发给用户的文件夹
  - 可以再加一个 `许可证说明.txt` 说明这是 GPL 程序
- [ ] **在程序里写明许可证**
  - 界面加一个"关于"按钮，或在窗口标题/状态栏注明 GPL-3.0
- [ ] **检查依赖库的许可证兼容性**（GPL 要求兼容）
  - `python-docx` → MIT ✅ 兼容
  - `lxml` → BSD ✅ 兼容
  - 结论：无冲突，但建议在 README 里列出依赖及其许可证

---

## 四、exe 的正确发布方式（不要提交进仓库）⚠️

- [ ] ⚠️ **不要把 `dist/` 提交进 Git 仓库**
  - 现在 `dist/` 有 55 MB，且每次改代码都会变
  - 提交进去会让仓库迅速膨胀，且无法有效清理历史
- [ ] **改用 GitHub Releases**
  - 打一个 tag（如 `v1.0.0`），在 Releases 里上传 exe
  - README 里放下载链接
- [ ] **Release 说明里写清**
  - 版本号、变更内容、系统要求（64 位 Windows 7+）
  - 未签名提示的处理方法（"更多信息"→"仍要运行"）
  - 对应源码的 commit 链接（GPL 义务）

---

## 五、代码质量与可信度（强烈建议）

- [ ] **`tests/` 目录收拢自检脚本**，改成 pytest 风格或保留独立可运行
  - 现在有 `selftest_synthetic.py`（36 项）、`selftest_platform.py`（18 项）、
    `main.py --selftest`（11 项）
- [ ] **GitHub Actions 自动跑自检**
  见下方"附：CI 配置模板"
- [ ] **删除开发期的诊断脚本**（带下划线前缀的，对用户无意义）
  - `_check_structure.py`、`_verify_with_word.ps1`
- [ ] **README 加截图或动图**
  - GitHub 上有没有截图，直接影响别人是否愿意下载
- [ ] **写清适用边界，不要夸大**
  - 强项：结构较清晰的论文（有编号、有标题层级）
  - 弱项：完全没有编号的纯文本稿（识别会退化）
  - 已实测：识别精确率 86%、召回率 97%（以 Word 大纲级别为基准）

---

## 六、规范来源的说明（避免争议）

- [ ] **说明 `specs/sysu_undergraduate_thesis.json` 的来源**
  - 它是从中山大学本科毕业论文格式要求翻译来的
  - 建议 README 中注明来源，并说明"规则的具体数值是客观事实，
    此处以机器可读形式重新表达"
- [ ] **说明如何支持其他学校**
  - 只需按同样结构新增一个 JSON，程序不用改
  - 这是这个项目最有价值的扩展点，值得在 README 里突出

---

## 七、术语与定位

- [ ] **明确它不是"插件"**
  - 它是独立桌面程序，不是某个软件的扩展
  - README 标题和描述里避免用"插件"，否则找的人预期错位
- [ ] **仓库名建议**
  - `thesis-format-normalizer`
  - `word-thesis-formatter`
  - `论文格式改写器`（中文名也可以，但英文名更利于搜索）
- [ ] **GitHub 仓库描述（About）一句话**
  - 例：`把 Word 论文按格式规范重新排版：自动识别标题层级，套用字体字号，支持自定义预设`

---

## 附：`.gitignore` 模板

```gitignore
# 构建产物
dist/
build/
*.spec
__pycache__/
*.py[cod]
*.egg-info/

# 用户生成物（不应发布）
预设.json
本次使用的预设.json
自检报告.txt
*_已排版.docx
*_修复版.docx

# 论文样本（含个人信息，绝不提交）
*.docx
*.doc
!docs/sample-template.docx   # 若确实需要示例，放一个空白模板到此路径

# 编辑器与系统
.vscode/
.idea/
.DS_Store
Thumbs.db
desktop.ini

# 临时文件
_*.py
_*.ps1
_*.docx
发给Mac用户/
*_Mac版.zip
```

> 注意：`*.docx` 全局忽略是**故意的**。这个项目处理的就是 docx，
> 很容易误提交用户论文。若需要示例文档，用白名单单独放行。

---

## 附：CI 配置模板（`.github/workflows/selftest.yml`）

```yaml
name: 自检

on:
  push:
  pull_request:

jobs:
  selftest:
    runs-on: windows-latest
    strategy:
      matrix:
        python-version: ["3.10", "3.12"]
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: ${{ matrix.python-version }}
      - name: 安装依赖
        run: |
          python -m pip install --upgrade pip
          pip install python-docx lxml
      - name: 假想论文端到端自检
        run: python selftest_synthetic.py
      - name: 跨平台字体自检
        run: python selftest_platform.py
      - name: 程序自检
        run: python main.py --selftest
```

---

## 附：建议的仓库结构

```
.
├── README.md                  中文主文档
├── README.en.md               英文版（可选，利于国际可见度）
├── LICENSE                    GPL-3.0 全文
├── requirements.txt
├── .gitignore
├── .github/workflows/selftest.yml
│
├── 论文格式改写器/             主程序包（或平铺也行）
│   ├── 启动.py
│   ├── main.py
│   ├── app.py                 图形界面
│   ├── rewriter.py            改写引擎
│   ├── region_detector.py     区域识别
│   ├── numbering.py           编号处理
│   ├── superscript.py         上标处理
│   ├── platform_fonts.py      跨平台字体
│   └── rewrite_cli.py         命令行入口
│
├── specs/
│   └── sysu_undergraduate_thesis.json
│
├── tests/
│   ├── selftest_synthetic.py
│   └── selftest_platform.py
│
├── packaging/
│   ├── build_exe.py
│   ├── 一键改写.bat
│   └── 启动.command
│
└── docs/
    ├── 分发说明.md
    └── Mac使用说明.md
```

---

## 发布前的最后三步

- [ ] **在干净目录重新验证一遍**
  ```powershell
  git clone <你的仓库> /tmp/checkout
  cd /tmp/checkout
  pip install -r requirements.txt
  python tests/selftest_synthetic.py
  ```
  确认脱离你的开发环境也能跑通。
- [ ] **用 GitHub 的搜索确认没有敏感信息漏网**
  在仓库页面按 `t` 打开文件搜索，搜 `Administrator`、`wxid`、`大学文件`。
- [ ] **打 tag 发 Release**，上传 exe，附上源码链接（GPL 义务）。
