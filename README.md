# kaoyan-bilingual-docx

把**英文试题 PDF + 参考答案 PDF**，整理成一份**从头到尾、一句英文一句中文**的逐句对照
Word 文档（.docx）——完形填空按答案填空、阅读题带选项与正解、翻译用官方参考译文、
写作含范文。专为**考研英语（英语一/二）**、四六级、托福雅思等英文真题的「中英对照学习资料」制作。

> Turn an English exam PDF + its answer key into a clean, sentence-by-sentence
> bilingual (English→Chinese) Word document. Built for Chinese postgraduate
> entrance English exams and similar reading-based tests.

排版已经固定好、舒适不拥挤：完形选项 4 个并排、阅读选项 2 个并排、正确答案红色 ✔、
Part B 题号标注到选项后、中文统一深蓝色。**你只管内容，排版永远一致。**

![示例](examples/example_output.docx) ← 仓库里 `examples/example_output.docx` 是一份渲染样例。

---

## 它是怎么工作的

```
试题PDF + 答案PDF
      │  ① 抽取文本（extract_pdf_text.py）
      ▼
  纯文本（通读、切句）
      │  ② 逐句翻译 + 组装成 spec.json  ←★ 需要语言模型理解力
      ▼
   spec.json（内容规格）
      │  ③ 渲染（render_bilingual_docx.py，排版固定）
      ▼
  中英对照 .docx
```

**①③ 是确定性的纯脚本，谁都能跑。② 翻译切句这一步需要 AI**——
推荐用 [Claude Code](https://claude.com/claude-code) 安装本仓库自带的 skill 后直接对话完成（见下）。

---

## 安装

需要 Python 3.9+。

```bash
git clone https://github.com/LJL-6666/kaoyan-bilingual-docx.git
cd kaoyan-bilingual-docx
pip install -r requirements.txt
```

---

## 用法一：作为 Claude Code 技能（推荐，全自动）

本仓库的 `skill/` 目录就是一个 Claude Code 技能。安装后，给 Claude 两个 PDF，
说一句「把这套英语真题做成中英逐句对照 Word」，它就会自动走完抽取→翻译→渲染。

```bash
# 把技能装到 Claude Code 的个人技能目录
# macOS / Linux:
cp -r skill ~/.claude/skills/bilingual-exam-doc
# Windows (PowerShell):
Copy-Item -Recurse skill "$env:USERPROFILE\.claude\skills\bilingual-exam-doc"
```

然后在 Claude Code 里直接说：

> 把 `2025真题.pdf` 和 `2025答案.pdf` 做成中英逐句对照的 Word。

技能的全部约定（完形填空填答案、阅读选项排版、Part B 各题型如何处理等）都写在
`skill/SKILL.md` 里，Claude 会照着做。

---

## 用法二：当作独立工具手动跑（不依赖 AI）

如果你自己写好了 `spec.json`（内容规格，格式见下），只用渲染器就能出 docx：

```bash
# 抽取 PDF 文本
python make_bilingual.py extract 真题.pdf 真题.txt

# 校验 spec、看各类块数量
python make_bilingual.py check examples/example_spec.json

# 渲染成 Word
python make_bilingual.py render examples/example_spec.json 输出.docx
```

> Windows 下若终端中文乱码，命令前加 `set PYTHONIOENCODING=utf-8`（cmd）
> 或 `$env:PYTHONIOENCODING="utf-8"`（PowerShell）。
> 渲染时若目标 docx 正被 Word 打开会失败，先关闭它。

---

## spec.json 内容规格

一个最小可渲染的例子见 [`examples/example_spec.json`](examples/example_spec.json)。
顶层字段：`title`、`subtitle`、`note`、`blocks`。`blocks` 是按顺序排列的内容块，常用类型：

| 块类型 | 用途 |
|---|---|
| `pair` | 一句英文 + 一句中文（完形/阅读/翻译/范文本文都用它） |
| `cloze_opts` | 完形某题的 4 个选项（4 个并排，正确项标红 ✔） |
| `qhead` + `read_opts` | 阅读题干 + 选项（每行 2 个，正确项标红 ✔） |
| `sub` | Part B 小标题/匹配项，题号标注到选项后 |
| `numbered` | Part B 判断 T/F、排序等；带「编号 + 答案」 |
| `h1`/`h2`/`h3`、`note` | 分节标题与小字说明 |

每个块的完整字段说明，见 `skill/scripts/render_bilingual_docx.py` 文件顶部的 docstring。

---

## 目录结构

```
kaoyan-bilingual-docx/
├── README.md
├── requirements.txt
├── make_bilingual.py            # 便捷 CLI：extract / check / render
├── skill/                       # Claude Code 技能（核心）
│   ├── SKILL.md                 # 制作约定与流程
│   ├── scripts/
│   │   ├── extract_pdf_text.py  # 工具①：PDF → 文本
│   │   └── render_bilingual_docx.py  # 工具②：spec.json → docx
│   └── references/example_spec.json
└── examples/
    ├── example_spec.json        # 样例内容规格
    └── example_output.docx      # 样例渲染结果
```

---

## 已知限制 / 注意

- **翻译质量**：本工具产出的中文译文（阅读/完形本文部分）由模型生成，建议作为学习参考，
  关键处请对照原文与官方解析核对。
- **扫描图片版 PDF**：无文字层的扫描件，`extract` 抽不出文字，需要先 OCR（如 Claude 看图识字，
  或 tesseract）。
- **答案与官方译文**：以你提供的答案/解析 PDF 为准；OCR 可能有个别讹误，请抽查。

## License

[MIT](LICENSE)
