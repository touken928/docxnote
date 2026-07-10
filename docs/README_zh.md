<h1 align="center">Docxnote</h1>

<p align="center">
  <strong>轻量级 DOCX 批注引擎：在段落纯文本上添加与读取 Word 批注，仅依赖 <code>lxml</code>。</strong>
</p>

<p align="center">
  <a href="https://www.python.org/downloads/"><img src="https://img.shields.io/badge/python-3.12+-blue.svg?style=for-the-badge&logo=python" alt="Python 3.12+"></a>
  <a href="../LICENSE"><img src="https://img.shields.io/badge/license-Apache%202.0-blue.svg?style=for-the-badge" alt="License: Apache 2.0"></a>
  <a href="https://pypi.org/project/docxnote/"><img src="https://img.shields.io/pypi/v/docxnote.svg?style=for-the-badge&logo=pypi&logoColor=white&label=pypi" alt="PyPI version"></a>
  <a href="https://github.com/touken928/docxnote/stargazers"><img src="https://img.shields.io/github/stars/touken928/docxnote?style=for-the-badge&color=yellow&logo=github" alt="GitHub stars"></a>
</p>

<p align="center">
  <a href="../README.md">English</a> &middot; 简体中文
</p>

---

## 概览

**docxnote** 用于自动化 **Word 批注**：遍历 `Paragraph` / `Table` / `Cell`，用 `paragraph.comment(...)` 添加批注，并可通过 `paragraph.comments` 与 `doc.comments()` 读取批注，而不需要手工处理 Run。

**仓库：** [touken928/docxnote](https://github.com/touken928/docxnote)

---

## 安装

```
pip install docxnote
```

使用 [uv](https://github.com/astral-sh/uv)：

```
uv add docxnote
```

---

## 快速开始

```python
from docxnote import DocxDocument, Paragraph, Table

# 读取文档
with open("document.docx", "rb") as f:
    # 默认会先剥离原有批注，再写入新批注
    doc = DocxDocument.parse(f.read())

    # 如需保留原有批注并继续追加：
    # doc = DocxDocument.parse(f.read(), keep_comments=True)

# 遍历文档块
for block in doc.blocks():
    if isinstance(block, Paragraph):
        # 为段落添加批注
        if block.text:
            block.comment("请检查表述", end=5, author="reviewer")

    elif isinstance(block, Table):
        # 处理表格
        rows, cols = block.shape()
        for r in range(rows):
            for c in range(cols):
                cell = block[r, c]
                # 为单元格内容添加批注
                for inner in cell.blocks():
                    if isinstance(inner, Paragraph) and inner.text:
                        inner.comment("需复核", end=3, author="reviewer")

# 生成新文档
output = doc.render()
with open("output.docx", "wb") as f:
    f.write(output)
```

表格遍历会正确处理 Word 行中省略的首尾网格列。批注范围基于
`paragraph.text` 遵循 Python 切片语义，包括负数和超出范围的端点；规范化后
反向范围会成为锚定在规范化起点的空范围。详见 [API 参考](API_zh.md)。

---

## 命令行

安装 `docxnote` 会注册同名命令。所有读命令支持 `--json`，所有写命令必须显式指定
输出文件，全部子命令使用与库一致的可寻址路径（`p:0` / `t:0/r:1/c:2/p:0` / `p:0#3` …）。

```
docxnote list input.docx --text --json
docxnote show input.docx "t:0/r:1/c:2/p:0"
docxnote comments input.docx --json
docxnote annotate input.docx output.docx --path p:0 --text "请修改"
docxnote annotate input.docx output.docx --spec ops.json --keep-comments
```

完整说明：[CLI_zh.md](CLI_zh.md) · [CLI.md](CLI.md)。

`annotate` 会拒绝包括符号链接和硬链接在内的输入/输出文件别名，并先写入同目录
临时文件，再原子替换 `OUTPUT`。详见 [CLI 参考](CLI_zh.md)。

`start` / `end` 始终基于 `paragraph.text` 的字符偏移，遵循 Python 切片语义
`[start, end)`。docxnote 会自动处理 Run 拆分，包括超链接等嵌套段落内容中的精确锚点。

---

## 文档

完整 Python API（方法、参数、批注、路径、表格与高级用法）：

- [API_zh.md](API_zh.md) — 简体中文  
- [API.md](API.md) — English  
- [CLI_zh.md](CLI_zh.md) / [CLI.md](CLI.md) — CLI 参考  

当 `keep_comments=True` 时，原有批注会连同已有的批注 XML 元数据一起保留，并在其基础上追加新批注。

---

## 测试

所有测试文档使用 python-docx 动态生成，不依赖外部文件，详见 [tests/README.md](../tests/README.md)。

---

## SKILL

```
npx skills add touken928/docxnote
```

[github.com/touken928/docxnote](https://github.com/touken928/docxnote) 根目录下的 [`skills/docxnote`](https://github.com/touken928/docxnote/tree/main/skills/docxnote) 文件夹提供供对话型 / coding Agent 使用的技能说明。

---

## 许可证

本项目采用 Apache License 2.0 许可证。详见仓库根目录的 `LICENSE`。
