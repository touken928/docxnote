<h1 align="center">Docxnote</h1>

<p align="center">
  <strong>轻量级 DOCX 批注引擎：在段落纯文本上添加与读取 Word 批注，仅依赖 <code>lxml</code>。</strong>
</p>

<p align="center">
  <a href="https://www.python.org/downloads/"><img src="https://img.shields.io/badge/python-3.12+-blue.svg?style=for-the-badge&logo=python" alt="Python 3.12+"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-Apache%202.0-blue.svg?style=for-the-badge" alt="License: Apache 2.0"></a>
  <a href="https://pypi.org/project/docxnote/"><img src="https://img.shields.io/pypi/v/docxnote.svg?style=for-the-badge&logo=pypi&logoColor=white&label=pypi" alt="PyPI version"></a>
  <a href="https://github.com/touken928/docxnote/stargazers"><img src="https://img.shields.io/github/stars/touken928/docxnote?style=for-the-badge&color=yellow&logo=github" alt="GitHub stars"></a>
</p>

<p align="center">
  <a href="README.md">English</a> &middot; 简体中文
</p>

---

## 概览

**docxnote** 用于自动化 **Word 批注**：遍历 `Paragraph` / `Table` / `Cell`，用 `paragraph.comment(...)` 添加批注，并可通过 `paragraph.comments` 与 `doc.comments()` 读取批注。

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
    # 默认不保留原有批注（会清空）
    doc = DocxDocument.parse(f.read())

    # 如需保留原有批注并继续添加：
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

---

## 文档

完整 Python API（方法、参数、批注、路径、表格与高级用法）：

- [docs/API_zh.md](docs/API_zh.md) — 简体中文  
- [docs/API.md](docs/API.md) — English  

索引：[docs/README.md](docs/README.md)。

---

## 测试

所有测试文档使用 python-docx 动态生成，不依赖外部文件，详见 [tests/README.md](tests/README.md)。

---

## SKILL

[github.com/touken928/docxnote](https://github.com/touken928/docxnote) 根目录下的 [`skills`](https://github.com/touken928/docxnote/tree/main/skills) 文件夹提供供对话型 / coding Agent 使用的技能说明。若需同步到本机 `.cursor/`、`.claude/` 等目录，**告知对话智能体从上述仓库获取 `skills` 文件夹**，由其自行复制或下载即可。

---

## 许可证

本项目采用 Apache License 2.0 许可证。详见仓库根目录的 `LICENSE`。
