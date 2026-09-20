<h1 align="center">Docxnote</h1>

<p align="center">
  <a href="https://www.python.org/downloads/"><img src="https://img.shields.io/badge/python-3.12+-blue.svg?style=for-the-badge&logo=python" alt="Python 3.12+"></a>
  <a href="../LICENSE"><img src="https://img.shields.io/badge/license-Apache%202.0-blue.svg?style=for-the-badge" alt="License: Apache 2.0"></a>
  <a href="https://pypi.org/project/docxnote/"><img src="https://img.shields.io/pypi/v/docxnote.svg?style=for-the-badge&logo=pypi&logoColor=white&label=pypi" alt="PyPI version"></a>
  <a href="https://github.com/touken928/docxnote/stargazers"><img src="https://img.shields.io/github/stars/touken928/docxnote?style=for-the-badge&color=yellow&logo=github" alt="GitHub stars"></a>
</p>

基于段落纯文本和字符偏移，读取与添加 Word 批注。
运行时仅依赖 `lxml`。

[English](../README.md) · [Python API](API_zh.md) · [文档 Shell](SHELL_zh.md) · [开发与测试](../tests/README.md)

## 安装

```bash
pip install docxnote
# 或在 uv 项目中：
uv add docxnote
```

## 快速开始

读取已有 DOCX，为第一个非空段落添加批注，并另存为新文件：

```python
from pathlib import Path
from docxnote import DocxDocument

doc = DocxDocument.parse(Path("input.docx").read_bytes(), keep_comments=True)

for paragraph in doc.iter_paragraphs():
    if paragraph.text:
        paragraph.comment("请检查这段表述", author="reviewer")
        break

Path("reviewed.docx").write_bytes(doc.render())
```

`keep_comments=True` 保留旧批注及其元数据；默认值为 `False`，会剥离旧批注。
文件读取与保存由调用方负责。

## 选择接口

| 任务 | 从这里开始 |
| --- | --- |
| 用 Python 读取文本、检查表格或添加批注 | [Python API](API_zh.md) |
| 通过有输出上限的命令接口检索、批注，或接入 Agent 工具 | [DocxShell 指南](SHELL_zh.md) |
| 运行检查或贡献修复 | [开发与测试](../tests/README.md) |

段落遍历包含表格、嵌套表格和块级内容控件，合并单元格只访问一次。
批注偏移基于 `paragraph.text` 的 `[start, end)` 区间。文本框会保留，但不计入
该文本视图；批注范围读取仅支持单一段落。详见
[文本范围](API_zh.md#text)与[范围限制](API_zh.md#单段范围限制)。

## 文档

| 参考 | English | 简体中文 |
| --- | --- | --- |
| 安装与快速开始 | [Getting started](../README.md) | 本页 |
| Python API 与行为 | [API](API.md) | [API 参考](API_zh.md) |
| 文档 Shell 与接入 | [Shell](SHELL.md) | [Shell 指南](SHELL_zh.md) |

仓库维护约定见 [AGENTS.md](../AGENTS.md)。
