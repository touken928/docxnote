<h1 align="center">Docxnote</h1>

<p align="center">
  <a href="https://www.python.org/downloads/"><img src="https://img.shields.io/badge/python-3.12+-blue.svg?style=for-the-badge&logo=python" alt="Python 3.12+"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-Apache%202.0-blue.svg?style=for-the-badge" alt="License: Apache 2.0"></a>
  <a href="https://pypi.org/project/docxnote/"><img src="https://img.shields.io/pypi/v/docxnote.svg?style=for-the-badge&logo=pypi&logoColor=white&label=pypi" alt="PyPI version"></a>
  <a href="https://github.com/touken928/docxnote/stargazers"><img src="https://img.shields.io/github/stars/touken928/docxnote?style=for-the-badge&color=yellow&logo=github" alt="GitHub stars"></a>
</p>

<p align="center">
  Read and add Word comments using plain paragraph text and character offsets.
  Runtime dependency: <code>lxml</code>.
</p>

<p align="center">
  <a href="docs/README_zh.md">简体中文</a> · <a href="docs/API.md">Python API</a> · <a href="docs/SHELL.md">Document shell</a> · <a href="tests/README.md">Development and tests</a>
</p>

## Installation

```bash
pip install docxnote
# Or, in a uv project:
uv add docxnote
```

## Quick start

Read an existing DOCX, annotate the first nonempty paragraph, and save a copy:

```python
from pathlib import Path
from docxnote import DocxDocument

doc = DocxDocument.parse(Path("input.docx").read_bytes(), keep_comments=True)

for paragraph in doc.iter_paragraphs():
    if paragraph.text:
        paragraph.comment("Please review this wording", author="reviewer")
        break

Path("reviewed.docx").write_bytes(doc.render())
```

`keep_comments=True` preserves existing comments and their metadata. The default
is `False`, which strips existing comments. File reading and writing belong to
your application.

## Choose your interface

| Task | Start here |
| --- | --- |
| Read text, inspect tables, or add comments from Python | [Python API](docs/API.md) |
| Search and annotate with a bounded command interface, including agent tools | [DocxShell guide](docs/SHELL.md) |
| Run checks or contribute a fix | [Development and tests](tests/README.md) |

Paragraph traversal includes tables, nested tables, and block content controls,
with merged cells visited once. Comment offsets use `[start, end)` in
`paragraph.text`. Text boxes are preserved but excluded from this text view;
comment range reading supports single-paragraph ranges. See
[text scope](docs/API.md#text) and [range limits](docs/API.md#single-paragraph-scope).

## Documentation

| Reference | English | 简体中文 |
| --- | --- | --- |
| Installation and quick start | This page | [入门](docs/README_zh.md) |
| Python API and behavior | [API](docs/API.md) | [API 参考](docs/API_zh.md) |
| Document shell and integration | [Shell](docs/SHELL.md) | [Shell 指南](docs/SHELL_zh.md) |

Repository maintenance rules are in [AGENTS.md](AGENTS.md).
