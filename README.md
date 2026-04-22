<h1 align="center">Docxnote</h1>

<p align="center">
  <strong>Lightweight DOCX comment engine: add and read Word comments from plain paragraph text, with only an <code>lxml</code> dependency.</strong>
</p>

<p align="center">
  <a href="https://www.python.org/downloads/"><img src="https://img.shields.io/badge/python-3.12+-blue.svg?style=for-the-badge&logo=python" alt="Python 3.12+"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-Apache%202.0-blue.svg?style=for-the-badge" alt="License: Apache 2.0"></a>
  <a href="https://pypi.org/project/docxnote/"><img src="https://img.shields.io/pypi/v/docxnote.svg?style=for-the-badge&logo=pypi&logoColor=white&label=pypi" alt="PyPI version"></a>
  <a href="https://github.com/touken928/docxnote/stargazers"><img src="https://img.shields.io/github/stars/touken928/docxnote?style=for-the-badge&color=yellow&logo=github" alt="GitHub stars"></a>
</p>

<p align="center">
  <a href="docs/README_zh.md">简体中文</a>
</p>

---

## Overview

**docxnote** automates **Word comments** (annotations) without touching runs: you traverse `Paragraph` / `Table` / `Cell`, call `paragraph.comment(...)`, and optionally read comments via `paragraph.comments` and `doc.comments()`.

**Repository:** [touken928/docxnote](https://github.com/touken928/docxnote)

---

## Installation

```
pip install docxnote
```

With [uv](https://github.com/astral-sh/uv):

```
uv add docxnote
```

---

## Quick start

```python
from docxnote import DocxDocument, Paragraph, Table

# Load document
with open("document.docx", "rb") as f:
    # By default existing comments are discarded
    doc = DocxDocument.parse(f.read())

    # To keep existing comments and add more:
    # doc = DocxDocument.parse(f.read(), keep_comments=True)

# Walk block-level content
for block in doc.blocks():
    if isinstance(block, Paragraph):
        if block.text:
            block.comment("Please review wording", end=5, author="reviewer")

    elif isinstance(block, Table):
        rows, cols = block.shape()
        for r in range(rows):
            for c in range(cols):
                cell = block[r, c]
                for inner in cell.blocks():
                    if isinstance(inner, Paragraph) and inner.text:
                        inner.comment("Needs review", end=3, author="reviewer")

# Write output
output = doc.render()
with open("output.docx", "wb") as f:
    f.write(output)
```

---

## CLI

Installing `docxnote` adds a console script of the same name. Every read
command supports `--json`, every write command takes an explicit output path,
and they all share the library's addressable paths (`p:0`, `t:0/r:1/c:2/p:0`,
`p:0#3`, ...).

```
docxnote list input.docx --text --json
docxnote show input.docx "t:0/r:1/c:2/p:0"
docxnote comments input.docx --json
docxnote annotate input.docx output.docx --path p:0 --text "please revise"
docxnote annotate input.docx output.docx --spec ops.json --keep-comments
```

Full reference: [docs/CLI.md](docs/CLI.md) · [docs/CLI_zh.md](docs/CLI_zh.md).

---

## Documentation

Full Python API (methods, parameters, comments, paths, tables, and advanced patterns):

- [docs/API.md](docs/API.md) — English  
- [docs/API_zh.md](docs/API_zh.md) — 简体中文  
- [docs/CLI.md](docs/CLI.md) / [docs/CLI_zh.md](docs/CLI_zh.md) — CLI reference  
- [docs/README_zh.md](docs/README_zh.md) — 简体中文 overview (same scope as README)

---

## Tests

Test documents are generated with python-docx; no checked-in DOCX fixtures. See [tests/README.md](tests/README.md).

---

## SKILL

Agent-oriented docs live under [`skills`](https://github.com/touken928/docxnote/tree/main/skills) in [github.com/touken928/docxnote](https://github.com/touken928/docxnote). To mirror them under `.cursor/`, `.claude/`, or similar, **tell your coding agent to use that repository’s `skills` folder** and let it copy or fetch as needed.

---

## License

Licensed under the Apache License, Version 2.0. See `LICENSE` in the repository root.
