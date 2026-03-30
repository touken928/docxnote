# Docxnote

**[简体中文](README_zh-CN.md)**

**docxnote** is a lightweight **DOCX comment engine** with only an `lxml` dependency, for automating Word comments.

It works directly on **WordprocessingML**, treating a DOCX as **ZIP + XML**, and exposes a **text-oriented API**.

Unlike typical DOCX libraries, docxnote **hides Word’s Run structure entirely**; everything is based on **paragraph strings**.

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

## API

### DocxDocument

Represents a DOCX file.

#### parse

```python
DocxDocument.parse(docx_bytes, *, keep_comments=False)
```

Parses the DOCX and returns a document object.

- **keep_comments**: Whether to keep existing comments. Default `False` (strips them). Use `True` to preserve existing comments and append new ones.

---

#### blocks

```python
doc.blocks()
```

Returns block-level elements:

```python
(Paragraph | Table, ...)
```

Order matches the Word document.

---

#### render

```python
doc.render()
```

Returns new DOCX as `bytes`. Comments are written during this step.

#### Thread safety

A single `DocxDocument` instance is safe to use from multiple threads (internally serialized with a reentrant lock). Use separate instances for parallel work across threads. For multiple processes, call `parse` in each process.

---

### Paragraph

Represents a Word paragraph.

#### text

```python
text = paragraph.text
```

Full paragraph text, including `\n` and `\t`.

---

#### comment

```python
paragraph.comment(
    text,           # comment body
    start=0,        # start index (inclusive)
    end=None,       # end index (exclusive); None means end of paragraph
    *,
    author="docxnote",
    date=None,      # datetime (timezone-aware recommended); None = current system time
)
```

Adds a comment spanning the given character range in the paragraph. The `w:date` value in `comments.xml` is stored in UTC (`…Z`). A naive `datetime` (no `tzinfo`) is treated as UTC.

**Example:**

```python
paragraph.comment("Needs change", start=3, end=8, author="Alice")
```

docxnote handles run splitting, anchors, `comments.xml`, and relationship updates.

---

### Table

Represents a Word table.

#### shape

```python
rows, cols = table.shape()
```

Returns `(row_count, col_count)`.

---

#### Cell access

```python
cell = table[row, col]
```

Returns a `Cell`. All coordinates are addressable, including positions covered by merged cells.

---

### Cell

Represents a table cell.

#### blocks

```python
cell.blocks()
```

Block-level elements inside the cell:

```python
(Paragraph | Table, ...)
```

Order matches Word.

---

#### bounds

```python
top, left, bottom, right = cell.bounds()
```

Cell bounds `(top, left, bottom, right)` with half-open intervals `[top, bottom)` and `[left, right)`.

For a non-merged cell, returns `(r, c, r+1, c+1)`.

---

## Advanced

### Nested tables

```python
for block in doc.blocks():
    if isinstance(block, Table):
        rows, cols = block.shape()
        for r in range(rows):
            for c in range(cols):
                cell = block[r, c]
                for inner_block in cell.blocks():
                    if isinstance(inner_block, Table):
                        inner_rows, inner_cols = inner_block.shape()
                        # ...
```

### Multiple comments

```python
paragraph.comment("Note 1", start=0, end=5, author="Alice")
paragraph.comment("Note 2", start=10, end=15, author="Bob")
paragraph.comment("Note 3", start=20, end=25, author="Carol")
```

### Merged cells

```python
table = [b for b in doc.blocks() if isinstance(b, Table)][0]

cell = table[0, 0]
top, left, bottom, right = cell.bounds()

if bottom - top > 1 or right - left > 1:
    print(f"Merged cell spans {bottom - top} rows, {right - left} cols")
```

---

## Tests

Test documents are generated with python-docx; no checked-in DOCX fixtures. See [tests/README.md](tests/README.md).

---

## SKILL

This repo includes [`SKILL.md`](SKILL.md) to help coding agents use `docxnote` correctly. From the project root, download it into the agent skill folder (replace `main` if your default branch differs). On Windows PowerShell, use `curl.exe` if `curl` is aliased to `Invoke-WebRequest`.

**Cursor**

```bash
mkdir -p .cursor/docxnote
curl -fsSL -o .cursor/docxnote/SKILL.md https://raw.githubusercontent.com/touken928/docxnote/main/SKILL.md
```

**Claude Code**

```bash
mkdir -p .claude/docxnote
curl -fsSL -o .claude/docxnote/SKILL.md https://raw.githubusercontent.com/touken928/docxnote/main/SKILL.md
```

Point your agent at that file for install steps, suggested patterns, and caveats.

---

## License

Licensed under the Apache License, Version 2.0. See `LICENSE` in the repository root.
