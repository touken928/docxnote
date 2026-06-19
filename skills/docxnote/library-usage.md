# docxnote Python library guide for agents

This guide is for coding agents that need to integrate `docxnote` into Python
 code. It focuses on practical workflows, safe patterns, and copyable examples.
The authoritative API reference is `docs/API.md`; this file is the agent-facing
playbook.

---

## Key ideas

- A DOCX file is treated as `ZIP + XML`, but the public API stays high-level.
- You work with full paragraph text, not Word run internals.
- Comments are accumulated in memory and written only when `doc.render()` is
  called.
- One `DocxDocument` instance is safe to share across threads.

## Install and import

Prefer one of these install commands when generating setup steps:

```bash
pip install docxnote
```

or:

```bash
uv add docxnote
```

In Python:

```python
from docxnote import DocxDocument, Paragraph, Table
```

## Recommended workflow

1. Parse the DOCX bytes.
2. Traverse paragraphs and tables.
3. Add or inspect comments through paragraph text.
4. Render a new DOCX and write it out.

```python
from docxnote import DocxDocument, Paragraph, Table


def annotate_docx(input_path: str, output_path: str, *, keep_comments: bool = False) -> None:
    with open(input_path, "rb") as f:
        doc = DocxDocument.parse(f.read(), keep_comments=keep_comments)

    for block in doc.blocks():
        if isinstance(block, Paragraph):
            handle_paragraph(block)
        elif isinstance(block, Table):
            handle_table(block)

    with open(output_path, "wb") as f:
        f.write(doc.render())


def handle_paragraph(paragraph: Paragraph) -> None:
    text = paragraph.text or ""
    if "TODO" in text:
        paragraph.comment(
            "Please resolve this TODO.",
            start=0,
            end=min(5, len(text)),
            author="reviewer",
        )


def handle_table(table: Table) -> None:
    rows, cols = table.shape()
    for r in range(rows):
        for c in range(cols):
            cell = table[r, c]
            for inner in cell.blocks():
                if isinstance(inner, Paragraph):
                    handle_paragraph(inner)
                elif isinstance(inner, Table):
                    handle_table(inner)
```

## `DocxDocument`

### Parse

```python
doc = DocxDocument.parse(docx_bytes, keep_comments=False)
```

- `keep_comments=False` strips existing comments and starts fresh.
- `keep_comments=True` preserves existing comments and lets you append new ones.

If the user says "append", "keep existing comments", or "build on reviewed
files", explicitly pass `keep_comments=True`.

### Traverse top-level blocks

```python
for block in doc.blocks():
    if isinstance(block, Paragraph):
        ...
    elif isinstance(block, Table):
        ...
```

The order matches Word's visible document order.

## `Paragraph`

### Text view

```python
text = paragraph.text
```

- It is the full paragraph string.
- It preserves `\n` and `\t`.
- All comment offsets are based on this exact string.

### Read paragraph comments

```python
from docxnote import Comment

for c in paragraph.comments:
    assert isinstance(c, Comment)
    print(c.path, c.start, c.end, c.text, c.author)
```

- `c.path` is a stable address like `t:0/r:0/c:0/p:0#3`.
- `c.start` / `c.end` are character offsets into `paragraph.text` using
  Python slice semantics: `[start, end)`.
- `c.text` is the comment body.
- `c.date` comes from the comment's `w:date` value.

### Add a comment

```python
new = paragraph.comment(
    text="Needs revision",
    start=0,
    end=None,
    author="docxnote",
    date=None,
)
```

- `end=None` means "to the end of the paragraph".
- Returned `Comment.path` can be stored and later resolved with
  `doc.resolve(...)`.
- Offsets are character-based, not byte-based.

Typical precise-anchor pattern:

```python
phrase = "risk"
idx = paragraph.text.find(phrase)
if idx != -1:
    paragraph.comment(
        "Please justify this risk statement.",
        start=idx,
        end=idx + len(phrase),
        author="reviewer",
    )
```

Whole-paragraph comment:

```python
if paragraph.text.strip():
    paragraph.comment("Please review this paragraph.", author="reviewer")
```

Multiple comments on one paragraph:

```python
paragraph.comment("Comment 1", start=0, end=5, author="A")
paragraph.comment("Comment 2", start=10, end=15, author="B")
```

## `Table` and `Cell`

### Traverse a table

```python
rows, cols = table.shape()
for r in range(rows):
    for c in range(cols):
        cell = table[r, c]
        ...
```

### Traverse cell content

```python
for inner in cell.blocks():
    if isinstance(inner, Paragraph):
        ...
    elif isinstance(inner, Table):
        ...
```

Use the same recursion strategy for nested tables.

### Handle merged cells

```python
top, left, bottom, right = cell.bounds()
```

- Bounds use half-open ranges: `[top, bottom)` and `[left, right)`.
- An unmerged cell at `(r, c)` has bounds `(r, c, r + 1, c + 1)`.
- Accessing a coordinate covered by a merge still returns the logical cell view.

## Document-wide comment access

```python
for c in doc.comments():
    ...
```

This walks the whole document, including paragraphs inside tables and nested
tables, in document order.

## Stable paths

Every `Paragraph`, `Table`, `Cell`, and `Comment` has a stable `path` string.

Examples:

- `p:0`
- `t:0`
- `t:0/r:1/c:2`
- `t:0/r:1/c:2/p:0`
- `t:0/r:1/c:2/p:0#3`

Useful APIs:

```python
para = doc.resolve("p:0")
cell = doc.resolve("t:0/r:1/c:2")
comment = doc.resolve("t:0/r:1/c:2/p:0#3")

for p in doc.iter_paragraphs():
    print(p.path)
```

## Recommended patterns

### Rule-based review

```python
def apply_rules_to_paragraph(paragraph: Paragraph) -> None:
    text = paragraph.text or ""
    if not text.strip():
        return

    if len(text) > 200:
        paragraph.comment("Paragraph is too long; consider splitting it.")

    for term in ("significantly improved", "dramatically reduced"):
        idx = text.find(term)
        if idx != -1:
            paragraph.comment(
                f"Please provide quantitative support for '{term}'.",
                start=idx,
                end=idx + len(term),
                author="review-bot",
            )
```

### Skip blank paragraphs

```python
if not (paragraph.text or "").strip():
    return
```

## Common mistakes

- Do not manipulate runs or OOXML directly unless absolutely necessary.
- Do not forget to call `doc.render()`; comments only exist in memory before
  render.
- Do not compute offsets from bytes; use normal Python string indexing.

## When to prefer docxnote

Choose `docxnote` when the task involves:

- Automated comments on many DOCX files.
- Paragraph-string-driven logic instead of run-level editing.
- Comments inside tables, including merged and nested tables.
- Appending machine comments on top of existing human comments.

## Quick reference

- Parse: `DocxDocument.parse(bytes, keep_comments=False)`
- Top-level blocks: `doc.blocks()`
- Paragraph text: `paragraph.text`
- Add comment: `paragraph.comment(...)`
- Paragraph comments: `paragraph.comments`
- All comments: `doc.comments()`
- Resolve by path: `doc.resolve("p:0")`
- Iterate all paragraphs: `doc.iter_paragraphs()`
- Table shape: `table.shape()`
- Cell access: `table[r, c]`
- Cell blocks: `cell.blocks()`
- Cell bounds: `cell.bounds()`
- Render: `doc.render()`
