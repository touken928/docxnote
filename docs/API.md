# Python API reference

Full interface and usage details for **docxnote**. For a minimal example, see the project [README](../README.md).

---

## DocxDocument

Represents a DOCX file.

### parse

```python
DocxDocument.parse(docx_bytes, *, keep_comments=False)
```

Parses the DOCX and returns a document object.

- **keep_comments**: Whether to keep existing comments. Default `False` (strips them). Use `True` to preserve existing comments, keep their existing comment XML metadata, and append new ones.

---

### blocks

```python
doc.blocks()
```

Returns block-level elements:

```python
(Paragraph | Table, ...)
```

Order matches the Word document.

---

### render

```python
doc.render()
```

Returns new DOCX as `bytes`. Comments are written during this step.

### Thread safety

A single `DocxDocument` instance is safe to use from multiple threads (internally serialized with a reentrant lock). Use separate instances for parallel work across threads. For multiple processes, call `parse` in each process.

---

## Addressable units (paths)

Every `Paragraph`, `Table`, `Cell` and `Comment` has a stable string **path** that identifies its position in the document:

- `p:N` — Nth paragraph at this level
- `t:N` — Nth table at this level
- `t:N/r:R/c:C` — cell at row `R`, column `C` (based on its merge origin)
- `t:N/r:R/c:C/p:M` — paragraph inside a cell (recurses for nested tables)
- `<paragraph_path>#<id>` — a specific comment on a paragraph (where `<id>` is Word's internal `w:id`)

```python
for block in doc.blocks():
    print(block.path, type(block).__name__)

# Round-trip: path → object
para = doc.resolve("p:0")
cell = doc.resolve("t:0/r:1/c:2")
comment = doc.resolve("t:0/r:1/c:2/p:0#3")
```

`doc.resolve(path)` returns a `Paragraph` / `Table` / `Cell` / `Comment`. Use `doc.iter_paragraphs()` to walk every paragraph (including inside tables and nested tables, de-duplicating merged cells) — each yielded paragraph carries its own `path`.

---

## Paragraph

Represents a Word paragraph.

---

## Comment reading

```python
from docxnote import Comment
```

Each comment attached to a paragraph is exposed as a `Comment` object via `paragraph.comments`:

```python
for c in paragraph.comments:
    assert isinstance(c, Comment)
    print(c.path, c.start, c.end, c.text, c.author)
```

Fields:

- `path` — addressable path of this comment, e.g. `"t:0/r:0/c:0/p:0#3"`; can be passed back to `doc.resolve(...)`
- `start` / `end` — character offsets into `paragraph.text` (`[start, end)`)
- `text` / `author` / `date` — comment body, author, and `w:date` (UTC)

### text

```python
text = paragraph.text
```

Full paragraph text, including `\n` and `\t`.

---

### comment

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

Adds a comment spanning the given character range in the paragraph. The range is always interpreted against `paragraph.text` using Python slice semantics (`[start, end)`). docxnote splits runs and places anchors automatically, including inside nested paragraph content such as hyperlinks. The `w:date` value in `comments.xml` is stored in UTC (`…Z`). A naive `datetime` (no `tzinfo`) is treated as UTC.

`paragraph.comment(...)` returns the newly created `Comment`, whose `path` can be used later with `doc.resolve(...)`.

**Example:**

```python
new = paragraph.comment("Needs change", start=3, end=8, author="Alice")
print(new.path)           # e.g. "p:0#0"
same = doc.resolve(new.path)
assert same.text == "Needs change"
```

docxnote handles run splitting, anchors, `comments.xml`, and relationship updates.

---

## Reading comments on the whole document

```python
comments = doc.comments()
for c in comments:
    # c.paragraph is the owning Paragraph
    ...
```

This walks all paragraphs in the document (including inside tables and nested tables) and returns all comments in document order. Respecting `keep_comments`: with `keep_comments=False` only comments added in the current session are visible; with `keep_comments=True` existing comments from the DOCX are also exposed and preserved on render.

---

## Table

Represents a Word table.

### shape

```python
rows, cols = table.shape()
```

Returns `(row_count, col_count)`.

---

### Cell access

```python
cell = table[row, col]
```

Returns a `Cell`. All coordinates are addressable, including positions covered by merged cells.

---

## Cell

Represents a table cell.

### blocks

```python
cell.blocks()
```

Block-level elements inside the cell:

```python
(Paragraph | Table, ...)
```

Order matches Word.

---

### bounds

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
