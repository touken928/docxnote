# Python API reference

[Getting started](../README.md) · [简体中文](API_zh.md) · [Shell guide](SHELL.md)

Full interface and usage details for **docxnote**. For a minimal example, see the project [README](../README.md).

Use the Python API to parse DOCX bytes, inspect or add comments, and render
output bytes. Your application handles reading and writing files.

## Contents

- [DocxDocument](#docxdocument): parse, traverse, resolve, read comments, render
- [Paragraph](#paragraph): text and comment creation
- [Comment](#comment-object): fields, reading order, and range limits
- [Table](#table) / [Cell](#cell): grid coordinates and cell contents
- [Paths](#addressable-units-paths): addresses and object resolution
- [Advanced](#advanced): nested tables, multiple comments, merged cells
- [DocxShell / ShellResult](SHELL.md#python-api): separate command-interface reference

## DocxDocument

Represents a DOCX file.

### parse

```python
DocxDocument.parse(docx_bytes, keep_comments=False)  # -> DocxDocument
```

Parse DOCX bytes and return a document object. `keep_comments` is keyword-only.

- **keep_comments**: Whether to keep existing comments. Default `False` (strips them). Use `True` to preserve existing comments, keep their existing comment XML metadata, and append new ones.

The comments part is located through the document relationship, including relative
and package-absolute targets; its filename need not be `word/comments.xml`.
With `keep_comments=True`, the part stays at its original path, its document
relationship and its own relationships are preserved, and new comment IDs follow
the existing maximum. With `False`, the old part and its own relationship part
are removed; if new comments are added, they use the resolved part path without
reusing the old comment content or its relationships.

### blocks

```python
doc.blocks()
```

Returns block-level elements:

```python
(Paragraph | Table, ...)
```

Order matches the Word document. Block-level content controls (`w:sdt`) are
transparent: paragraphs and tables in their `w:sdtContent`, including nested
controls, appear in place. Controls do not add path segments; `p:N` and `t:N`
count the expanded blocks at that level. This also applies to `Cell.blocks()`.

### iter_paragraphs

```python
doc.iter_paragraphs()  # Iterator[Paragraph]
```

Walk paragraphs in the document body in document order, including tables,
nested tables, and block content controls, with merged cells visited once.
Each paragraph has a `path`. Wrappers are snapshotted under the document lock
when iteration begins; see [text scope](#text).

### resolve

```python
doc.resolve(path)  # Paragraph | Table | Cell | Comment
```

Locate an object by [path](#addressable-units-paths). Invalid path structure raises
`ValueError`; a valid path to a missing object raises `LookupError`. Reading an
unsupported comment range raises
[UnsupportedCommentRangeError](#single-paragraph-scope).

<a id="reading-comments-on-the-whole-document"></a>

### comments

```python
comments = doc.comments()
for c in comments:
    # c.paragraph is the owning Paragraph
    ...
```

This walks all paragraphs in the document (including inside tables and nested tables) and returns all comments in document order. Each paragraph wrapper computes its current comments when accessed, so an older wrapper observes comments added through another wrapper. Respecting `keep_comments`: with `keep_comments=False` only comments added in the current session are visible; with `keep_comments=True` existing comments from the DOCX are also exposed and preserved on render. If any paragraph holds a cross-paragraph or unclosed comment range, `doc.comments()` raises `UnsupportedCommentRangeError`.

### add_comment

```python
doc.add_comment(text, author="docxnote", date=None)  # int
```

Allocate a comment ID and store its body without creating a paragraph anchor.
Applications should normally use [`paragraph.comment(...)`](#paragraph-comment), which
creates both the comment and its anchors and returns a `Comment`.
Unanchored comments do not appear in `doc.comments()`.

### render

```python
doc.render()
```

Returns new DOCX as `bytes`. Comments are written during this step.

### Thread safety

A single `DocxDocument` instance is safe to use from multiple threads (internally serialized with a reentrant lock). Use separate instances for parallel work across threads. For multiple processes, call `parse` in each process.

All views and Shell sessions for a document share that lock. Iterators release it
before yielding wrappers to the caller.

## Paragraph

Represents a Word paragraph.

### text

```python
text = paragraph.text
```

Paragraph text, including `\n` and `\t`, with hyperlinks and inline content
controls included. Embedded text boxes inside runs belong to separate paragraphs:
their text is excluded from the host paragraph text and character offsets. Text
boxes are preserved during render but are not exposed by `blocks()`,
`iter_paragraphs()`, or the comment range views. This scope also applies to
`DocxShell`. Reading text, splitting runs, and reading/writing comment anchors
use the same host-paragraph coordinates.

<a id="paragraph-comment"></a>

### comment

```python
paragraph.comment(
    text,           # comment body
    start=0,        # start index (inclusive)
    end=None,       # end index (exclusive); None means end of paragraph
    author="docxnote",
    date=None,      # datetime (timezone-aware recommended); None = current system time
)
```

`author` and `date` are keyword-only. A paragraph with no text runs raises `ValueError`.

Adds a comment spanning the given character range in the paragraph. The range is always interpreted against `paragraph.text` using Python slice semantics (`[start, end)`), including negative and oversized endpoints. `end=None` means the paragraph end; if the normalized end precedes the normalized start, the range becomes zero-length at the normalized start. docxnote splits runs and places anchors automatically, including inside nested paragraph content such as hyperlinks. The `w:date` value in `comments.xml` is stored in UTC (`…Z`). A naive `datetime` (no `tzinfo`) is treated as UTC. Newly created comments always carry a concrete `datetime` — `date=None` means the current system time (timezone-aware), never `None`.

`paragraph.comment(...)` returns the newly created `Comment`, whose `path` can be used later with `doc.resolve(...)`.

**Example:**

```python
new = paragraph.comment("Needs change", start=3, end=8, author="Alice")
print(new.path)           # e.g. "p:0#0"
same = doc.resolve(new.path)
assert same.text == "Needs change"
```

docxnote handles run splitting, anchors, `comments.xml`, and relationship updates.

### comments

```python
paragraph.comments  # tuple[Comment, ...]
```

Return a snapshot of this paragraph's current comments. See
[Comment](#comment-object) for fields, ordering, and unsupported ranges.

<a id="comment-reading"></a>

<a id="comment-object"></a>

## Comment

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

- `paragraph` — owning `Paragraph`; use `comment.paragraph.text[start:end]` to read the annotated text.
- `path` — addressable path of this comment, e.g. `"t:0/r:0/c:0/p:0#3"`; can be passed back to `doc.resolve(...)`
- `start` / `end` — character offsets into `paragraph.text` (`[start, end)`)
- `text` / `author` — comment body and author
- `date` — `datetime | None`; parsed from the comment's `w:date`, retaining its timezone offset (or remaining naive if the source has no timezone). It is `None` when the source attribute is missing, blank, or invalid. With `keep_comments=True`, the source `w:date` attribute is preserved verbatim on render (missing stays missing, invalid strings are re-emitted unchanged).

Comments are returned in the XML document order of their `commentRangeStart` markers, so for nested ranges the outer comment comes before the inner one. This ordering also applies to equal-start and zero-length ranges, regardless of their closing order.

### Single-paragraph scope

Comment ranges are scoped to a single paragraph: `commentRangeStart` and `commentRangeEnd` must both live in the same paragraph. The high-level range views — `paragraph.comments`, `doc.comments()`, and `doc.resolve("p:0#N")` — raise `UnsupportedCommentRangeError` (a `ValueError` subclass, importable from `docxnote`) when a document contains a range that crosses paragraphs or is left unclosed. `DocxDocument.parse` and `doc.render()` are unaffected: such XML is passed through unchanged, so you can still strip it (`keep_comments=False`) or re-emit it (`keep_comments=True`) without inspecting ranges.

## Table

Represents a Word table.

### shape

```python
rows, cols = table.shape()
```

Returns `(row_count, col_count)`.

### Cell access

```python
cell = table[row, col]
```

Returns a `Cell`. All coordinates are addressable, including positions covered by merged cells.

The table uses the logical grid defined by `tblGrid`. When a row omits leading
or trailing grid coordinates with `gridBefore` or `gridAfter`, those omitted
coordinates are still addressable and return synthetic empty `Cell` objects.
Their `blocks()` result is `()`, their `bounds()` is the one-cell half-open
range `(row, col, row + 1, col + 1)`, and their `path` is the normal logical
coordinate path (for example `t:0/r:0/c:0`). That path can be passed to
`doc.resolve()` and resolves to the same synthetic cell coordinate.

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

Order matches Word, including expanded block content controls, using the same
numbering rules as `doc.blocks()`.

### bounds

```python
top, left, bottom, right = cell.bounds()
```

Cell bounds `(top, left, bottom, right)` with half-open intervals `[top, bottom)` and `[left, right)`.

For a non-merged cell, returns `(r, c, r+1, c+1)`.

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

Indices are zero-based; paragraphs and tables are counted separately. Comment
paths are resolved canonically, so harmless whitespace or separator differences
are accepted. Coordinates covered by a merged cell resolve to its origin cell,
whose `path` uses the origin coordinates. See [iter_paragraphs](#iter_paragraphs)
for traversal behavior.

<a id="path-helpers"></a>

Use an object’s `.path` and `doc.resolve(path)` to work with addresses.
Path construction and parsing helpers are internal implementation details.

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
