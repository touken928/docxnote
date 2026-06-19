---
name: docxnote
description: >
  Lightweight Python library and `docxnote` CLI for DOCX Word comments: parse
  ZIP/XML, traverse paragraphs and tables through plain-text views, add and read
  comments by character range, handle merged and nested tables, and resolve
  stable address paths. Use when the user mentions docxnote, DOCX review
  automation, or Word comments. Full Python API: docs/API.md. Full CLI
  reference: docs/CLI.md. Agent-oriented usage: library-usage.md and
  cli-usage.md.
---

# docxnote

`docxnote` is designed for paragraph-text-driven manipulation of Word `.docx`
 files. It reads document text, adds or reads comments, and writes the required
 `comments.xml` and relationship updates during render. It keeps the public API
 at the paragraph / table / cell / comment level rather than exposing raw runs
 or OOXML details.

## Core capabilities

- Parse and render DOCX bytes with `DocxDocument.parse(...)` and `doc.render()`.
- Traverse top-level blocks with `doc.blocks()` and recursively traverse table
  content with `cell.blocks()`.
- Work with `paragraph.text` as a plain-text view, including `\n` and `\t`.
- Add comments with `paragraph.comment(...)` and read them with
  `paragraph.comments` or `doc.comments()`.
- Preserve existing comments with `keep_comments=True`, or strip them on parse
  with the default `keep_comments=False`.
- Use stable paths such as `paragraph.path` and `comment.path`, then resolve
  them later with `doc.resolve(path)`.
- Use the `docxnote` CLI (`list`, `show`, `comments`, `annotate`) on the same
  path system as the library.
- Share one `DocxDocument` safely across threads through an internal re-entrant
  lock.

## Documentation map

| File | Purpose |
|------|---------|
| `docs/API.md` | Full Python API reference. |
| `docs/CLI.md` | Full CLI reference. |
| [`library-usage.md`](library-usage.md) | Agent-facing Python workflows, patterns, and examples. |
| [`cli-usage.md`](cli-usage.md) | Agent-facing CLI workflows, pitfalls, and examples. |
| `SKILL.md` | Short capability summary for fast agent loading. |

## When to use docxnote

Use it for automated review flows, rule-based comment insertion, comments inside
tables, and workflows that append machine comments on top of existing human
comments.

Avoid dropping to low-level OOXML editing unless absolutely necessary; prefer
the high-level `DocxDocument` / `Paragraph` / `Table` / `Cell` / `Comment` API.
