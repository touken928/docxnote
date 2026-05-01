# docxnote CLI usage (for agents)

This document is the **Agent-oriented** guide for driving `docxnote` from a
shell. The authoritative reference is `docs/CLI.md`; this
file adds patterns, pitfalls and LLM-friendly workflows.

---

## When to use the CLI

- You want to operate on a `.docx` without writing Python.
- You are orchestrating from a shell / Make / CI pipeline.
- You are an LLM that emits shell commands or JSON tool calls.

For anything that needs fine-grained logic (per-run reasoning, custom
filtering, XML-level changes) use the Python library — see
[`library-usage.md`](library-usage.md).

---

## Mental model (one screen)

`docxnote` exposes 4 subcommands built on the same path system as the
library:

| command | read/write | purpose |
|---------|------------|---------|
| `list`     | read  | enumerate every paragraph's `path` (+ preview) |
| `show`     | read  | dump a paragraph / table / cell / comment by `path` |
| `comments` | read  | dump every comment in the document |
| `annotate` | write | add one or many comments, emit a **new** `.docx` |

Every read command supports `--json`. Every object the library exposes has a
stable `path` string; the CLI speaks exclusively in those strings.

Path primer (full grammar in `docs/API.md`):

```
p:N                        # N-th top-level paragraph
t:N                        # N-th top-level table
t:N/r:R/c:C                # cell (R,C) inside table N (top-left of merge)
t:N/r:R/c:C/p:M            # paragraph inside that cell
<paragraph-path>#<id>      # a comment anchored at that paragraph
```

---

## Standard agent workflow

The safe pattern is **locate → annotate → verify**. Each step is a single
CLI call with JSON output.

### 1. Locate

```sh
docxnote list input.docx --text --text-limit 120 --json > index.json
```

`index.json` is a JSON array of `{"path": "...", "text": "..."}` covering
every paragraph (including nested / merged cells, de-duplicated).

### 2. Decide

Turn `index.json` into an ops file. Any tool works; below is LLM-style
pseudo-logic:

```python
ops = []
for item in index:
    if "TODO" in item["text"]:
        ops.append({"path": item["path"], "text": "resolve this TODO",
                    "author": "reviewer"})
```

Persist as `ops.json`:

```json
[
  {"path": "p:1", "text": "resolve this TODO", "author": "reviewer"},
  {"path": "t:0/r:2/c:1/p:0", "text": "check number", "start": 0, "end": 3}
]
```

### 3. Apply

```sh
docxnote annotate input.docx annotated.docx --spec ops.json --json
```

Non-zero exit → nothing was written. Capture stderr.

### 4. Verify

```sh
docxnote comments annotated.docx --json
```

`path` values inside the result echo the paragraphs you targeted; use that
as your acceptance check.

---

## Anchoring a comment precisely

`start`/`end` are **character offsets** into `paragraph.text` (which joins
runs and inserts `\n`/`\t`). Typical strategy:

1. `docxnote show input.docx <para-path> --json` to fetch full text.
2. Compute `start = text.index(phrase)` and `end = start + len(phrase)`.
3. Put them in the op.

If you only need to comment the whole paragraph, omit `start`/`end` — they
default to `0` and *end of paragraph*.

---

## Preserving existing comments

- `list` / `show` / `comments` default to **`--keep-comments`** (they're reads).
- `annotate` defaults to **`--no-keep-comments`** (the library default). Add
  `--keep-comments` to stack new comments on top of existing ones.

Rule of thumb: if the user says "append", pass `--keep-comments` to
`annotate` and to the post-check `comments`/`show` call.

---

## Common pitfalls

- **Targeting a table, row or non-paragraph with `annotate`** → exits `2`.
  Only paragraph paths are writable (`p:N`, `t:.../p:M`).
- **Editing in place** isn't supported; pass two distinct files. If the
  caller wants in-place, they must overwrite `INPUT` themselves after a
  successful run.
- **Mixing `--path/--text` with `--spec`** is allowed; the single op is
  *appended* after spec entries.
- **Non-UTF-8 spec files**: the CLI reads `--spec` as UTF-8.
- **Shell quoting** with Chinese or emoji content: on Windows PowerShell,
  prefer `--spec` over long inline `--text` arguments.

---

## Cheat sheet

```sh
# what's in this doc?
docxnote list input.docx --text --json

# pull one comment’s full text
docxnote show input.docx "p:2#0" --json

# single comment
docxnote annotate in.docx out.docx \
  --path p:2 --text "please cite source" --author editor

# bulk
docxnote annotate in.docx out.docx --spec ops.json --json

# append instead of overwrite
docxnote annotate in.docx out.docx --spec ops.json --keep-comments

# inspect result
docxnote comments out.docx --json
```

---

## Related

- `docs/CLI.md` — exhaustive CLI reference.
- [`library-usage.md`](library-usage.md) — Python library patterns.
- `docs/API.md` — full Python API including path grammar.
