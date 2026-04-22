# `docxnote` CLI reference

Installing the package registers a `docxnote` console script powered by
[`src/docxnote/cli.py`](../src/docxnote/cli.py). Every read command can emit
JSON (`--json`) so it plugs directly into LLM pipelines, shell tools and
spreadsheets; every write command requires an explicit output file.

```
docxnote --help
docxnote <command> --help
```

All commands operate on **addressable paths** — the same string identifiers
returned by `paragraph.path`, `Comment.path`, `Table.path`, `Cell.path` and
accepted by `DocxDocument.resolve(path)`. See [API.md](API.md#addressable-units-paths)
for the path grammar.

Run it inside a checkout without installing:

```
uv run docxnote <command> ...
# or
python -m docxnote.cli <command> ...
```

---

## `docxnote list FILE`

Enumerate every paragraph in the document (including those inside tables and
nested tables), de-duplicating merged cells.

Options:

| flag | default | description |
|------|---------|-------------|
| `--text` | off | include paragraph text after the path (plain mode only) |
| `--text-limit N` | `80` | truncate text preview to N chars; `0` disables truncation |
| `--json` | off | emit a JSON array of `{path, text}` |
| `--keep-comments` / `--no-keep-comments` | strip | preserve existing Word comments while parsing |

Example:

```
$ docxnote list report.docx --text --text-limit 40
p:0     Quarterly Report
p:1     Revenue grew by 12% quarter over quar…
t:0/r:0/c:0/p:0 Region
t:0/r:0/c:1/p:0 Revenue
```

JSON mode is stable and UTF-8 safe:

```
$ docxnote list report.docx --json | jq '.[0]'
{
  "path": "p:0",
  "text": "Quarterly Report"
}
```

---

## `docxnote show FILE PATH`

Resolve `PATH` to a `Paragraph`, `Table`, `Cell` or `Comment` and print it.
Default `--keep-comments` is **on** because `show` is a read operation.

JSON schemas:

```jsonc
// paragraph
{"type": "paragraph", "path": "p:0", "text": "...", "comments": [ /* see below */ ]}

// table
{"type": "table", "path": "t:0", "shape": [rows, cols]}

// cell
{"type": "cell", "path": "t:0/r:1/c:2",
 "bounds": [top, left, bottom, right],
 "blocks": [{"kind": "paragraph", "path": "t:0/r:1/c:2/p:0"}, ...]}

// comment
{"type": "comment", "path": "p:1#3", "paragraph": "p:1",
 "start": 0, "end": 5, "text": "...", "author": "...", "date": "ISO-8601"}
```

Plain output is human-readable and suitable for piping into less/grep.

---

## `docxnote comments FILE`

List every Word comment in the document in document order. Default
`--keep-comments` is **on**.

Plain rows: `<comment_path>\t<start>:<end>\t<author>\t<preview>`.

```
$ docxnote comments report.docx --json | jq '.[0]'
{
  "path": "p:1#0",
  "paragraph": "p:1",
  "start": 0,
  "end": 7,
  "text": "Please cite source.",
  "author": "reviewer",
  "date": "2026-04-22T09:12:00+00:00"
}
```

---

## `docxnote annotate INPUT OUTPUT`

Add one or more comments to `INPUT` and write the modified document to
`OUTPUT`. Never writes in place.

### Single-shot

```
docxnote annotate in.docx out.docx \
  --path "p:1" \
  --text "Please revise this opening." \
  --start 0 --end 5 \
  --author reviewer
```

`--path` and `--text` must be provided together. `--start` defaults to `0`
and `--end` defaults to *end of paragraph*.

### Batch via JSON spec

`--spec ops.json` reads an array of operation objects. The single-shot form
may be combined with `--spec` and is appended last.

```json
[
  {"path": "p:0", "text": "title needs a year", "author": "editor"},
  {"path": "t:0/r:1/c:2/p:0", "text": "check number", "start": 0, "end": 3}
]
```

### Behavior

- `--keep-comments` / `--no-keep-comments` controls whether the *input*
  document's existing comments are preserved before adding new ones;
  default is to **strip** (matches the Python `DocxDocument.parse` default).
- If any op's `path` does not resolve to a `Paragraph`, the command exits
  with status `2` and **no file is written**.
- With `--json`, stdout is `{"output": "<path>", "added": [<comment dict>, ...]}`.

---

## Exit codes

| code | meaning |
|------|---------|
| `0` | success |
| `2` | user error: bad arguments, missing file, unresolved path, invalid spec |

---

## Patterns

**Locate → annotate loop** (shell):

```sh
# 1. inspect paragraphs
docxnote list input.docx --text --json > index.json

# 2. choose targets and build ops.json (any tool: jq, LLM, script)
jq '[.[] | select(.text | test("TODO")) | {path, text: "resolve TODO"}]' \
   index.json > ops.json

# 3. apply
docxnote annotate input.docx annotated.docx --spec ops.json --json
```

**Round-trip sanity check**:

```
docxnote annotate in.docx tmp.docx --path p:0 --text "hi"
docxnote comments tmp.docx --json
```
