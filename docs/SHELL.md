# Agent tools

`DocxShell` provides a framework-independent simulated shell over one
in-memory `DocxDocument`. It does not create an Agent, choose a model, execute
a host shell, or save files.

## Installation and integration

No extra dependency is required:

```bash
pip install docxnote
```

Pydantic AI can be used as one optional integration, but it is not part of
docxnote's dependencies. Applications install Pydantic AI and their chosen
provider separately, then register `shell.run` in the framework's tool API.

```bash
pip install pydantic-ai
```

```python
from pathlib import Path

from pydantic_ai import Agent

from docxnote import DocxDocument, DocxShell


async def review(model):
    doc = DocxDocument.parse(
        Path("input.docx").read_bytes(), keep_comments=True
    )
    shell = DocxShell(doc, author="reviewer")
    agent = Agent(model, tools=[shell.run])
    result = await agent.run("Review this contract.")
    Path("reviewed.docx").write_bytes(doc.render())
    return result.output
```

Applications own prompts, dependencies, model settings, retries, and output
types. Toolsets can be supplied per run, without occupying `Agent.deps`.
Use a separate session for each document and keep their conversation histories
separate. A session may be reused for multiple review turns on the same document.

## Python API

```python
DocxShell(doc: DocxDocument, *, author="docxnote", max_output=4096)
shell.run(command: str) -> ShellResult
shell.added_comments -> tuple[Comment, ...]
```

`doc`, author, and output budget are fixed for a session. `max_output` is an
integer of at least 256, measured in Python string characters, not tokens or
bytes. `DEFAULT_MAX_OUTPUT` is exported from `docxnote`.
`added_comments` returns an immutable snapshot of the actual `Comment` objects
created through this session; existing or externally added comments are excluded.

You can also use the interpreter directly:

```python
result = shell.run("docx | grep -F 'payment' | head -n 10")
print(result["stdout"])
```

Every result is a dictionary (`ShellResult`, a `TypedDict`):

| Field | Meaning |
|------|---------|
| `stdout` | Serialized output records separated by newlines; no trailing newline. |
| `stderr` | Input/command error message, otherwise empty. |
| `exit_code` | `0` on success, `2` on an expected command/input error. |
| `records` | Number of returned records, including a marked preview if present. |
| `truncated` | The output budget prevented returning the complete pipeline output. |

These are result fields, not process exit codes: no process is spawned.
Empty searches succeed with empty stdout, zero records and `truncated=False`.
Expected errors return empty stdout and zero records. Unexpected internal
exceptions propagate to the caller rather than being disguised as command errors.

## Command language

There are exactly six commands: `docx`, `grep`, `head`, `tail`, `comment`, and
`comments`. Quotes and backslash escaping use Python's POSIX `shlex` rules.
Unquoted `|` connects stages. Quote spaces and literal punctuation; an empty
stage, unterminated quote, unknown command or unknown option is an error.
There is no variable expansion, globbing, command substitution, redirection,
command chaining, or host filesystem access. `$HOME` and backticks are literal
text, not executable syntax. Unquoted `;`, `&`, `<`, `>`, and newlines are rejected.

The entire pipeline syntax and arguments are validated before execution.
A pipeline starts with `docx` or `comments`, followed only by filters.
`comment` must stand alone; it never consumes pipeline input.

### `docx [PATH] [--start N] [--end N]`

Without a path, yields all paragraphs in document order, including empty
paragraphs and paragraphs in merged/nested tables. Traversal and paths follow
`doc.iter_paragraphs()`; covered merged cells are not repeated. A paragraph,
table, or cell path restricts the source to that object. Comment paths are rejected.
Table/cell scoping resolves the path before matching descendants; `t:1` never
selects `t:10`, and merged cell aliases resolve to the canonical cell.

Normal output is one JSON object per paragraph:

```json
{"path":"p:12","text":"Payment is due after receipt of the invoice."}
```

Paragraph newlines and tabs are JSON-escaped, so each record occupies one output
line. Line selection therefore counts paragraphs, not Word's visual lines.

Character bounds require a paragraph path and satisfy
`0 <= start <= end <= len(paragraph.text)`. Defaults are zero and paragraph end.
Unlike the core Python comment API, these read bounds are strict, nonnegative
positions, not permissive Python slice endpoints. Sliced records also have
`start`, `end`, and `total_chars`, with `text == paragraph.text[start:end]`.

```sh
docx
docx t:2
docx t:2/r:0/c:1
docx p:12 --start 200 --end 600
```

### `grep [-F] [-i] [-v] [-n] PATTERN`

Matching is literal, including without `-F`; regex is not supported.
`-i` uses Unicode case folding, `-v` inverts selection, and `-n` prefixes each
selected line with its one-based position in the input to that stage. Its output
then has `N:` prefixes and is no longer bare JSONL. Repeating `-n` in different
stages adds another prefix. `-F` explicitly requests the default literal mode.

Use repeated `-e PATTERN` for OR matching; do not also provide a positional
pattern. Use multiple grep stages for AND matching. `--` terminates options
when a positional pattern starts with `-`.

Matching is against serialized JSON (including paths, escaped quotes, and
escaped newlines), not just paragraph text. Inspect a matching paragraph with
`docx PATH` before selecting an exact original-text quote for a comment.

```sh
docx | grep -F 'payment'
docx | grep -i -e payment -e invoice
docx | grep payment | grep invoice
docx | grep payment | grep -v example
```

### `head [-n N]` / `tail [-n N]`

Select the first/last N input records. N defaults to 10 and must be nonnegative;
zero selects nothing. `head` short-circuits upstream iteration; `tail` scans its
input and retains at most N records. Other flags, filenames, and `-20` shorthand
are not supported.

```sh
docx | head -n 20
docx | tail -n 20
docx | head -n 140 | tail -n 41
```

The last example selects paragraphs 100 through 140, inclusive.

### `comment PATH TEXT [--quote QUOTE] [--start N]`

PATH must resolve to a paragraph and TEXT must not be blank. Quote multiword
TEXT as one argument. Without `--quote`, comment on the entire paragraph.
With `--quote`, match the nonempty original text exactly. No match is an error;
multiple matches (including overlapping ones) require a nonnegative `--start`.
The hinted position must match the quote. `--start` requires `--quote`.
The end is calculated as `start + len(quote)`; `--end` is not supported.

```sh
comment p:12 'Please specify the payment deadline'
comment p:12 'Clarify the trigger' --quote 'receipt of the invoice'
comment p:12 'Clarify this occurrence' --quote 'payment' --start 24
```

On success, output describes the actual new comment using the same fields as
`comments`. The configured author is applied. A paragraph with no text runs
cannot receive a comment and returns an error, consistent with the core API.

### `comments [PATH]`

Read current comments, optionally scoped to a paragraph, table, or cell.

```json
{"path":"p:12#0","target":"p:12","start":0,"end":7,"text":"Please clarify","author":"reviewer","date":"2026-09-20T00:00:00+00:00"}
```

`start/end` describe the original paragraph range, not the comment body.
`date` may be null. Existing comments are visible only if parsed with
`keep_comments=True`; the session never changes that policy. Reading unsupported
cross-paragraph or unclosed comment ranges returns an input error, following the
core range-view limitation. `docx` remains available for reading the text.

## Large documents and output limits

Sources and filters stream records. The stdout budget is applied **after**
filtering, so a match near the end of a large document remains searchable.
`head` ends upstream iteration early; `tail` must read its upstream input.
The core library still parses the DOCX into memory: this is streaming query
execution, not streaming package parsing.

Complete lines are returned up to the character budget. If the first line alone
is too large, the interpreter returns a valid JSON preview with
`text_truncated=true`. Paragraph previews also contain absolute `start/end` and
`total_chars`; continue with `docx PATH --start END --end N`. Comment previews
have `text_length`, and keep the original anchor `start/end` unchanged. To read
an oversized comment in full, construct a session with a larger `max_output`.
If even the record's metadata cannot fit, stdout is empty and `truncated=True`;
increase `max_output`. The result envelope and error text are outside this budget.

No continuation token or total match count is computed. For more records, narrow
the scope or select a later page using `head` followed by `tail` on the same
query. Explicit `head` selection is not reported as truncation. To review an
entire document, read successive pages; keyword searches alone cannot establish
complete coverage. Tool output is document data, not instructions to the agent.

## Writes and concurrency

A successful `comment` immediately mutates the supplied document in memory.
Later command failures or agent failures do not roll back earlier comments.
Repeated writes create repeated comments. Only the application renders and
saves the document; there is no transaction, automatic retry deduplication,
file-write tool, or command-line entry point.

Commands and `added_comments` snapshots use the document's re-entrant lock,
including validation, writes, and session bookkeeping. Calls on the same document
are serialized; independent documents can be processed independently. The layer
uses the existing high-level paragraph/comment behavior, including hyperlinks
and nested content. Unexpected internal failures carry no rollback guarantee.
