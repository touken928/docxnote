"""A small, streaming command interpreter over DOCX records (not a host shell)."""

from __future__ import annotations

import argparse
import json
import shlex
from collections import deque
from collections.abc import Iterable, Iterator
from dataclasses import dataclass, replace
from itertools import islice
from typing import NoReturn, TypedDict

from docxnote import Comment, DocxDocument, Paragraph, Table, Cell

DEFAULT_MAX_OUTPUT = 4096

DOCX_SHELL_INSTRUCTIONS = """Inspect and annotate the bound Word document with a restricted in-memory command language.

The tool never executes a host shell, reads files, starts processes, expands variables, or saves documents.
Use paths returned by `docx` or `comments`.

Sources:
  docx [PATH] [--start N] [--end N]
  Without PATH, output all paragraphs as JSONL records in document order.
  PATH may target a paragraph, table, or cell. Character bounds require a
  paragraph and use strict zero-based [start, end) offsets.
  Each record has `path` and `text`; sliced records also have `start`, `end`,
  and `total_chars`. Empty paragraphs are included.

Filters may follow a source with `|`:
  grep [-F] [-i] [-v] [-n] PATTERN
  head [-n N]
  tail [-n N]
Matching is literal, including without `-F`; `grep` searches serialized JSONL.
Use repeated `-e PATTERN` for OR matching and multiple grep stages for AND.
`head` and `tail` select records, not visual Word lines.

Comments:
  comment PATH TEXT [--quote QUOTE] [--start N]
  comments [PATH]
`comment` must run alone. Without `--quote`, it annotates the whole paragraph.
With `--quote`, the quote must exactly match the original paragraph text; an
ambiguous quote requires `--start`. Successful comments immediately modify the
in-memory document and return the actual comment path and range. The caller
must render and save the document.

Results contain `stdout`, `stderr`, `exit_code`, `records`, and `truncated`.
Output limits apply after filtering. A truncated first paragraph record has
`text_truncated=true`, absolute `start`/`end`, and `total_chars`; continue with
`docx PATH --start END --end N`. Do not treat truncated output as a complete
document review. Expected command errors have exit code 2 and empty stdout.
"""


class ShellResult(TypedDict):
    """Output budget applies to stdout characters, not the result envelope."""

    stdout: str
    stderr: str
    exit_code: int
    records: int
    truncated: bool


class _CommandError(ValueError):
    pass


class _Parser(argparse.ArgumentParser):
    def error(self, message: str) -> NoReturn:
        raise _CommandError(f"{self.prog}: {message}")


def _natural(value: str) -> int:
    try:
        number = int(value)
    except ValueError:
        raise argparse.ArgumentTypeError("expected a nonnegative integer") from None
    if number < 0:
        raise argparse.ArgumentTypeError("expected a nonnegative integer")
    return number


def _stages(command: str) -> list[list[str]]:
    """Split unquoted pipes, leaving quote/escape handling to shlex."""
    stages: list[list[str]] = []
    start = 0
    quote: str | None = None
    escaped = False
    for index, char in enumerate(command):
        if escaped:
            escaped = False
            continue
        if char == "\\" and quote != "'":
            escaped = True
            continue
        if quote:
            if char == quote:
                quote = None
            continue
        if char in "\"'":
            quote = char
        elif char == "|":
            stages.append(shlex.split(command[start:index]))
            start = index + 1
        elif char in ";&<>\n\r":
            raise _CommandError("only pipes are supported; quote literal punctuation")
    stages.append(shlex.split(command[start:]))
    if any(not stage for stage in stages):
        raise _CommandError("expected a command on each side of a pipe")
    return stages


def _parse(argv: list[str]) -> argparse.Namespace:
    name, *args = argv
    parser = _Parser(prog=name, add_help=False, allow_abbrev=False)
    if name in ("docx", "comments"):
        parser.add_argument("path", nargs="?")
        if name == "docx":
            parser.add_argument("--start", type=_natural)
            parser.add_argument("--end", type=_natural)
    elif name in ("head", "tail"):
        parser.add_argument("-n", type=_natural, default=10)
    elif name == "grep":
        parser.add_argument("-F", action="store_true")
        parser.add_argument("-i", action="store_true")
        parser.add_argument("-v", action="store_true")
        parser.add_argument("-n", action="store_true")
        parser.add_argument("-e", action="append", default=[])
        parser.add_argument("pattern", nargs="?")
    elif name == "comment":
        parser.add_argument("path")
        parser.add_argument("text")
        parser.add_argument("--quote")
        parser.add_argument("--start", type=_natural)
    else:
        raise _CommandError(f"unsupported command: {name}")
    options = parser.parse_args(args)
    options.command = name
    if name == "grep":
        if options.pattern is not None and options.e:
            raise _CommandError("grep: use a pattern or repeated -e patterns, not both")
        if options.pattern is None and not options.e:
            raise _CommandError("grep: missing pattern")
        options.patterns = options.e or [options.pattern]
    if name == "comment":
        if not options.text.strip():
            raise _CommandError("comment: text must not be blank")
        if options.quote == "":
            raise _CommandError("comment: quote must not be empty")
        if options.start is not None and options.quote is None:
            raise _CommandError("comment: --start requires --quote")
    return options


def _json(data: dict) -> str:
    return json.dumps(data, ensure_ascii=False, separators=(",", ":"))


@dataclass(frozen=True)
class _Line:
    data: dict
    prefix: str = ""

    def render(self) -> str:
        return self.prefix + _json(self.data)


def _grep(lines: Iterable[_Line], options: argparse.Namespace) -> Iterator[_Line]:
    patterns = [p.casefold() if options.i else p for p in options.patterns]
    for number, line in enumerate(lines, 1):
        text = line.render()
        if options.i:
            text = text.casefold()
        matched = any(pattern in text for pattern in patterns)
        if matched != options.v:
            yield replace(line, prefix=f"{number}:{line.prefix}") if options.n else line


def _comment_data(comment: Comment) -> dict:
    return {
        "path": comment.path,
        "target": comment.paragraph.path,
        "start": comment.start,
        "end": comment.end,
        "text": comment.text,
        "author": comment.author,
        "date": comment.date.isoformat() if comment.date is not None else None,
    }


def _preview(line: _Line, limit: int) -> str | None:
    """Fit the first oversized record without ever cutting serialized JSON."""
    data = line.data.copy()
    text = data["text"]
    data["text_truncated"] = True
    is_comment = "target" in data
    if is_comment:
        data["text_length"] = len(text)
    else:
        data.setdefault("start", 0)
        data.setdefault("total_chars", len(text))

    def render(length: int) -> str:
        data["text"] = text[:length]
        if not is_comment:
            data["end"] = data["start"] + length
        return line.prefix + _json(data)

    if len(render(0)) > limit:
        return None
    low, high = 0, len(text)
    while low < high:
        middle = (low + high + 1) // 2
        if len(render(middle)) <= limit:
            low = middle
        else:
            high = middle - 1
    return render(low)


class DocxShell:
    """Bind six in-memory commands to a document. No Pydantic AI dependency.

    max_output limits stdout characters after the pipeline, with complete JSON
    records and a marked preview when its first record exceeds the budget.
    """

    def __init__(
        self,
        doc: DocxDocument,
        *,
        author: str = "docxnote",
        max_output: int = DEFAULT_MAX_OUTPUT,
    ) -> None:
        if (
            isinstance(max_output, bool)
            or not isinstance(max_output, int)
            or max_output < 256
        ):
            raise ValueError("max_output must be an integer of at least 256 characters")
        self._doc = doc
        self._author = author
        self._max_output = max_output
        self._added_comments: list[Comment] = []

    @property
    def added_comments(self) -> tuple[Comment, ...]:
        """Snapshot of comments successfully added through this session."""
        with self._doc._lock:
            return tuple(self._added_comments)

    def run(self, command: str) -> ShellResult:
        """Execute the command language described by DOCX_SHELL_INSTRUCTIONS."""
        with self._doc._lock:
            try:
                try:
                    argv_stages = _stages(command)
                except ValueError as exc:
                    raise _CommandError(str(exc)) from exc
                stages = [_parse(argv) for argv in argv_stages]
                first, *filters = stages
                if first.command not in ("docx", "comments", "comment"):
                    raise _CommandError("pipeline must start with docx or comments")
                if first.command == "comment" and filters:
                    raise _CommandError("comment must be a standalone command")
                if any(
                    stage.command not in ("grep", "head", "tail") for stage in filters
                ):
                    raise _CommandError("only grep, head and tail may follow a pipe")
                if first.command == "comment":
                    lines: Iterable[_Line] = [self._comment(first)]
                else:
                    lines = self._source(first)
                for stage in filters:
                    if stage.command == "grep":
                        lines = _grep(lines, stage)
                    elif stage.command == "head":
                        lines = islice(lines, stage.n)
                    else:
                        lines = deque(lines, maxlen=stage.n) if stage.n else ()
                return self._collect(lines)
            except _CommandError as exc:
                return {
                    "stdout": "",
                    "stderr": str(exc),
                    "exit_code": 2,
                    "records": 0,
                    "truncated": False,
                }

    def _resolve(self, path: str):
        try:
            return self._doc.resolve(path)
        except (ValueError, LookupError) as exc:
            raise _CommandError(str(exc)) from exc

    def _paragraphs(self, path: str | None) -> Iterable[Paragraph]:
        if path is None:
            return self._doc.iter_paragraphs()
        obj = self._resolve(path)
        if isinstance(obj, Paragraph):
            return (obj,)
        if not isinstance(obj, (Table, Cell)):
            raise _CommandError("expected a paragraph, table or cell path")
        prefix = obj.path + "/"
        return (p for p in self._doc.iter_paragraphs() if p.path.startswith(prefix))

    def _source(self, options: argparse.Namespace) -> Iterable[_Line]:
        paragraphs = self._paragraphs(options.path)
        if options.command == "comments":
            return self._comments(paragraphs)
        sliced = options.start is not None or options.end is not None
        if sliced:
            if options.path is None or not isinstance(
                self._resolve(options.path), Paragraph
            ):
                raise _CommandError("docx: character bounds require a paragraph path")
            paragraph = next(iter(paragraphs))
            text = paragraph.text
            start = options.start if options.start is not None else 0
            end = options.end if options.end is not None else len(text)
            if not 0 <= start <= end <= len(text):
                raise _CommandError(f"docx: expected 0 <= start <= end <= {len(text)}")
            return [
                _Line(
                    {
                        "path": paragraph.path,
                        "text": text[start:end],
                        "start": start,
                        "end": end,
                        "total_chars": len(text),
                    }
                )
            ]
        return (_Line({"path": p.path, "text": p.text}) for p in paragraphs)

    @staticmethod
    def _comments(paragraphs: Iterable[Paragraph]) -> Iterator[_Line]:
        from docxnote import UnsupportedCommentRangeError

        for paragraph in paragraphs:
            try:
                comments = paragraph.comments
            except UnsupportedCommentRangeError as exc:
                raise _CommandError(str(exc)) from exc
            for comment in comments:
                yield _Line(_comment_data(comment))

    def _comment(self, options: argparse.Namespace) -> _Line:
        paragraph = self._resolve(options.path)
        if not isinstance(paragraph, Paragraph):
            raise _CommandError("comment: expected a paragraph path")
        start, end = 0, len(paragraph.text)
        if options.quote is not None:
            quote = options.quote
            if options.start is None:
                start = paragraph.text.find(quote)
                if start < 0:
                    raise _CommandError(
                        "comment: quote not found; read the original text again"
                    )
                if paragraph.text.find(quote, start + 1) >= 0:
                    raise _CommandError("comment: ambiguous quote; supply --start")
            else:
                start = options.start
            end = start + len(quote)
            if paragraph.text[start:end] != quote:
                raise _CommandError("comment: quote does not match at --start")
        try:
            comment = paragraph.comment(
                options.text, start=start, end=end, author=self._author
            )
        except ValueError as exc:
            raise _CommandError(str(exc)) from exc
        self._added_comments.append(comment)
        return _Line(_comment_data(comment))

    def _collect(self, lines: Iterable[_Line]) -> ShellResult:
        output: list[str] = []
        length = 0
        truncated = False
        for line in lines:
            rendered = line.render()
            required = len(rendered) + bool(output)
            if length + required > self._max_output:
                truncated = True
                if not output:
                    preview = _preview(line, self._max_output)
                    if preview is not None:
                        output.append(preview)
                break
            output.append(rendered)
            length += required
        return {
            "stdout": "\n".join(output),
            "stderr": "",
            "exit_code": 0,
            "records": len(output),
            "truncated": truncated,
        }


# Keep the callable's framework-visible documentation and the exported
# instructions byte-for-byte identical, so tool adapters cannot drift.
DocxShell.run.__doc__ = DOCX_SHELL_INSTRUCTIONS
