"""docxnote 命令行工具

提供 4 个子命令：``list`` / ``show`` / ``comments`` / ``annotate``。
所有读命令都把对象的 ``path`` 打印出来，便于下游（LLM / 脚本）用它回头
精确定位；写命令显式要求输出文件路径，避免原地覆盖造成意外。

典型用法::

    docxnote list input.docx --text --json
    docxnote show input.docx "t:0/r:1/c:2/p:0"
    docxnote comments input.docx --json
    docxnote annotate input.docx output.docx \\
        --path "p:0" --text "Needs revision" --start 0 --end 5 --author reviewer
    docxnote annotate input.docx output.docx --spec ops.json
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Sequence

from . import Cell, Comment, DocxDocument, Paragraph, Table


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _read_doc(path: Path, *, keep_comments: bool) -> DocxDocument:
    return DocxDocument.parse(Path(path).read_bytes(), keep_comments=keep_comments)


def _truncate(s: str, limit: int | None) -> str:
    if limit is None or limit <= 0 or len(s) <= limit:
        return s
    return s[: max(0, limit)] + "…"


def _paragraph_to_dict(p: Paragraph, *, text_limit: int | None) -> dict[str, Any]:
    return {"path": p.path, "text": _truncate(p.text, text_limit)}


def _comment_to_dict(c: Comment) -> dict[str, Any]:
    return {
        "path": c.path,
        "paragraph": c.paragraph.path,
        "start": c.start,
        "end": c.end,
        "text": c.text,
        "author": c.author,
        "date": c.date.isoformat(),
    }


def _dump_json(payload: Any) -> None:
    json.dump(payload, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")


# ---------------------------------------------------------------------------
# subcommands
# ---------------------------------------------------------------------------


def cmd_list(args: argparse.Namespace) -> int:
    doc = _read_doc(args.file, keep_comments=args.keep_comments)
    paras = list(doc.iter_paragraphs())
    limit = args.text_limit if args.text_limit > 0 else None

    if args.json:
        _dump_json([_paragraph_to_dict(p, text_limit=limit) for p in paras])
    else:
        for p in paras:
            if args.text:
                text = (
                    _truncate(p.text, limit).replace("\n", "\\n").replace("\t", "\\t")
                )
                print(f"{p.path}\t{text}")
            else:
                print(p.path)
    return 0


def cmd_show(args: argparse.Namespace) -> int:
    doc = _read_doc(args.file, keep_comments=args.keep_comments)
    obj = doc.resolve(args.path)
    payload = _describe(obj)

    if args.json:
        _dump_json(payload)
    else:
        _print_show_plain(payload)
    return 0


def _describe(obj: Paragraph | Table | Cell | Comment) -> dict[str, Any]:
    if isinstance(obj, Paragraph):
        return {
            "type": "paragraph",
            "path": obj.path,
            "text": obj.text,
            "comments": [_comment_to_dict(c) for c in obj.comments],
        }
    if isinstance(obj, Table):
        rows, cols = obj.shape()
        return {"type": "table", "path": obj.path, "shape": [rows, cols]}
    if isinstance(obj, Cell):
        top, left, bottom, right = obj.bounds()
        return {
            "type": "cell",
            "path": obj.path,
            "bounds": [top, left, bottom, right],
            "blocks": [
                {
                    "kind": "paragraph" if isinstance(b, Paragraph) else "table",
                    "path": b.path,
                }
                for b in obj.blocks()
            ],
        }
    if isinstance(obj, Comment):
        return {"type": "comment", **_comment_to_dict(obj)}
    raise ValueError(f"unsupported object type: {type(obj).__name__}")


def _print_show_plain(payload: dict[str, Any]) -> None:
    kind = payload["type"]
    if kind == "paragraph":
        print(f"path: {payload['path']}")
        print("text:")
        print(payload["text"])
        if payload["comments"]:
            print("comments:")
            for c in payload["comments"]:
                print(
                    f"  - {c['path']}  [{c['start']}:{c['end']}]  "
                    f"{c['author']}  {c['date']}"
                )
                for line in (c["text"] or "").splitlines() or [""]:
                    print(f"    {line}")
    elif kind == "table":
        rows, cols = payload["shape"]
        print(f"path: {payload['path']}")
        print(f"shape: {rows} x {cols}")
    elif kind == "cell":
        print(f"path: {payload['path']}")
        top, left, bottom, right = payload["bounds"]
        print(f"bounds: top={top} left={left} bottom={bottom} right={right}")
        for b in payload["blocks"]:
            print(f"  {b['kind']}\t{b['path']}")
    elif kind == "comment":
        print(f"path: {payload['path']}")
        print(f"paragraph: {payload['paragraph']}")
        print(f"range: [{payload['start']}:{payload['end']}]")
        print(f"author: {payload['author']}")
        print(f"date: {payload['date']}")
        print("text:")
        print(payload["text"])


def cmd_comments(args: argparse.Namespace) -> int:
    doc = _read_doc(args.file, keep_comments=args.keep_comments)
    items = list(doc.comments())

    if args.json:
        _dump_json([_comment_to_dict(c) for c in items])
    else:
        for c in items:
            preview = (c.text or "").replace("\n", "\\n")
            if len(preview) > 80:
                preview = preview[:80] + "…"
            print(f"{c.path}\t{c.start}:{c.end}\t{c.author}\t{preview}")
    return 0


def cmd_annotate(args: argparse.Namespace) -> int:
    doc = _read_doc(args.input, keep_comments=args.keep_comments)

    ops = _collect_annotate_ops(args)
    if not ops:
        print(
            "error: no annotation specified; use --path/--text or --spec",
            file=sys.stderr,
        )
        return 2

    results: list[dict[str, Any]] = []
    for op in ops:
        comment = _apply_annotate_op(doc, op)
        results.append(_comment_to_dict(comment))

    Path(args.output).write_bytes(doc.render())

    if args.json:
        _dump_json({"output": str(args.output), "added": results})
    else:
        print(f"wrote: {args.output}")
        for r in results:
            print(f"  + {r['path']}  {r['author']}  [{r['start']}:{r['end']}]")
    return 0


def _collect_annotate_ops(args: argparse.Namespace) -> list[dict[str, Any]]:
    ops: list[dict[str, Any]] = []

    if args.spec is not None:
        data = json.loads(Path(args.spec).read_text(encoding="utf-8"))
        if not isinstance(data, list):
            raise ValueError("--spec file must contain a JSON array")
        for item in data:
            if not isinstance(item, dict):
                raise ValueError(f"invalid op entry in --spec: {item!r}")
            ops.append(item)

    if args.path is not None or args.text is not None:
        if args.path is None or args.text is None:
            raise ValueError("--path and --text must be used together")
        ops.append(
            {
                "path": args.path,
                "text": args.text,
                "start": args.start,
                "end": args.end,
                "author": args.author,
            }
        )

    return ops


def _apply_annotate_op(doc: DocxDocument, op: dict[str, Any]) -> Comment:
    path = op.get("path")
    text = op.get("text")
    if not isinstance(path, str) or not path:
        raise ValueError(f"op missing 'path': {op!r}")
    if not isinstance(text, str):
        raise ValueError(f"op missing 'text': {op!r}")

    target = doc.resolve(path)
    if not isinstance(target, Paragraph):
        raise ValueError(
            f"path does not target a paragraph: {path!r} -> {type(target).__name__}"
        )

    start = int(op.get("start") or 0)
    end_raw = op.get("end")
    end = int(end_raw) if end_raw is not None else None
    author = str(op.get("author") or "docxnote")

    return target.comment(text, start=start, end=end, author=author)


# ---------------------------------------------------------------------------
# argparse wiring
# ---------------------------------------------------------------------------


def _add_keep_comments(p: argparse.ArgumentParser, *, default: bool) -> None:
    """Add symmetric --keep-comments / --no-keep-comments switches."""
    group = p.add_mutually_exclusive_group()
    group.add_argument(
        "--keep-comments",
        dest="keep_comments",
        action="store_true",
        help="preserve existing DOCX comments while parsing",
    )
    group.add_argument(
        "--no-keep-comments",
        dest="keep_comments",
        action="store_false",
        help="strip existing DOCX comments while parsing",
    )
    p.set_defaults(keep_comments=default)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docxnote",
        description=(
            "Read paragraphs and Word comments, or annotate a DOCX by addressable paths."
        ),
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # list
    p_list = sub.add_parser("list", help="list all paragraph paths")
    p_list.add_argument("file", type=Path, help="input .docx file")
    p_list.add_argument(
        "--text", action="store_true", help="include paragraph text preview"
    )
    p_list.add_argument(
        "--text-limit",
        type=int,
        default=80,
        help="truncate text preview to N chars (0 = no limit, default 80)",
    )
    p_list.add_argument("--json", action="store_true", help="output JSON")
    _add_keep_comments(p_list, default=False)
    p_list.set_defaults(func=cmd_list)

    # show
    p_show = sub.add_parser("show", help="resolve a path and print the target")
    p_show.add_argument("file", type=Path, help="input .docx file")
    p_show.add_argument(
        "path",
        help='addressable path, e.g. "p:0", "t:0/r:1/c:2/p:0", "p:0#3"',
    )
    p_show.add_argument("--json", action="store_true", help="output JSON")
    _add_keep_comments(p_show, default=True)
    p_show.set_defaults(func=cmd_show)

    # comments
    p_cmts = sub.add_parser("comments", help="list every Word comment in the document")
    p_cmts.add_argument("file", type=Path, help="input .docx file")
    p_cmts.add_argument("--json", action="store_true", help="output JSON")
    _add_keep_comments(p_cmts, default=True)
    p_cmts.set_defaults(func=cmd_comments)

    # annotate
    p_ann = sub.add_parser("annotate", help="add comment(s) and write a new .docx")
    p_ann.add_argument("input", type=Path, help="input .docx file")
    p_ann.add_argument("output", type=Path, help="output .docx file")
    p_ann.add_argument("--path", help='paragraph path, e.g. "p:0"')
    p_ann.add_argument("--text", help="comment body")
    p_ann.add_argument(
        "--start", type=int, default=0, help="start char offset (default 0)"
    )
    p_ann.add_argument(
        "--end",
        type=int,
        default=None,
        help="end char offset (default: end of paragraph)",
    )
    p_ann.add_argument(
        "--author", default="docxnote", help="comment author (default: docxnote)"
    )
    p_ann.add_argument(
        "--spec",
        type=Path,
        help="JSON array file: [{path, text, start?, end?, author?}, ...]",
    )
    p_ann.add_argument(
        "--json", action="store_true", help="emit JSON summary to stdout"
    )
    _add_keep_comments(p_ann, default=False)
    p_ann.set_defaults(func=cmd_annotate)

    return parser


# ---------------------------------------------------------------------------
# entry point
# ---------------------------------------------------------------------------


def main(argv: Sequence[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    try:
        return int(args.func(args) or 0)
    except FileNotFoundError as e:
        print(f"error: file not found: {e.filename or e}", file=sys.stderr)
        return 2
    except (LookupError, ValueError) as e:
        print(f"error: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
