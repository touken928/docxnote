"""Construct and navigate public views; callers hold the shared document lock."""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from typing import TYPE_CHECKING

from lxml import etree

from ._paths import build_segment, comment_path, join_path, parse_path
from ._state import DocumentOwner
from ._xml.blocks import iter_block_elements
from ._xml.namespaces import local_name

if TYPE_CHECKING:
    from .comments import Comment
    from .paragraph import Paragraph
    from .table import Cell, Table


def collect_blocks(
    container: etree._Element | None, owner: DocumentOwner, parent_path: str = ""
) -> tuple[Paragraph | Table, ...]:
    # The factory is the only layer that needs concrete view constructors.
    # Deferred imports let Cell delegate child construction back to the factory.
    from .paragraph import Paragraph
    from .table import Table

    if container is None:
        return ()
    blocks: list[Paragraph | Table] = []
    counts = {"p": 0, "t": 0}
    for element in iter_block_elements(container):
        kind = "p" if local_name(element) == "p" else "t"
        path = join_path(parent_path, build_segment(kind, counts[kind]))
        counts[kind] += 1
        view = (
            Paragraph(element, owner, path)
            if kind == "p"
            else Table(element, owner, path)
        )
        blocks.append(view)
    return tuple(blocks)


def iter_paragraphs(blocks: Iterable[Paragraph | Table]) -> Iterator[Paragraph]:
    from .paragraph import Paragraph

    for block in blocks:
        if isinstance(block, Paragraph):
            yield block
        else:
            seen: set[Cell] = set()
            rows, cols = block.shape()
            for row in range(rows):
                for col in range(cols):
                    cell = block[row, col]
                    if cell not in seen:
                        seen.add(cell)
                        yield from iter_paragraphs(cell.blocks())


def resolve_path(owner: DocumentOwner, path: str) -> Paragraph | Table | Cell | Comment:
    from .paragraph import Paragraph

    segments, comment_id = parse_path(path)
    current = _navigate(owner, segments)
    if comment_id is None:
        return current
    if not isinstance(current, Paragraph):
        raise ValueError(
            f"comment path must target a paragraph, got {type(current).__name__}"
        )
    canonical_path = comment_path(current.path, comment_id)
    for comment in current.comments:
        if comment.path == canonical_path:
            return comment
    raise LookupError(f"comment {comment_id} not found on paragraph {current.path!r}")


def _select_block(
    blocks: tuple[Paragraph | Table, ...], kind: str, index: int, parent_path: str = ""
) -> Paragraph | Table:
    from .paragraph import Paragraph
    from .table import Table

    block_type = Paragraph if kind == "p" else Table
    matches = [block for block in blocks if isinstance(block, block_type)]
    if index >= len(matches):
        suffix = f" in {parent_path!r}" if parent_path else ""
        name = "paragraph" if kind == "p" else "table"
        raise LookupError(f"{name} index out of range: {kind}:{index}{suffix}")
    return matches[index]


def _navigate(
    owner: DocumentOwner, segments: list[tuple[str, int]]
) -> Paragraph | Table | Cell:
    from .table import Cell, Table

    first_kind, first_index = segments[0]
    if first_kind not in {"p", "t"}:
        raise ValueError(f"first segment must be 'p' or 't', got {first_kind!r}")
    current: Paragraph | Table | Cell = _select_block(
        collect_blocks(owner._state.body, owner), first_kind, first_index
    )
    position = 1
    while position < len(segments):
        kind, index = segments[position]
        if isinstance(current, Table):
            if kind != "r":
                raise ValueError(f"after table expected 'r:', got {kind!r} in path")
            if position + 1 >= len(segments):
                raise ValueError("path has 'r:' without following 'c:'")
            column_kind, column = segments[position + 1]
            if column_kind != "c":
                raise ValueError(f"after 'r:' expected 'c:', got {column_kind!r}")
            rows, cols = current.shape()
            if index >= rows or column >= cols:
                raise LookupError(
                    f"cell out of bounds: r:{index}/c:{column} in {current.path!r}"
                )
            current = current[index, column]
            position += 2
        elif isinstance(current, Cell):
            if kind not in {"p", "t"}:
                raise ValueError(f"after cell expected 'p:' or 't:', got {kind!r}")
            current = _select_block(current.blocks(), kind, index, current.path)
            position += 1
        else:
            raise ValueError(
                "paragraph has no navigable children; use '#<id>' for comments"
            )
    return current
