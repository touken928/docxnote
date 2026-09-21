"""String addresses for document blocks, cells, and anchored comments.

Paragraph/table indices are counted independently within their parent. A cell
uses its merge-origin row and column, e.g. ``t:0/r:1/c:2/p:0``. A ``#N`` suffix
addresses Word comment ID N on that paragraph. Parsing validates segments;
object navigation validates their structural order.
"""

from __future__ import annotations

SEP = "/"
COMMENT_SEP = "#"

_SEGMENT_KINDS = ("p", "t", "r", "c")


def build_segment(kind: str, index: int) -> str:
    """Build a segment such as ``p:3``, rejecting unknown kinds and negative indices."""
    if kind not in _SEGMENT_KINDS:
        raise ValueError(f"invalid segment kind: {kind!r}")
    if index < 0:
        raise ValueError(f"segment index must be >= 0, got {index}")
    return f"{kind}:{index}"


def join_path(*parts: str) -> str:
    """Join nonempty path fragments with slashes."""
    return SEP.join(p for p in parts if p)


def comment_path(paragraph_path: str, comment_id: int) -> str:
    """Append a nonnegative Word comment ID to a paragraph address."""
    if comment_id < 0:
        raise ValueError(f"comment_id must be >= 0, got {comment_id}")
    return f"{paragraph_path}{COMMENT_SEP}{comment_id}"


def parse_path(path: str) -> tuple[list[tuple[str, int]], int | None]:
    """Return ``(segments, comment_id)`` from a string address.

    Each segment is a ``(kind, index)`` pair. Surrounding whitespace and empty
    slash-separated fragments are ignored for compatibility with existing paths.
    """
    if not isinstance(path, str):
        raise TypeError("path must be a str")

    raw = path.strip()
    if not raw:
        raise ValueError("path is empty")

    comment_id: int | None = None
    if COMMENT_SEP in raw:
        base, cid_str = raw.split(COMMENT_SEP, 1)
        cid_str = cid_str.strip()
        if not cid_str:
            raise ValueError(f"missing comment id after '#': {path!r}")
        try:
            comment_id = int(cid_str)
        except ValueError as e:
            raise ValueError(f"invalid comment id: {cid_str!r}") from e
        if comment_id < 0:
            raise ValueError(f"comment_id must be >= 0, got {comment_id}")
        raw = base

    segments: list[tuple[str, int]] = []
    for part in raw.split(SEP):
        part = part.strip()
        if not part:
            continue
        if ":" not in part:
            raise ValueError(f"invalid path segment: {part!r}")
        kind, idx_str = part.split(":", 1)
        kind = kind.strip()
        idx_str = idx_str.strip()
        if kind not in _SEGMENT_KINDS:
            raise ValueError(f"invalid segment kind: {kind!r}")
        try:
            idx = int(idx_str)
        except ValueError as e:
            raise ValueError(f"invalid segment index: {idx_str!r}") from e
        if idx < 0:
            raise ValueError(f"segment index must be >= 0, got {idx}")
        segments.append((kind, idx))

    if not segments:
        raise ValueError(f"path has no segments: {path!r}")

    return segments, comment_id
