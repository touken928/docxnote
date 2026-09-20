"""Public comment views; XML records and anchors remain internal."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import TYPE_CHECKING

from ._xml.ranges import UnsupportedCommentRangeError as UnsupportedCommentRangeError

if TYPE_CHECKING:
    from .paragraph import Paragraph


@dataclass(frozen=True)
class Comment:
    """An immutable snapshot of a comment attached to a paragraph.

    ``start`` and ``end`` are half-open offsets in ``paragraph.text``. The path
    ends in the stable Word comment ID, e.g. ``t:0/r:0/c:0/p:0#3``.
    ``text`` is plain text with newlines between comment paragraphs. ``date``
    is None for a missing, blank, or invalid source w:date attribute; valid
    source timestamps retain their timezone information, if present.
    """

    paragraph: Paragraph
    path: str
    start: int
    end: int
    text: str
    author: str
    date: datetime | None
