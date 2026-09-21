"""Public paragraph text and comment views."""

from datetime import datetime

from lxml import etree

from ._comment_store import CommentRecord
from ._paths import comment_path
from ._state import DocumentOwner
from ._xml.ranges import collect_comment_ranges, insert_comment_markers
from ._xml.text import iter_runs, paragraph_text
from .comments import Comment


class Paragraph:
    """A host paragraph; embedded text boxes have separate text coordinates."""

    def __init__(
        self, element: etree._Element, document: DocumentOwner, path: str = ""
    ) -> None:
        self._element = element
        self._state = document._state
        self._path = path
        self._text_cache: str | None = None

    @property
    def path(self) -> str:
        """The paragraph's address, such as ``p:0`` or ``t:0/r:1/c:2/p:0``."""
        return self._path

    @property
    def text(self) -> str:
        """Read text, tabs, and line breaks without descending into run contents."""
        with self._state.lock:
            # Comment insertion changes runs and markers, but not visible text.
            if self._text_cache is None:
                self._text_cache = paragraph_text(self._element)
            return self._text_cache

    def comment(
        self,
        text: str,
        start: int = 0,
        end: int | None = None,
        *,
        author: str = "docxnote",
        date: datetime | None = None,
    ) -> Comment:
        """Add a comment anchored to a normalized Python-slice text range.

        Reversed ranges collapse at the normalized start. ``date=None`` uses
        the current timezone-aware local time. Paragraphs without runs raise
        ValueError before allocating a comment record.
        """
        with self._state.lock:
            safe_start, safe_end, _ = slice(start, end).indices(len(self.text))
            safe_end = max(safe_start, safe_end)
            if next(iter_runs(self._element), None) is None:
                raise ValueError(
                    f"Cannot add comment to paragraph '{self._path}': paragraph has no text runs"
                )
            record = self._state.comments.create(text, author, date)
            insert_comment_markers(
                self._element, record.comment_id, safe_start, safe_end
            )
            return self._comment_view(record, safe_start, safe_end)

    def _comment_view(self, record: CommentRecord, start: int, end: int) -> Comment:
        return Comment(
            paragraph=self,
            path=comment_path(self._path, record.comment_id),
            start=start,
            end=end,
            text=record.text,
            author=record.author,
            date=record.date,
        )

    @property
    def comments(self) -> tuple[Comment, ...]:
        """Read current comments in start-marker order, with outer ranges first.

        Cross-paragraph or unpaired anchors raise UnsupportedCommentRangeError.
        Each access reads shared state, so older wrappers observe new comments.
        """
        with self._state.lock:
            ranges = collect_comment_ranges(self._element, self._path)
            result: list[Comment] = []
            length = len(self.text)
            for span in ranges:
                record = self._state.comments.get(span.comment_id)
                if record is None:
                    continue
                start = max(0, min(span.start, length))
                end = max(start, min(span.end, length))
                result.append(self._comment_view(record, start, end))
            return tuple(result)
