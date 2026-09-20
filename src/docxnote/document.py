"""Public document facade for traversal, comment operations, and serialization."""

from collections.abc import Iterator
from datetime import datetime

from lxml import etree

from ._navigation import collect_blocks, iter_paragraphs, resolve_path
from ._package import DocxPackage
from ._state import DocumentState
from .comments import Comment
from .paragraph import Paragraph
from .table import Cell, Table


class DocxDocument:
    """A DOCX document with serialized access through one reentrant lock.

    Use ``parse`` to load document XML. Separate instances have independent
    mutable state; applications own file I/O and retain the bytes from ``render``.
    """

    def __init__(self, zip_data: bytes) -> None:
        self._state = DocumentState(DocxPackage(zip_data))

    @classmethod
    def parse(cls, docx_bytes: bytes, *, keep_comments: bool = False) -> "DocxDocument":
        """Load DOCX bytes, stripping existing comments unless requested otherwise."""
        document = cls(docx_bytes)
        with document._state.lock:
            document._state.load(keep_comments=keep_comments)
        return document

    def blocks(self) -> tuple[Paragraph | Table, ...]:
        """Return addressable body blocks in document order."""
        with self._state.lock:
            return collect_blocks(self._state.body, self)

    def iter_paragraphs(self) -> Iterator[Paragraph]:
        """Snapshot paragraphs, including nested tables, and release the lock before yielding."""
        with self._state.lock:
            paragraphs = tuple(iter_paragraphs(self.blocks()))
        yield from paragraphs

    def add_comment(
        self, text: str, author: str = "docxnote", *, date: datetime | None = None
    ) -> int:
        """Store an unanchored comment body and return its ID.

        Prefer ``Paragraph.comment`` to create a body and its paragraph anchors.
        ``date=None`` uses the current timezone-aware local time.
        """
        with self._state.lock:
            return self._state.comments.create(text, author, date).comment_id

    def comments(self) -> tuple[Comment, ...]:
        """Read anchored comments in document order.

        UnsupportedCommentRangeError propagates for cross-paragraph or unpaired
        anchors; parsing and rendering do not require reading these range views.
        """
        with self._state.lock:
            return tuple(
                comment
                for paragraph in self.iter_paragraphs()
                for comment in paragraph.comments
            )

    def resolve(self, path: str) -> Paragraph | Table | Cell | Comment:
        """Resolve an address to a view.

        Invalid syntax/structure raises ValueError; missing objects raise
        LookupError. A comment suffix requires a supported paragraph range.
        """
        with self._state.lock:
            return resolve_path(self, path)

    def render(self) -> bytes:
        """Serialize a consistent snapshot to DOCX bytes."""
        with self._state.lock:
            root = self._state.root
            if root is None:
                raise RuntimeError("Document XML is not loaded")
            document_xml = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            return self._state.package.render(
                document_xml,
                self._state.comments.render(),
                keep_comments=self._state.keep_comments,
            )
