"""Read paragraph text and add comments to existing DOCX files.

Public API:
    DocxDocument: Parse bytes, navigate document views, and render DOCX bytes.
    Paragraph, Table, Cell: Access document content and structure.
    Comment: Represent document comments.
    UnsupportedCommentRangeError: Raised by comment range views when a
        comment range crosses paragraphs or is left unclosed.

Comment ranges use ``[start, end)`` character offsets in ``paragraph.text``
and are scoped to a single paragraph.
"""

from .document import DocxDocument
from .paragraph import Paragraph
from .table import Table, Cell
from .comments import Comment, UnsupportedCommentRangeError
from .shell import DocxShell, ShellResult

__all__ = [
    "DocxDocument",
    "Paragraph",
    "Table",
    "Cell",
    "Comment",
    "UnsupportedCommentRangeError",
    "DocxShell",
    "ShellResult",
]
