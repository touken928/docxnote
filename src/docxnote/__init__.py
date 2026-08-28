"""Lightweight Python library for reading, editing, and annotating DOCX files.

Public API:
    DocxDocument: Open, create, modify, and save documents.
    Paragraph, Table, Cell: Access document content and structure.
    Comment: Represent document comments.
    UnsupportedCommentRangeError: Raised by comment range views when a
        comment range crosses paragraphs or is left unclosed.
    build_segment, comment_path, join_path, parse_path: Build and parse paths
        for stable references to paragraphs, tables, cells, and comments.

Comment ranges use ``[start, end)`` character offsets in ``paragraph.text``
and are scoped to a single paragraph.
"""

from .document import DocxDocument
from .paragraph import Paragraph
from .table import Table, Cell
from .comments import Comment, UnsupportedCommentRangeError
from .paths import build_segment, comment_path, join_path, parse_path

__all__ = [
    "DocxDocument",
    "Paragraph",
    "Table",
    "Cell",
    "Comment",
    "UnsupportedCommentRangeError",
    "build_segment",
    "comment_path",
    "join_path",
    "parse_path",
]
