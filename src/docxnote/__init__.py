"""Lightweight Python library for reading, editing, and annotating DOCX files.

Public API:
    DocxDocument: Open, create, modify, and save documents.
    Paragraph, Table, Cell: Access document content and structure.
    Comment: Represent document comments.
    build_segment, comment_path, join_path, parse_path: Build and parse paths
        for stable references to paragraphs, tables, cells, and comments.

Comment ranges use ``[start, end)`` character offsets in ``paragraph.text``.
"""

from .document import DocxDocument
from .paragraph import Paragraph
from .table import Table, Cell
from .comments import Comment
from .paths import build_segment, comment_path, join_path, parse_path

__all__ = [
    "DocxDocument",
    "Paragraph",
    "Table",
    "Cell",
    "Comment",
    "build_segment",
    "comment_path",
    "join_path",
    "parse_path",
]
