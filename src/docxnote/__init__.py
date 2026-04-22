"""docxnote - 轻量级 DOCX 批注引擎"""

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
