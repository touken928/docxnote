"""docxnote - 轻量级 DOCX 批注引擎"""

from .document import DocxDocument
from .paragraph import Paragraph
from .table import Table, Cell
from .comments import Comment

__all__ = ["DocxDocument", "Paragraph", "Table", "Cell", "Comment"]
