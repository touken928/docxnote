"""测试：嵌套批注按 commentRangeStart 的 XML 文档顺序返回。

``paragraph.comments`` / ``doc.comments()`` 的顺序必须跟随
``commentRangeStart`` 标记在 XML 中的出现顺序（外层批注先于嵌套内层
批注），而不是 ``commentRangeEnd`` 的闭合顺序。
"""

from tests.comments._helpers import build_docx
from docxnote import DocxDocument, Paragraph


def _first_paragraph(doc: DocxDocument) -> Paragraph:
    for block in doc.blocks():
        if isinstance(block, Paragraph) and block.text:
            return block
    raise AssertionError("no non-empty paragraph found")


class TestNestedCommentOrder:
    def test_nested_comments_ordered_by_range_start_document_order(self):
        doc = DocxDocument.parse(build_docx(["ABCDEFGHIJ"]))
        para = _first_paragraph(doc)

        para.comment("outer", start=0, end=6, author="u1")
        para.comment("inner", start=2, end=4, author="u2")
        para.comment("tail", start=8, end=10, author="u3")

        comments = para.comments
        assert [(c.text, c.start, c.end) for c in comments] == [
            ("outer", 0, 6),
            ("inner", 2, 4),
            ("tail", 8, 10),
        ]

    def test_same_start_nested_comments_follow_start_marker_order(self):
        """同一起点的嵌套批注，外层（先插入的 start 标记）在前。"""
        doc = DocxDocument.parse(build_docx(["ABCDEFGHIJ"]))
        para = _first_paragraph(doc)

        para.comment("outer", start=0, end=6, author="u1")
        para.comment("inner", start=0, end=3, author="u2")

        comments = para.comments
        assert [(c.text, c.start, c.end) for c in comments] == [
            ("outer", 0, 6),
            ("inner", 0, 3),
        ]

    def test_nested_order_survives_render_roundtrip(self):
        doc = DocxDocument.parse(build_docx(["ABCDEFGHIJ"]))
        para = _first_paragraph(doc)
        para.comment("outer", start=0, end=6, author="u1")
        para.comment("inner", start=2, end=4, author="u2")

        out = doc.render()
        doc2 = DocxDocument.parse(out, keep_comments=True)
        para2 = _first_paragraph(doc2)

        comments = para2.comments
        assert [(c.text, c.start, c.end) for c in comments] == [
            ("outer", 0, 6),
            ("inner", 2, 4),
        ]

    def test_document_comments_follow_same_order(self):
        doc = DocxDocument.parse(build_docx(["ABCDEFGHIJ"]))
        para = _first_paragraph(doc)
        para.comment("outer", start=0, end=6, author="u1")
        para.comment("inner", start=2, end=4, author="u2")

        comments = doc.comments()
        assert [c.text for c in comments] == ["outer", "inner"]
