"""测试：跨段落 / 未闭合批注范围在高层 range view 上明确抛错。

``DocxDocument.parse`` 与 ``doc.render()`` 对这类 XML 原样透传（不修改、
不报错）；但访问高层范围视图（``paragraph.comments``、``doc.comments()``、
``doc.resolve("p:0#N")``）时必须抛出 ``UnsupportedCommentRangeError``
（``ValueError`` 子类）。
"""

import zipfile
from io import BytesIO

import pytest
from lxml import etree

from docxnote import DocxDocument, Paragraph, UnsupportedCommentRangeError
from docxnote.namespaces import NS
from tests.comments._helpers import make_docx_with_comment_markers


def _cross_paragraph_docx() -> bytes:
    """commentRangeStart 在第 0 段，commentRangeEnd 在第 1 段。"""
    return make_docx_with_comment_markers(
        ["AAA", "BBB"],
        {0: [("commentRangeStart", 0)], 1: [("commentRangeEnd", 0)]},
        comment_meta={0: ("cross", "orig", "2024-01-01T00:00:00Z")},
    )


def _unclosed_docx() -> bytes:
    """只有 commentRangeStart，没有任何 commentRangeEnd。"""
    return make_docx_with_comment_markers(
        ["AAA", "BBB"],
        {0: [("commentRangeStart", 0)]},
        comment_meta={0: ("unclosed", "orig", "2024-01-01T00:00:00Z")},
    )


class TestUnsupportedCommentRangeError:
    def test_error_is_public_valueerror_subclass(self):
        assert issubclass(UnsupportedCommentRangeError, ValueError)

    def test_parse_and_render_pass_through_cross_paragraph_range(self):
        docx_bytes = _cross_paragraph_docx()

        doc = DocxDocument.parse(docx_bytes, keep_comments=True)  # 不应抛错
        out = doc.render()  # 不应抛错

        with zipfile.ZipFile(BytesIO(out)) as z:
            tree = etree.fromstring(z.read("word/document.xml"))
        starts = tree.findall(".//w:commentRangeStart", NS)
        ends = tree.findall(".//w:commentRangeEnd", NS)
        assert len(starts) == 1
        assert len(ends) == 1
        assert starts[0].get(f"{{{NS['w']}}}id") == "0"
        assert ends[0].get(f"{{{NS['w']}}}id") == "0"

    def test_cross_paragraph_range_raises_on_paragraph_comments(self):
        doc = DocxDocument.parse(_cross_paragraph_docx(), keep_comments=True)
        paras = [b for b in doc.blocks() if isinstance(b, Paragraph)]
        assert len(paras) == 2

        with pytest.raises(UnsupportedCommentRangeError):
            _ = paras[0].comments
        with pytest.raises(UnsupportedCommentRangeError):
            _ = paras[1].comments

    def test_cross_paragraph_range_raises_on_document_comments(self):
        doc = DocxDocument.parse(_cross_paragraph_docx(), keep_comments=True)

        with pytest.raises(UnsupportedCommentRangeError):
            doc.comments()

    def test_unclosed_range_raises_on_paragraph_comments(self):
        doc = DocxDocument.parse(_unclosed_docx(), keep_comments=True)
        paras = [b for b in doc.blocks() if isinstance(b, Paragraph)]

        with pytest.raises(UnsupportedCommentRangeError):
            _ = paras[0].comments

    def test_unclosed_range_raises_on_document_comments(self):
        doc = DocxDocument.parse(_unclosed_docx(), keep_comments=True)

        with pytest.raises(UnsupportedCommentRangeError):
            doc.comments()

    def test_unclosed_range_raises_on_resolve(self):
        doc = DocxDocument.parse(_unclosed_docx(), keep_comments=True)

        with pytest.raises(UnsupportedCommentRangeError):
            doc.resolve("p:0#0")

    def test_default_parse_strips_markers_so_range_view_is_empty(self):
        """keep_comments=False 剥离标记后，range view 正常返回空。"""
        doc = DocxDocument.parse(_cross_paragraph_docx())
        paras = [b for b in doc.blocks() if isinstance(b, Paragraph)]

        assert paras[0].comments == ()
        assert paras[1].comments == ()
        assert doc.comments() == ()

    def test_well_formed_range_in_same_package_still_works(self):
        """对照组：闭合的单段范围不受影响。"""
        docx_bytes = make_docx_with_comment_markers(
            ["AAA"],
            {0: [("commentRangeStart", 0), ("commentRangeEnd", 0)]},
            comment_meta={0: ("closed", "orig", "2024-01-01T00:00:00Z")},
        )
        doc = DocxDocument.parse(docx_bytes, keep_comments=True)
        paras = [b for b in doc.blocks() if isinstance(b, Paragraph)]

        comments = paras[0].comments
        assert len(comments) == 1
        assert comments[0].text == "closed"
        assert (comments[0].start, comments[0].end) == (0, 0)
