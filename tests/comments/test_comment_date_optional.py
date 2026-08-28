"""测试：``Comment.date`` 为 ``datetime | None``。

源 ``w:date`` 缺失、空白或非法时，高层视图返回 ``None``；合法日期与
新建批注的日期行为保持不变；``keep_comments=True`` 渲染时原样保留源
``w:date`` 属性（缺失仍缺失、非法字符串原样保留，不得改写为当前时间）。
"""

import zipfile
from datetime import datetime, timezone
from io import BytesIO

from lxml import etree

from docxnote import DocxDocument, Paragraph
from docxnote.namespaces import NS
from tests.comments._helpers import make_docx_with_comment_markers


def _first_paragraph(doc: DocxDocument) -> Paragraph:
    for block in doc.blocks():
        if isinstance(block, Paragraph) and block.text:
            return block
    raise AssertionError("no non-empty paragraph found")


def _single_comment_docx(w_date: str | None) -> bytes:
    """构造一个闭合批注范围 + 指定 w:date（None 表示不写该属性）的 docx。"""
    return make_docx_with_comment_markers(
        ["AAA"],
        {0: [("commentRangeStart", 0), ("commentRangeEnd", 0)]},
        comment_meta={0: ("body", "orig", w_date)},
    )


def _read_source_date_attr(docx_bytes: bytes) -> str | None:
    with zipfile.ZipFile(BytesIO(docx_bytes)) as z:
        tree = etree.fromstring(z.read("word/comments.xml"))
    comment = tree.find("./w:comment[@w:id='0']", NS)
    assert comment is not None
    return comment.get(f"{{{NS['w']}}}date")


class TestCommentDateOptional:
    def test_missing_w_date_returns_none(self):
        doc = DocxDocument.parse(_single_comment_docx(None), keep_comments=True)
        comments = _first_paragraph(doc).comments

        assert len(comments) == 1
        assert comments[0].date is None

    def test_blank_w_date_returns_none(self):
        for raw in ("", "   "):
            doc = DocxDocument.parse(_single_comment_docx(raw), keep_comments=True)
            comments = _first_paragraph(doc).comments

            assert len(comments) == 1
            assert comments[0].date is None, f"w:date={raw!r} should yield None"

    def test_invalid_w_date_returns_none(self):
        doc = DocxDocument.parse(
            _single_comment_docx("not-a-date"), keep_comments=True
        )
        comments = _first_paragraph(doc).comments

        assert len(comments) == 1
        assert comments[0].date is None

    def test_valid_w_date_still_parsed_as_datetime(self):
        doc = DocxDocument.parse(
            _single_comment_docx("2024-01-01T00:00:00Z"), keep_comments=True
        )
        comments = _first_paragraph(doc).comments

        assert len(comments) == 1
        assert comments[0].date == datetime(2024, 1, 1, tzinfo=timezone.utc)

    def test_new_comment_without_date_gets_aware_datetime(self):
        doc = DocxDocument.parse(_single_comment_docx(None), keep_comments=True)
        para = _first_paragraph(doc)

        created = para.comment("new", start=0, end=1, author="tester")

        assert created.date is not None
        assert created.date.tzinfo is not None

    def test_new_comment_with_explicit_date_kept(self):
        doc = DocxDocument.parse(_single_comment_docx(None), keep_comments=True)
        para = _first_paragraph(doc)
        fixed = datetime(2020, 6, 15, 12, 30, 0, tzinfo=timezone.utc)

        created = para.comment("dated", start=0, end=1, author="tester", date=fixed)

        assert created.date == fixed

    def test_keep_comments_render_preserves_missing_date_attribute(self):
        """源 comment 无 w:date → 渲染后仍无 w:date（不得补当前时间）。"""
        doc = DocxDocument.parse(_single_comment_docx(None), keep_comments=True)
        out = doc.render()

        assert _read_source_date_attr(out) is None

    def test_keep_comments_render_preserves_invalid_date_verbatim(self):
        """源 w:date 非法 → 渲染后原样保留该字符串。"""
        doc = DocxDocument.parse(
            _single_comment_docx("not-a-date"), keep_comments=True
        )
        out = doc.render()

        assert _read_source_date_attr(out) == "not-a-date"
