"""测试：run 拆分产生的首尾空白 w:t 片段必须设置 xml:space="preserve"。

Word 会剥离 ``w:t`` 首尾空白，除非该元素显式声明 ``xml:space="preserve"``。
docxnote 在字符边界拆分 run 时可能产生以空白开头/结尾（甚至纯空白）的
``w:t`` 片段；这些片段必须携带 ``xml:space="preserve"``，否则段落文本
视图在 Word 中会发生变化。
"""

import zipfile
from io import BytesIO

from lxml import etree

from docxnote import DocxDocument, Paragraph
from docxnote.namespaces import NS

XML_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"


def _docx(paragraph_text: str) -> bytes:
    from tests.comments._helpers import build_docx

    return build_docx([paragraph_text])


def _first_paragraph(doc: DocxDocument) -> Paragraph:
    for block in doc.blocks():
        if isinstance(block, Paragraph) and block.text:
            return block
    raise AssertionError("no non-empty paragraph found")


def _w_t_fragments(docx_bytes: bytes) -> list[str]:
    with zipfile.ZipFile(BytesIO(docx_bytes)) as z:
        tree = etree.fromstring(z.read("word/document.xml"))
    return [t.text or "" for t in tree.findall(".//w:t", NS)]


def _assert_whitespace_fragments_preserved(docx_bytes: bytes) -> list[str]:
    """所有首尾含空白的 w:t 都必须带 xml:space="preserve"，返回片段文本。"""
    with zipfile.ZipFile(BytesIO(docx_bytes)) as z:
        tree = etree.fromstring(z.read("word/document.xml"))
    whitespace_fragments: list[str] = []
    for t in tree.findall(".//w:t", NS):
        text = t.text or ""
        if text != text.strip():
            whitespace_fragments.append(text)
            assert t.get(XML_SPACE_ATTR) == "preserve", (
                f"w:t fragment {text!r} has leading/trailing whitespace but "
                "is missing xml:space='preserve'"
            )
    return whitespace_fragments


class TestSplitRunWhitespacePreservation:
    def test_split_before_whitespace_sets_preserve(self):
        """拆分出以空白开头的片段（" C"）时设置 xml:space=preserve。"""
        doc = DocxDocument.parse(_docx("AB CD"))
        para = _first_paragraph(doc)

        created = para.comment("mark", start=2, end=4, author="tester")
        assert (created.start, created.end) == (2, 4)

        fragments = _assert_whitespace_fragments_preserved(doc.render())
        assert " C" in fragments

    def test_split_after_whitespace_sets_preserve(self):
        """拆分出以空白结尾的片段（"AB "）时设置 xml:space=preserve。"""
        doc = DocxDocument.parse(_docx("AB CD"))
        para = _first_paragraph(doc)

        created = para.comment("mark", start=0, end=3, author="tester")
        assert (created.start, created.end) == (0, 3)

        fragments = _assert_whitespace_fragments_preserved(doc.render())
        assert "AB " in fragments

    def test_split_whitespace_only_fragment_sets_preserve(self):
        """拆分出纯空白片段（" "）时设置 xml:space=preserve。"""
        doc = DocxDocument.parse(_docx("A B"))
        para = _first_paragraph(doc)

        created = para.comment("mark", start=1, end=2, author="tester")
        assert (created.start, created.end) == (1, 2)

        fragments = _assert_whitespace_fragments_preserved(doc.render())
        assert " " in fragments

    def test_whitespace_text_view_survives_roundtrip(self):
        """拆分后重新解析，段落文本与批注范围保持不变。"""
        doc = DocxDocument.parse(_docx("A B"))
        para = _first_paragraph(doc)
        para.comment("mark", start=1, end=2, author="tester")

        out = doc.render()
        assert _w_t_fragments(out) == ["A", " ", "B"]

        doc2 = DocxDocument.parse(out, keep_comments=True)
        para2 = _first_paragraph(doc2)
        assert para2.text == "A B"
        assert [(c.start, c.end) for c in para2.comments] == [(1, 2)]
