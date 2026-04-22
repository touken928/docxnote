"""测试批注阅读功能（基于文本视图）"""

from datetime import datetime, timezone

from docxnote import DocxDocument, Paragraph, Comment


def _first_paragraph_with_text(doc: DocxDocument) -> Paragraph:
    for block in doc.blocks():
        if isinstance(block, Paragraph) and (block.text or "").strip():
            return block
    raise AssertionError("no non-empty paragraph found")


class TestCommentReading:
    def test_paragraph_comments_after_writing_and_reparse(self, simple_doc):
        """写入批注后重新解析，能够从段落上读出批注列表。"""
        doc = DocxDocument.parse(simple_doc)
        para = _first_paragraph_with_text(doc)
        original_text = para.text

        para.comment("测试批注", start=1, end=3, author="tester")

        out = doc.render()
        doc2 = DocxDocument.parse(out, keep_comments=True)
        para2 = _first_paragraph_with_text(doc2)

        comments = para2.comments
        assert isinstance(comments, tuple)
        assert len(comments) >= 1

        c0 = comments[0]
        assert isinstance(c0, Comment)
        assert c0.text == "测试批注"
        assert c0.author == "tester"
        assert c0.paragraph.text == original_text

        # 范围必须落在合法区间内
        assert 0 <= c0.start <= c0.end <= len(para2.text)

    def test_document_comments_collect_all(self, simple_doc):
        """DocxDocument.comments() 能返回文档中所有批注。"""
        doc = DocxDocument.parse(simple_doc)
        para = _first_paragraph_with_text(doc)

        para.comment("c1", start=0, end=2, author="u1")
        para.comment("c2", start=2, end=4, author="u2")

        out = doc.render()
        doc2 = DocxDocument.parse(out, keep_comments=True)

        comments = doc2.comments()
        assert isinstance(comments, tuple)
        # 至少包含我们刚刚添加的两个批注
        texts = {c.text for c in comments}
        assert "c1" in texts
        assert "c2" in texts

    def test_comments_respect_keep_comments_flag(self, simple_doc):
        """keep_comments=False 时不暴露原始批注；True 时可以读取。"""
        # 构造带批注的 DOCX
        doc = DocxDocument.parse(simple_doc)
        para = _first_paragraph_with_text(doc)
        para.comment("existing", author="orig")
        out = doc.render()

        # 默认：不保留批注
        doc_no = DocxDocument.parse(out)
        assert doc_no.comments() == ()
        p_no = _first_paragraph_with_text(doc_no)
        assert p_no.comments == ()

        # 显式保留：可以读取到
        doc_yes = DocxDocument.parse(out, keep_comments=True)
        comments_yes = doc_yes.comments()
        assert any(c.text == "existing" and c.author == "orig" for c in comments_yes)

    def test_comment_date_roundtrip_in_read_model(self, simple_doc):
        """自定义日期的批注，在读取模型中也能拿到同样的 datetime（UTC 时间）。"""
        fixed = datetime(2021, 1, 2, 3, 4, 5, tzinfo=timezone.utc)
        doc = DocxDocument.parse(simple_doc)
        para = _first_paragraph_with_text(doc)

        para.comment("dated", author="tester", date=fixed)
        out = doc.render()

        doc2 = DocxDocument.parse(out, keep_comments=True)
        para2 = _first_paragraph_with_text(doc2)
        cs = para2.comments
        assert len(cs) >= 1

        matched = [c for c in cs if c.text == "dated" and c.author == "tester"]
        assert matched
        # 使用 UTC 时区比较
        assert matched[0].date.astimezone(timezone.utc) == fixed
