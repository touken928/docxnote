"""测试批注功能"""

import pytest
import zipfile
from datetime import datetime, timezone
from io import BytesIO

from docx import Document as PythonDocxDocument
from lxml import etree

from docxnote import DocxDocument, Paragraph, Table
from docxnote.namespaces import NS


class TestComments:
    """测试批注添加和渲染"""

    def test_add_single_comment(self, simple_doc):
        """测试添加单个批注"""
        doc = DocxDocument.parse(simple_doc)

        # 添加批注
        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                block.comment("测试批注", end=5, author="tester")
                break

        # 验证批注被添加
        assert len(doc._comments) == 1

        comment_id, text, author, _date = doc._comments[0]
        assert text == "测试批注"
        assert author == "tester"

    def test_add_multiple_comments(self, simple_doc):
        """测试添加多个批注"""
        doc = DocxDocument.parse(simple_doc)

        # 添加多个批注
        count = 0
        for block in doc.blocks():
            if isinstance(block, Paragraph) and len(block.text) > 5:
                block.comment(f"批注{count}", end=5, author=f"作者{count}")
                count += 1
                if count >= 3:
                    break

        assert len(doc._comments) >= 3

    def test_comment_on_table_cell(self, table_doc):
        """测试为表格单元格添加批注"""
        doc = DocxDocument.parse(table_doc)

        # 查找表格
        tables = [b for b in doc.blocks() if isinstance(b, Table)]
        assert len(tables) > 0

        table = tables[0]
        cell = table[0, 0]

        # 为单元格内的段落添加批注
        for block in cell.blocks():
            if isinstance(block, Paragraph) and block.text:
                block.comment("单元格批注", end=3, author="tester")
                break

        # 渲染应该成功
        output = doc.render()
        assert output is not None

        # 验证批注存在
        with zipfile.ZipFile(BytesIO(output)) as z:
            assert "word/comments.xml" in z.namelist()

    def test_comment_range(self, simple_doc):
        """测试批注范围"""
        doc = DocxDocument.parse(simple_doc)

        for block in doc.blocks():
            if isinstance(block, Paragraph) and len(block.text) >= 10:
                # 添加不同范围的批注
                block.comment("前5个字符", start=0, end=5, author="tester1")
                block.comment("中间部分", start=5, end=10, author="tester2")
                break

        assert len(doc._comments) >= 2

    def test_comment_full_paragraph(self, simple_doc):
        """测试为整个段落添加批注"""
        doc = DocxDocument.parse(simple_doc)

        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                text_len = len(block.text)
                block.comment("整段批注", start=0, end=text_len, author="tester")
                break

        output = doc.render()
        assert output is not None

    def test_comment_with_default_author(self, simple_doc):
        """测试使用默认作者的批注"""
        doc = DocxDocument.parse(simple_doc)

        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                block.comment("默认作者批注", end=5)
                break

        # 验证使用了默认作者
        comment_id, text, author, _date = doc._comments[0]
        assert author == "docxnote"

    def test_comment_custom_date_in_comments_xml(self, simple_doc):
        """自定义批注时间写入 comments.xml 的 w:date"""
        fixed = datetime(2020, 6, 15, 12, 30, 0, tzinfo=timezone.utc)
        doc = DocxDocument.parse(simple_doc)

        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                block.comment("dated", end=3, author="tester", date=fixed)
                break

        out = doc.render()
        with zipfile.ZipFile(BytesIO(out)) as z:
            root = etree.fromstring(z.read("word/comments.xml"))
        el = root.find(f"{{{NS['w']}}}comment")
        assert el is not None
        assert el.get(f"{{{NS['w']}}}date") == "2020-06-15T12:30:00Z"

    def test_comment_empty_text(self, simple_doc):
        """测试空批注文本"""
        doc = DocxDocument.parse(simple_doc)

        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                block.comment("", end=5, author="tester")
                break

        # 应该能够渲染
        output = doc.render()
        assert output is not None

    def test_comment_on_empty_paragraph_raises(self):
        """空段落无法添加批注，应抛出 ValueError"""
        pd_doc = PythonDocxDocument()
        pd_doc.add_paragraph("")
        buf = BytesIO()
        pd_doc.save(buf)
        buf.seek(0)

        doc = DocxDocument.parse(buf.getvalue())

        for block in doc.blocks():
            if isinstance(block, Paragraph) and not block.text:
                with pytest.raises(ValueError, match="no text runs"):
                    block.comment("空段落批注", author="tester")
                return

        pytest.fail("No empty paragraph found in test document")

    def test_comment_beyond_text_length(self, simple_doc):
        """测试超出文本长度的批注"""
        doc = DocxDocument.parse(simple_doc)

        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                text_len = len(block.text)
                # end 超出文本长度
                block.comment("超长批注", start=0, end=text_len + 100, author="tester")
                break

        # 应该能够渲染
        output = doc.render()
        assert output is not None

    def test_comment_zero_length(self, simple_doc):
        """测试零长度批注"""
        doc = DocxDocument.parse(simple_doc)

        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                # start == end
                block.comment("零长度批注", start=0, end=0, author="tester")
                break

        output = doc.render()
        assert output is not None

    def test_render_without_comments(self, simple_doc):
        """测试不添加批注的渲染"""
        doc = DocxDocument.parse(simple_doc)

        # 不添加任何批注
        output = doc.render()

        assert output is not None

        # 不应该有 comments.xml
        with zipfile.ZipFile(BytesIO(output)) as z:
            assert "word/comments.xml" not in z.namelist()

    def test_comment_preservation_after_render(self, simple_doc):
        """测试渲染后批注信息保留"""
        doc = DocxDocument.parse(simple_doc)

        # 添加批注
        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                block.comment("测试批注", end=5, author="tester")
                break

        # 第一次渲染
        output1 = doc.render()

        # 第二次渲染
        output2 = doc.render()

        # 两次渲染的结果应该一致
        assert len(output1) == len(output2)
