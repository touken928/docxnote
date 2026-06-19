"""测试生成的 DOCX 文件 XML 语法合法性"""

import zipfile
from io import BytesIO
from lxml import etree

from docxnote import DocxDocument, Paragraph


class TestXMLValidity:
    """测试生成的 XML 结构合法性"""

    def test_document_xml_well_formed(self, simple_doc):
        """测试 document.xml 格式良好"""
        doc = DocxDocument.parse(simple_doc)

        # 添加批注
        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                doc.add_comment("测试批注", "tester")
                break

        output = doc.render()

        # 验证是有效的 ZIP
        with zipfile.ZipFile(BytesIO(output)) as z:
            # 验证 document.xml 可以解析
            doc_xml = z.read("word/document.xml")
            doc_tree = etree.fromstring(doc_xml)
            assert doc_tree is not None

            # 验证根元素
            assert etree.QName(doc_tree.tag).localname == "document"

    def test_comments_xml_well_formed(self, simple_doc):
        """测试 comments.xml 格式良好"""
        doc = DocxDocument.parse(simple_doc)

        # 添加批注
        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                block.comment("测试批注", end=5, author="tester")
                break

        output = doc.render()

        with zipfile.ZipFile(BytesIO(output)) as z:
            # 验证 comments.xml 存在且可解析
            assert "word/comments.xml" in z.namelist()

            comments_xml = z.read("word/comments.xml")
            comments_tree = etree.fromstring(comments_xml)
            assert comments_tree is not None

            # 验证根元素
            assert etree.QName(comments_tree.tag).localname == "comments"

            # 验证至少有一个批注
            assert len(comments_tree) > 0

    def test_rels_valid(self, simple_doc):
        """测试 rels 文件合法性"""
        doc = DocxDocument.parse(simple_doc)

        # 添加批注
        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                block.comment("测试", end=3)
                break

        output = doc.render()

        with zipfile.ZipFile(BytesIO(output)) as z:
            rels_xml = z.read("word/_rels/document.xml.rels")
            rels_tree = etree.fromstring(rels_xml)
            assert rels_tree is not None

            # 检查是否有 comments 关系
            has_comments_rel = False
            for rel in rels_tree:
                rel_type = rel.get("Type", "")
                if "comments" in rel_type:
                    has_comments_rel = True
                    assert rel.get("Target") == "comments.xml"
                    break

            assert has_comments_rel, "缺少 comments 关系"

    def test_content_types_valid(self, simple_doc):
        """测试 Content Types 合法性"""
        doc = DocxDocument.parse(simple_doc)

        # 添加批注
        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                block.comment("测试", end=3)
                break

        output = doc.render()

        with zipfile.ZipFile(BytesIO(output)) as z:
            ct_xml = z.read("[Content_Types].xml")
            ct_tree = etree.fromstring(ct_xml)
            assert ct_tree is not None

            # 检查是否有 comments.xml 的 Override
            has_comments_override = False
            for override in ct_tree:
                part_name = override.get("PartName", "")
                if part_name == "/word/comments.xml":
                    has_comments_override = True
                    content_type = override.get("ContentType", "")
                    assert "comments" in content_type
                    break

            assert has_comments_override, "缺少 comments.xml 的 Override"

    def test_no_comments_when_none_added(self, simple_doc):
        """测试未添加批注时不生成 comments.xml"""
        doc = DocxDocument.parse(simple_doc)

        # 不添加任何批注
        output = doc.render()

        with zipfile.ZipFile(BytesIO(output)) as z:
            # 不应该有 comments.xml
            assert "word/comments.xml" not in z.namelist()

    def test_multiple_comments_valid(self, simple_doc):
        """测试多个批注的 XML 合法性"""
        doc = DocxDocument.parse(simple_doc)

        # 添加多个批注
        count = 0
        for block in doc.blocks():
            if isinstance(block, Paragraph) and len(block.text) > 5:
                block.comment(
                    f"批注{count}", end=min(5, len(block.text)), author=f"作者{count}"
                )
                count += 1
                if count >= 3:
                    break

        output = doc.render()

        with zipfile.ZipFile(BytesIO(output)) as z:
            comments_xml = z.read("word/comments.xml")
            comments_tree = etree.fromstring(comments_xml)

            # 验证批注数量
            assert len(comments_tree) >= count

            # 验证每个批注的结构
            for comment in comments_tree:
                assert etree.QName(comment.tag).localname == "comment"
                assert (
                    comment.get(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id"
                    )
                    is not None
                )
                assert (
                    comment.get(
                        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author"
                    )
                    is not None
                )

    def test_special_characters_in_comments(self, simple_doc):
        """测试批注中的特殊字符"""
        doc = DocxDocument.parse(simple_doc)

        special_text = "批注<>&\"'测试\n换行\t制表符"

        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                block.comment(special_text, end=5, author="测试者<>")
                break

        output = doc.render()

        # 验证可以正常解析
        with zipfile.ZipFile(BytesIO(output)) as z:
            comments_xml = z.read("word/comments.xml")
            comments_tree = etree.fromstring(comments_xml)
            assert comments_tree is not None

    def test_unicode_in_comments(self, simple_doc):
        """测试批注中的 Unicode 字符"""
        doc = DocxDocument.parse(simple_doc)

        unicode_text = "批注 🎉 emoji 中文 English 日本語 한글"

        for block in doc.blocks():
            if isinstance(block, Paragraph) and block.text:
                block.comment(unicode_text, end=5, author="测试者👨‍💻")
                break

        output = doc.render()

        # 验证可以正常解析
        with zipfile.ZipFile(BytesIO(output)) as z:
            comments_xml = z.read("word/comments.xml")
            comments_tree = etree.fromstring(comments_xml)
            assert comments_tree is not None
