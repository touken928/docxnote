"""测试可寻址单元（Paragraph / Table / Cell / Comment 的 path 与 resolve）"""

import pytest
from datetime import datetime, timezone

from docxnote import Cell, Comment, DocxDocument, Paragraph, Table


class TestBlockPaths:
    """顶层块的 path 计数独立进行（段落与表格各自从 0 开始）。"""

    def test_top_level_paragraph_paths(self, simple_doc):
        doc = DocxDocument.parse(simple_doc)
        blocks = doc.blocks()
        paragraph_paths = [b.path for b in blocks if isinstance(b, Paragraph)]
        assert paragraph_paths[:3] == ["p:0", "p:1", "p:2"]

    def test_top_level_table_paths(self, table_doc):
        doc = DocxDocument.parse(table_doc)
        tables = [b for b in doc.blocks() if isinstance(b, Table)]
        assert tables
        assert tables[0].path == "t:0"

    def test_paragraph_and_table_counters_are_independent(self, complex_doc):
        doc = DocxDocument.parse(complex_doc)
        blocks = doc.blocks()
        para_paths = [b.path for b in blocks if isinstance(b, Paragraph)]
        table_paths = [b.path for b in blocks if isinstance(b, Table)]

        # 段落与表格各自从 0 递增
        assert para_paths == [f"p:{i}" for i in range(len(para_paths))]
        assert table_paths == [f"t:{i}" for i in range(len(table_paths))]


class TestCellPaths:
    def test_cell_path_uses_origin(self, table_doc):
        doc = DocxDocument.parse(table_doc)
        table = next(b for b in doc.blocks() if isinstance(b, Table))
        cell = table[1, 2]
        assert cell.path == "t:0/r:1/c:2"

    def test_cell_paragraph_path(self, table_doc):
        doc = DocxDocument.parse(table_doc)
        table = next(b for b in doc.blocks() if isinstance(b, Table))
        cell = table[0, 0]
        paras = [b for b in cell.blocks() if isinstance(b, Paragraph)]
        assert paras
        assert paras[0].path == "t:0/r:0/c:0/p:0"

    def test_merged_cell_path_is_origin(self, merged_table_doc):
        doc = DocxDocument.parse(merged_table_doc)
        table = next(b for b in doc.blocks() if isinstance(b, Table))
        # Merged 行(0,0)-(0,1)：两个坐标指向同一单元格，path 为原点
        origin_path = table[0, 0].path
        assert origin_path == "t:0/r:0/c:0"
        assert table[0, 1].path == origin_path


class TestNestedTablePaths:
    def test_nested_paragraph_path(self, nested_table_doc):
        doc = DocxDocument.parse(nested_table_doc)
        outer = next(b for b in doc.blocks() if isinstance(b, Table))
        cell = outer[0, 1]
        inner = next(b for b in cell.blocks() if isinstance(b, Table))
        assert inner.path.startswith(cell.path + "/t:")

        inner_cell = inner[1, 0]
        assert inner_cell.path.startswith(inner.path + "/r:")

        inner_para = next(b for b in inner_cell.blocks() if isinstance(b, Paragraph))
        assert inner_para.path.startswith(inner_cell.path + "/p:")


class TestResolveByPath:
    def test_resolve_top_paragraph(self, simple_doc):
        doc = DocxDocument.parse(simple_doc)
        para = doc.resolve("p:1")
        assert isinstance(para, Paragraph)
        assert para.path == "p:1"

    def test_resolve_table(self, table_doc):
        doc = DocxDocument.parse(table_doc)
        table = doc.resolve("t:0")
        assert isinstance(table, Table)
        assert table.path == "t:0"

    def test_resolve_cell(self, table_doc):
        doc = DocxDocument.parse(table_doc)
        cell = doc.resolve("t:0/r:1/c:2")
        assert isinstance(cell, Cell)
        assert cell.path == "t:0/r:1/c:2"

    def test_resolve_cell_paragraph(self, table_doc):
        doc = DocxDocument.parse(table_doc)
        para = doc.resolve("t:0/r:0/c:0/p:0")
        assert isinstance(para, Paragraph)
        assert para.path == "t:0/r:0/c:0/p:0"

    def test_resolve_roundtrip_text(self, simple_doc):
        """通过 path 回溯得到的段落文本与原对象一致。"""
        doc = DocxDocument.parse(simple_doc)
        original = next(b for b in doc.blocks() if isinstance(b, Paragraph) and b.text)
        same = doc.resolve(original.path)
        assert isinstance(same, Paragraph)
        assert same.text == original.text

    def test_resolve_comment(self, simple_doc):
        """根据 ``#id`` 可以精确回溯到单条批注。"""
        doc = DocxDocument.parse(simple_doc)
        para = next(b for b in doc.blocks() if isinstance(b, Paragraph) and b.text)
        added = para.comment("hello", start=0, end=3, author="tester")

        got = doc.resolve(added.path)
        assert isinstance(got, Comment)
        assert got.path == added.path
        assert got.text == "hello"
        assert got.author == "tester"

    def test_resolve_invalid_path(self, simple_doc):
        doc = DocxDocument.parse(simple_doc)
        with pytest.raises(ValueError):
            doc.resolve("")

    def test_resolve_out_of_range(self, simple_doc):
        doc = DocxDocument.parse(simple_doc)
        with pytest.raises(LookupError):
            doc.resolve("p:9999")

    def test_resolve_bad_structure_after_table(self, table_doc):
        doc = DocxDocument.parse(table_doc)
        # 表格后不该直接跟 p:
        with pytest.raises(ValueError):
            doc.resolve("t:0/p:0")

    def test_resolve_missing_comment_raises_lookup(self, simple_doc):
        doc = DocxDocument.parse(simple_doc)
        para = next(b for b in doc.blocks() if isinstance(b, Paragraph) and b.text)
        with pytest.raises(LookupError):
            doc.resolve(f"{para.path}#9999")

    def test_resolve_comment_requires_paragraph(self, table_doc):
        doc = DocxDocument.parse(table_doc)
        with pytest.raises(ValueError):
            # 指向表格却带 comment suffix
            doc.resolve("t:0#0")


class TestIterParagraphs:
    def test_includes_table_cells(self, table_doc):
        doc = DocxDocument.parse(table_doc)
        paths = [p.path for p in doc.iter_paragraphs()]
        # 顶层段落存在
        assert any(p == "p:0" for p in paths)
        # 至少有若干表格内段落
        assert any(p.startswith("t:0/r:") and "/p:" in p for p in paths)

    def test_does_not_duplicate_merged_cells(self, merged_table_doc):
        """合并单元格覆盖多个坐标，但其中段落只应被遍历一次。"""
        doc = DocxDocument.parse(merged_table_doc)
        paths = [p.path for p in doc.iter_paragraphs()]
        assert len(paths) == len(set(paths))


class TestCommentPathField:
    def test_comment_path_format(self, simple_doc):
        doc = DocxDocument.parse(simple_doc)
        para = next(b for b in doc.blocks() if isinstance(b, Paragraph) and b.text)
        c = para.comment("p", start=0, end=1, author="u")
        assert c.path.startswith(para.path + "#")
        # 后缀是整数
        suffix = c.path.split("#", 1)[1]
        assert int(suffix) >= 0

    def test_comment_path_stable_after_render(self, simple_doc):
        doc = DocxDocument.parse(simple_doc)
        para = next(b for b in doc.blocks() if isinstance(b, Paragraph) and b.text)
        added = para.comment(
            "hi",
            start=0,
            end=2,
            author="u",
            date=datetime(2023, 1, 1, 0, 0, 0, tzinfo=timezone.utc),
        )

        out = doc.render()
        doc2 = DocxDocument.parse(out, keep_comments=True)
        all_comments = doc2.comments()
        found = [c for c in all_comments if c.text == "hi"]
        assert found
        # path 使用段落路径 + Word 内部 id
        assert "#" in found[0].path
        assert found[0].path.split("#", 1)[0] == added.path.split("#", 1)[0]

    def test_document_comments_path_resolvable(self, simple_doc):
        """doc.comments() 返回的每条批注都能通过 resolve(path) 回溯。"""
        doc = DocxDocument.parse(simple_doc)
        para = next(b for b in doc.blocks() if isinstance(b, Paragraph) and b.text)
        para.comment("a", start=0, end=2, author="u1")
        para.comment("b", start=2, end=4, author="u2")

        out = doc.render()
        doc2 = DocxDocument.parse(out, keep_comments=True)

        for c in doc2.comments():
            resolved = doc2.resolve(c.path)
            assert isinstance(resolved, Comment)
            assert resolved.path == c.path
            assert resolved.text == c.text
            assert resolved.author == c.author
