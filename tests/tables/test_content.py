"""Physical cells retain empty and multiple paragraphs, formatting, and special text."""

from docx import Document
import pytest

from docxnote import DocxDocument, Paragraph, Table
from tests.support.docx import save_docx


@pytest.mark.parametrize(
    "texts",
    [
        [""],
        ["one"],
        ["first", "second", "third"],
        ["<>&\"'", "中文 🎉", "line\nbreak\ttab"],
    ],
)
def test_cell_paragraphs_and_comments_roundtrip(texts):
    source = Document()
    cell = source.add_table(rows=1, cols=1).cell(0, 0)
    for index, text in enumerate(texts):
        paragraph = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
        paragraph.add_run(text).bold = True
    document = DocxDocument.parse(save_docx(source))
    table = document.resolve("t:0")
    assert isinstance(table, Table)
    blocks = table[0, 0].blocks()
    assert isinstance(blocks, tuple)
    assert all(isinstance(block, Paragraph) for block in blocks)
    assert [p.text for p in blocks if isinstance(p, Paragraph)] == texts
    if texts[0]:
        paragraph = next(document.iter_paragraphs())
        paragraph.comment("cell note", 0, 1)
        (saved,) = DocxDocument.parse(document.render(), keep_comments=True).comments()
        assert (saved.path, saved.start, saved.end) == ("t:0/r:0/c:0/p:0#0", 0, 1)
    else:
        paragraph = next(document.iter_paragraphs())
        assert paragraph.text == ""
        assert paragraph.path == "t:0/r:0/c:0/p:0"
