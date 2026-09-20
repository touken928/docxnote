"""Split text must keep xml:space so Word preserves boundary whitespace."""

import pytest

from docxnote import DocxDocument
from docxnote.namespaces import NS
from tests.support.docx import build_docx, xml_part


@pytest.mark.parametrize(
    "text,start,end,fragments",
    [
        pytest.param("AB CD", 2, 4, ["AB", " C", "D"], id="leading"),
        pytest.param("AB CD", 0, 3, ["AB ", "CD"], id="trailing"),
        pytest.param("A B", 1, 2, ["A", " ", "B"], id="whitespace-only"),
    ],
)
def test_split_whitespace_is_preserved_in_xml_and_text(text, start, end, fragments):
    document = DocxDocument.parse(build_docx([text]))
    next(document.iter_paragraphs()).comment("mark", start, end)
    output = document.render()
    nodes = xml_part(output).findall(".//w:t", NS)
    assert [node.text for node in nodes] == fragments
    for node in nodes:
        if node.text and node.text != node.text.strip():
            assert node.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
    (saved,) = DocxDocument.parse(output, keep_comments=True).comments()
    assert (saved.start, saved.end, saved.paragraph.text) == (start, end, text)
