"""Existing comment XML retains metadata, formatting, and multiline content."""

from lxml import etree
import pytest

from docxnote import DocxDocument
from docxnote.namespaces import NS
from tests.support.comments import make_docx_with_comment_markers
from tests.support.docx import read_parts, write_parts, xml_part

W = "{" + NS["w"] + "}"


@pytest.mark.parametrize("append", [False, True], ids=["render-only", "append-comment"])
def test_source_comment_xml_is_preserved_verbatim_in_structure(append):
    data = make_docx_with_comment_markers(
        ["Hello world"],
        {0: [("commentRangeStart", 0), ("commentRangeEnd", 0)]},
        comment_meta={0: ("第一段批注", "orig", "2024-01-01T00:00:00Z")},
    )
    parts = read_parts(data)
    root = etree.fromstring(parts["word/comments.xml"])
    root.set("{urn:test}source", "retained")
    comment = root[0]
    comment.set(W + "initials", "ZZ")
    comment.set("{http://schemas.microsoft.com/office/word/2012/wordml}done", "1")
    run = comment.find("./w:p/w:r", NS)
    assert run is not None
    properties = etree.Element(W + "rPr")
    etree.SubElement(properties, W + "b")
    run.insert(0, properties)
    paragraph = etree.SubElement(comment, W + "p")
    run = etree.SubElement(paragraph, W + "r")
    etree.SubElement(run, W + "t").text = "第二段批注"
    etree.SubElement(run, W + "tab")
    etree.SubElement(run, W + "t").text = "tab"
    etree.SubElement(run, W + "br")
    etree.SubElement(run, W + "t").text = "break"
    parts["word/comments.xml"] = etree.tostring(root)
    expected_xml = etree.tostring(comment, method="c14n")
    expected_text = "第一段批注\n第二段批注\ttab\nbreak"
    document = DocxDocument.parse(write_parts(parts), keep_comments=True)
    assert document.comments()[0].text == expected_text
    if append:
        assert next(document.iter_paragraphs()).comment("new", 0, 5).path == "p:0#1"
    output = document.render()
    result = xml_part(output, "word/comments.xml")
    assert dict(result.attrib) == dict(root.attrib)
    assert len(result) == 1 + append
    assert etree.tostring(result[0], method="c14n") == expected_xml
    reopened = DocxDocument.parse(output, keep_comments=True)
    assert [c.text for c in reopened.comments()] == (
        [expected_text, "new"] if append else [expected_text]
    )
