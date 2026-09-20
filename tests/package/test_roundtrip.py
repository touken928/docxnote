"""Package serialization writes consistent references and preserves unrelated parts."""

from lxml import etree
import pytest

from docxnote import DocxDocument
from docxnote.namespaces import NS
from tests.support.comments import (
    COMMENTS_CONTENT_TYPE,
    COMMENTS_REL_TYPE,
    make_docx_with_comment_markers,
)
from tests.support.docx import build_docx, read_parts, xml_part


@pytest.mark.parametrize("count", [0, 1, 3])
def test_new_comments_have_matching_parts_relationships_and_anchors(count):
    data = build_docx(["ABCDEFGHIJ"])
    document = DocxDocument.parse(data)
    paragraph = next(document.iter_paragraphs())
    for index in range(count):
        paragraph.comment(f"note {index}", index, index + 1, author=f"author {index}")
    output = document.render()
    parts = read_parts(output)
    assert read_parts(document.render()) == parts
    changed = {
        "word/document.xml",
        "word/comments.xml",
        "word/_rels/document.xml.rels",
        "[Content_Types].xml",
    }
    for name, original in read_parts(data).items():
        if name not in changed:
            assert parts[name] == original
    root = xml_part(output)
    assert root.tag == f"{{{NS['w']}}}document"
    rels = [
        r
        for r in etree.fromstring(parts["word/_rels/document.xml.rels"])
        if r.get("Type") == COMMENTS_REL_TYPE
    ]
    overrides = [
        r
        for r in etree.fromstring(parts["[Content_Types].xml"])
        if r.get("PartName") == "/word/comments.xml"
    ]
    assert len(rels) == len(overrides) == int(count > 0)
    expected_ids = [str(index) for index in range(count)]
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        assert (
            sorted(
                element.get(f"{{{NS['w']}}}id") or ""
                for element in root.findall(f".//w:{tag}", NS)
            )
            == expected_ids
        )
    if count:
        assert rels[0].get("Target") == "comments.xml"
        assert overrides[0].get("ContentType") == COMMENTS_CONTENT_TYPE
        comments = etree.fromstring(parts["word/comments.xml"])
        assert comments.tag == f"{{{NS['w']}}}comments"
        assert [c.get(f"{{{NS['w']}}}id") for c in comments] == expected_ids
        assert [c.get(f"{{{NS['w']}}}author") for c in comments] == [
            f"author {index}" for index in range(count)
        ]
    else:
        assert "word/comments.xml" not in parts
    reopened = DocxDocument.parse(output, keep_comments=True)
    assert [(c.text, c.start, c.end) for c in reopened.comments()] == [
        (f"note {index}", index, index + 1) for index in range(count)
    ]


@pytest.mark.parametrize("keep", [False, True], ids=["strip", "preserve"])
def test_existing_comments_follow_preservation_policy(keep):
    data = make_docx_with_comment_markers(
        ["ABCDE"],
        {0: [("commentRangeStart", 7), ("commentRangeEnd", 7)]},
        comment_meta={7: ("original", "orig", None)},
    )
    document = DocxDocument.parse(data, keep_comments=keep)
    expected = ["original"] if keep else []
    assert [c.text for c in document.comments()] == expected
    assert [c.text for c in next(document.iter_paragraphs()).comments] == expected
    output = document.render()
    if not keep:
        parts = read_parts(output)
        assert "word/comments.xml" not in parts
        assert not any(
            r.get("Type") == COMMENTS_REL_TYPE
            for r in xml_part(output, "word/_rels/document.xml.rels")
        )
        assert not any(
            r.get("PartName") == "/word/comments.xml"
            for r in xml_part(output, "[Content_Types].xml")
        )
        for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
            assert xml_part(output).findall(f".//w:{tag}", NS) == []
    assert [
        c.text for c in DocxDocument.parse(output, keep_comments=True).comments()
    ] == expected
    created = next(document.iter_paragraphs()).comment("new")
    assert created.path == ("p:0#8" if keep else "p:0#0")
