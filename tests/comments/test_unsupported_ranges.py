"""Unsupported anchors reject range views but still allow preservation or stripping."""

import pytest
from lxml import etree

from docxnote import DocxDocument, UnsupportedCommentRangeError
from tests.support.comments import make_docx_with_comment_markers
from tests.support.docx import xml_part


@pytest.mark.parametrize(
    "markers,affected",
    [
        pytest.param(
            {0: [("commentRangeStart", 0)], 1: [("commentRangeEnd", 0)]},
            [0, 1],
            id="cross-paragraph",
        ),
        pytest.param({0: [("commentRangeStart", 0)]}, [0], id="unclosed"),
    ],
)
def test_unsupported_ranges_preserve_xml_reject_views_and_can_be_stripped(
    markers, affected
):
    data = make_docx_with_comment_markers(
        ["AAA", "BBB"], markers, comment_meta={0: ("body", "orig", None)}
    )
    document = DocxDocument.parse(data, keep_comments=True)
    assert issubclass(UnsupportedCommentRangeError, ValueError)
    assert etree.tostring(xml_part(document.render()), method="c14n") == etree.tostring(
        xml_part(data), method="c14n"
    )
    paragraphs = tuple(document.iter_paragraphs())
    assert len(paragraphs) == 2
    for index in affected:
        with pytest.raises(UnsupportedCommentRangeError):
            _ = paragraphs[index].comments
        with pytest.raises(UnsupportedCommentRangeError):
            document.resolve(f"p:{index}#0")
    with pytest.raises(UnsupportedCommentRangeError):
        document.comments()
    stripped = DocxDocument.parse(data)
    assert stripped.comments() == ()
    assert all(p.comments == () for p in stripped.iter_paragraphs())


def test_closed_source_range_has_zero_width_at_its_markers():
    data = make_docx_with_comment_markers(
        ["AAA"],
        {0: [("commentRangeStart", 0), ("commentRangeEnd", 0)]},
        comment_meta={0: ("closed", "orig", None)},
    )
    (comment,) = DocxDocument.parse(data, keep_comments=True).comments()
    assert (comment.text, comment.start, comment.end) == ("closed", 0, 0)
