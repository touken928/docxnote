"""Overlapping anchors remain exact before and after repeated serialization."""

import pytest

from docxnote import DocxDocument
from tests.support.docx import build_docx
from tests.support.comments import make_docx_with_comment_markers


@pytest.mark.parametrize(
    "ranges",
    [
        pytest.param([(1, 3), (0, 2)], id="overlapping"),
        pytest.param([(0, 2), (2, 4), (4, 6)], id="adjacent"),
        pytest.param([(0, 6), (2, 4), (8, 10)], id="nested"),
        pytest.param([(0, 6), (0, 3)], id="same-start"),
        pytest.param([(0, 3), (3, 6), (6, 9)], id="sequential"),
        pytest.param([(2, 5), (1, 4), (3, 6)], id="reannotated"),
        pytest.param([(0, 10), (1, 3)], id="whole-then-partial"),
        pytest.param(
            [(0, 3), (0, 5), (5, 10), (7, 10), (2, 8)], id="shared-boundaries"
        ),
    ],
)
@pytest.mark.parametrize(
    "reparse_each", [False, True], ids=["in-memory", "reparse-between-writes"]
)
def test_ranges_keep_coordinates_and_start_order(ranges, reparse_each):
    document = DocxDocument.parse(build_docx(["ABCDEFGHIJ"]))
    expected = []
    for index, (start, end) in enumerate(ranges):
        paragraph = next(document.iter_paragraphs())
        added = paragraph.comment(str(index), start, end, author=f"u{index}")
        assert (added.start, added.end) == (start, end)
        expected.append((str(index), start, end, f"u{index}"))
        if reparse_each:
            document = DocxDocument.parse(document.render(), keep_comments=True)
        ordered = sorted(expected, key=lambda row: row[1])
        for current in (
            document,
            DocxDocument.parse(document.render(), keep_comments=True),
        ):
            paragraph = next(current.iter_paragraphs())
            assert paragraph.text == "ABCDEFGHIJ"
            assert [
                (c.text, c.start, c.end, c.author) for c in paragraph.comments
            ] == ordered
            assert [
                (c.text, c.start, c.end, c.author) for c in current.comments()
            ] == ordered


def test_reused_active_count_does_not_reorder_equal_start_markers():
    data = make_docx_with_comment_markers(
        ["text"],
        {
            0: [
                ("commentRangeStart", 0),
                ("commentRangeStart", 1),
                ("commentRangeEnd", 0),
                ("commentRangeStart", 2),
                ("commentRangeEnd", 2),
                ("commentRangeEnd", 1),
            ]
        },
        comment_meta={i: (str(i), "author", None) for i in range(3)},
    )
    document = DocxDocument.parse(data, keep_comments=True)
    for current in (
        document,
        DocxDocument.parse(document.render(), keep_comments=True),
    ):
        paragraph = next(current.iter_paragraphs())
        expected = ["p:0#0", "p:0#1", "p:0#2"]
        assert [c.path for c in paragraph.comments] == expected
        assert [c.path for c in current.comments()] == expected
