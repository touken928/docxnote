"""Creation contracts: immutable views, metadata, IDs, and normalized slices."""

from dataclasses import FrozenInstanceError

import pytest

from docxnote import Comment, DocxDocument
from tests.support.docx import build_docx


@pytest.mark.parametrize(
    "text,author",
    [
        ("note", "docxnote"),
        ("", "tester"),
        ("批注<>&\"'\n换行\t制表符", "测试者<>"),
        ("🎉 中文 English 日本語 한글", "作者👨‍💻"),
    ],
)
def test_comment_body_and_author_roundtrip(text, author):
    document = DocxDocument.parse(build_docx(["ABCDE"]))
    paragraph = next(document.iter_paragraphs())
    options = {} if author == "docxnote" else {"author": author}
    comment = paragraph.comment(text, 1, 3, **options)
    assert isinstance(comment, Comment)
    assert comment.paragraph is paragraph
    assert (comment.path, comment.start, comment.end, comment.text, comment.author) == (
        "p:0#0",
        1,
        3,
        text,
        author,
    )
    with pytest.raises(FrozenInstanceError):
        setattr(comment, "text", "changed")
    assert isinstance(paragraph.comments, tuple)
    assert isinstance(document.comments(), tuple)
    reopened = DocxDocument.parse(document.render(), keep_comments=True)
    (saved,) = reopened.comments()
    assert (saved.path, saved.start, saved.end, saved.text, saved.author) == (
        "p:0#0",
        1,
        3,
        text,
        author,
    )
    assert saved.paragraph.text == "ABCDE"


@pytest.mark.parametrize(
    "start,end,expected",
    [
        pytest.param(0, None, (0, 5), id="whole-paragraph"),
        pytest.param(1, 3, (1, 3), id="partial-run"),
        pytest.param(-4, 999, (1, 5), id="negative-and-oversized"),
        pytest.param(-2, -5, (3, 3), id="reversed"),
        pytest.param(2, 2, (2, 2), id="zero-length"),
        pytest.param(0, 0, (0, 0), id="start-boundary"),
        pytest.param(99, 100, (5, 5), id="end-boundary"),
    ],
)
def test_slice_normalization_survives_render(start, end, expected):
    document = DocxDocument.parse(build_docx(["ABCDE"]))
    paragraph = next(document.iter_paragraphs())
    comment = paragraph.comment("range", start, end)
    assert (comment.start, comment.end) == expected
    reopened = DocxDocument.parse(document.render(), keep_comments=True)
    (saved,) = reopened.comments()
    assert (saved.start, saved.end) == expected
    assert saved.paragraph.text == "ABCDE"
