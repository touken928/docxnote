"""Unanchored records share IDs and package storage with paragraph comments."""

from datetime import datetime, timezone
from io import BytesIO
import zipfile

from lxml import etree
import pytest

from docxnote import DocxDocument, Paragraph
from docxnote.namespaces import NS
from tests.comments._helpers import build_docx


FIXED_DATE = datetime(2020, 6, 15, 12, 30, tzinfo=timezone.utc)


def test_unanchored_records_survive_reparse_and_reserve_ids():
    document = DocxDocument.parse(build_docx(["ABCDE"]))
    first_id = document.add_comment("unanchored", "record author", date=FIXED_DATE)
    assert first_id == 0
    assert document.comments() == ()
    with pytest.raises(LookupError):
        document.resolve("p:0#0")

    paragraph = document.resolve("p:0")
    assert isinstance(paragraph, Paragraph)
    anchored = paragraph.comment("anchored", 1, 4, date=FIXED_DATE)
    assert anchored.path == "p:0#1"
    rendered = document.render()
    with zipfile.ZipFile(BytesIO(rendered)) as archive:
        root = etree.fromstring(archive.read("word/comments.xml"))
        assert [element.get(f"{{{NS['w']}}}id") for element in root] == ["0", "1"]
        first = root[0]
        assert first.get(f"{{{NS['w']}}}author") == "record author"
        assert first.get(f"{{{NS['w']}}}date") == "2020-06-15T12:30:00Z"
        text = first.find("./w:p/w:r/w:t", NS)
        assert text is not None and text.text == "unanchored"

    reopened = DocxDocument.parse(rendered, keep_comments=True)
    assert [
        (comment.path, comment.start, comment.end) for comment in reopened.comments()
    ] == [("p:0#1", 1, 4)]
    assert reopened.add_comment("another unanchored", date=FIXED_DATE) == 2
    paragraph = reopened.resolve("p:0")
    assert isinstance(paragraph, Paragraph)
    assert paragraph.comment("next anchor", date=FIXED_DATE).path == "p:0#3"
    final = DocxDocument.parse(reopened.render(), keep_comments=True)
    assert {comment.path for comment in final.comments()} == {"p:0#1", "p:0#3"}


def test_stripping_unanchored_records_resets_ids_and_removes_part():
    document = DocxDocument.parse(build_docx(["ABCDE"]))
    document.add_comment("unanchored", date=FIXED_DATE)
    stripped = DocxDocument.parse(document.render())
    assert stripped.comments() == ()
    with zipfile.ZipFile(BytesIO(stripped.render())) as archive:
        assert "word/comments.xml" not in archive.namelist()
    paragraph = stripped.resolve("p:0")
    assert isinstance(paragraph, Paragraph)
    assert paragraph.comment("replacement", date=FIXED_DATE).path == "p:0#0"


def test_rejected_empty_paragraph_does_not_allocate_an_id():
    document = DocxDocument.parse(build_docx(["", "ABCDE"]))
    empty = document.resolve("p:0")
    assert isinstance(empty, Paragraph)
    with pytest.raises(ValueError, match="no text runs"):
        empty.comment("rejected")
    paragraph = document.resolve("p:1")
    assert isinstance(paragraph, Paragraph)
    assert paragraph.comment("accepted", date=FIXED_DATE).path == "p:1#0"
