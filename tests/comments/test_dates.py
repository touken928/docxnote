"""Source dates are interpreted without rewriting their XML representation."""

from datetime import datetime, timezone, timedelta

import pytest

from docxnote import DocxDocument
from docxnote.namespaces import NS
from tests.support.comments import make_docx_with_comment_markers
from tests.support.docx import build_docx, xml_part


@pytest.mark.parametrize(
    "raw,expected",
    [
        (None, None),
        ("", None),
        ("   ", None),
        ("not-a-date", None),
        ("2024-01-01T00:00:00Z", datetime(2024, 1, 1, tzinfo=timezone.utc)),
    ],
)
def test_source_date_view_and_verbatim_preservation(raw, expected):
    data = make_docx_with_comment_markers(
        ["AAA"],
        {0: [("commentRangeStart", 0), ("commentRangeEnd", 0)]},
        comment_meta={0: ("body", "orig", raw)},
    )
    document = DocxDocument.parse(data, keep_comments=True)
    assert [c.date for c in document.comments()] == [expected]
    output = document.render()
    assert xml_part(output, "word/comments.xml")[0].get(f"{{{NS['w']}}}date") == raw
    assert [
        c.date for c in DocxDocument.parse(output, keep_comments=True).comments()
    ] == [expected]


@pytest.mark.parametrize(
    "date",
    [
        None,
        datetime(2020, 6, 15, 12, 30, tzinfo=timezone.utc),
        datetime(2020, 6, 15, 20, 30, tzinfo=timezone(timedelta(hours=8))),
    ],
)
def test_new_date_is_aware_and_serialized_in_utc(date):
    document = DocxDocument.parse(build_docx(["ABCDE"]))
    before = datetime.now().astimezone()
    created = next(document.iter_paragraphs()).comment("dated", date=date)
    after = datetime.now().astimezone()
    assert created.date is not None and created.date.tzinfo is not None
    if date is None:
        assert before <= created.date <= after
    else:
        assert created.date == date
    output = document.render()
    expected = created.date.astimezone(timezone.utc).replace(microsecond=0)
    assert xml_part(output, "word/comments.xml")[0].get(
        f"{{{NS['w']}}}date"
    ) == expected.strftime("%Y-%m-%dT%H:%M:%SZ")
    assert DocxDocument.parse(output, keep_comments=True).comments()[0].date == expected
