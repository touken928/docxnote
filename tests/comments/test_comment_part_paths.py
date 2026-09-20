"""Comments parts follow package relationships, including their own relationships."""

from io import BytesIO
import posixpath
import zipfile

import pytest
from lxml import etree

from docxnote import DocxDocument, Paragraph
from docxnote.namespaces import NS
from tests.comments._helpers import build_docx, COMMENTS_REL_TYPE


def _relocated_comments(target, part):
    doc = DocxDocument.parse(build_docx(["ABCDE"]))
    paragraph = doc.resolve("p:0")
    assert isinstance(paragraph, Paragraph)
    paragraph.comment("original", 0, 3)
    with zipfile.ZipFile(BytesIO(doc.render())) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    comments = etree.fromstring(parts.pop("word/comments.xml"))
    paragraph = comments.find("./w:comment/w:p", NS)
    assert paragraph is not None
    hyperlink = etree.SubElement(
        paragraph, f"{{{NS['w']}}}hyperlink", {f"{{{NS['r']}}}id": "rId1"}
    )
    hyperlink.append(paragraph[0])
    parts[part] = etree.tostring(comments)
    rels = etree.fromstring(parts["word/_rels/document.xml.rels"])
    comment_rel = next(r for r in rels if r.get("Type") == COMMENTS_REL_TYPE)
    comment_rel.set("Target", target)
    original_rel = dict(comment_rel.attrib)
    parts["word/_rels/document.xml.rels"] = etree.tostring(rels)
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"/word/comments.xml", ("/" + part).encode()
    )
    rels_part = posixpath.join(
        posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels"
    )
    parts[rels_part] = (
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
        b'Target="https://example.com" TargetMode="External"/></Relationships>'
    )
    out = BytesIO()
    with zipfile.ZipFile(out, "w") as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    return out.getvalue(), rels_part, parts[rels_part], original_rel


@pytest.mark.parametrize(
    "target,part",
    [
        ("notes.xml", "word/notes.xml"),
        ("notes/../notes/comments1.xml", "word/notes/comments1.xml"),
        ("/custom/comments1.xml", "custom/comments1.xml"),
    ],
)
def test_keep_relocated_comments_and_append_without_collisions(target, part):
    data, rels_part, rels_data, original_rel = _relocated_comments(target, part)
    doc = DocxDocument.parse(data, keep_comments=True)
    assert [(c.path, c.text) for c in doc.comments()] == [("p:0#0", "original")]
    paragraph = doc.resolve("p:0")
    assert isinstance(paragraph, Paragraph)
    assert paragraph.comment("new", 3, 5).path == "p:0#1"
    rendered = doc.render()
    with zipfile.ZipFile(BytesIO(rendered)) as archive:
        assert archive.namelist().count(part) == 1
        assert "word/comments.xml" not in archive.namelist()
        assert archive.read(rels_part) == rels_data
        comments = etree.fromstring(archive.read(part))
        hyperlink = comments.find("./w:comment/w:p/w:hyperlink", NS)
        assert hyperlink is not None
        assert hyperlink.get(f"{{{NS['r']}}}id") == "rId1"
        types = etree.fromstring(archive.read("[Content_Types].xml"))
        assert sum(r.get("PartName") == "/" + part for r in types) == 1
        rels = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
        assert [dict(r.attrib) for r in rels if r.get("Type") == COMMENTS_REL_TYPE] == [
            original_rel
        ]
    reopened = DocxDocument.parse(rendered, keep_comments=True)
    assert [(c.text, c.start, c.end) for c in reopened.comments()] == [
        ("original", 0, 3),
        ("new", 3, 5),
    ]


@pytest.mark.parametrize("append", [False, True])
def test_strip_relocated_comments_removes_old_part_relationships(append):
    part = "word/notes/comments1.xml"
    data, rels_part, _, _ = _relocated_comments("notes/comments1.xml", part)
    doc = DocxDocument.parse(data)
    if append:
        paragraph = doc.resolve("p:0")
        assert isinstance(paragraph, Paragraph)
        paragraph.comment("replacement")
    rendered = doc.render()
    with zipfile.ZipFile(BytesIO(rendered)) as archive:
        assert rels_part not in archive.namelist()
        if not append:
            assert part not in archive.namelist()
            types = etree.fromstring(archive.read("[Content_Types].xml"))
            assert not any(r.get("PartName") == "/" + part for r in types)
    assert [
        c.text for c in DocxDocument.parse(rendered, keep_comments=True).comments()
    ] == (["replacement"] if append else [])
