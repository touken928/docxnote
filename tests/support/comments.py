"""Build source comments independently of the engine being tested.

Markers are inserted before the first run, in the supplied order. This permits
cross-paragraph/unpaired anchors and raw dates that the public writer cannot emit.
"""

from lxml import etree

from docxnote.namespaces import NS
from .docx import build_docx, read_parts, write_parts

W = "{" + NS["w"] + "}"
COMMENTS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)


def make_docx_with_comment_markers(
    paragraph_texts: list[str],
    markers: dict[int, list[tuple[str, int]]],
    *,
    comment_meta: dict[int, tuple[str, str, str | None]] | None = None,
) -> bytes:
    parts = read_parts(build_docx(paragraph_texts))
    root = etree.fromstring(parts["word/document.xml"])
    paragraphs = root.findall("./w:body/w:p", NS)
    for index, entries in markers.items():
        paragraph = paragraphs[index]
        offset = int(len(paragraph) > 0 and paragraph[0].tag == W + "pPr")
        for position, (tag, comment_id) in enumerate(entries, offset):
            paragraph.insert(
                position, etree.Element(W + tag, {W + "id": str(comment_id)})
            )
    parts["word/document.xml"] = etree.tostring(root)

    comments = etree.Element(W + "comments", nsmap=NS)
    for comment_id, (text, author, date) in (comment_meta or {}).items():
        comment = etree.SubElement(
            comments, W + "comment", {W + "id": str(comment_id), W + "author": author}
        )
        if date is not None:
            comment.set(W + "date", date)
        paragraph = etree.SubElement(comment, W + "p")
        run = etree.SubElement(paragraph, W + "r")
        etree.SubElement(run, W + "t").text = text
    parts["word/comments.xml"] = etree.tostring(comments)

    rels = etree.fromstring(parts["word/_rels/document.xml.rels"])
    etree.SubElement(
        rels,
        "Relationship",
        {"Id": "testComments", "Type": COMMENTS_REL_TYPE, "Target": "comments.xml"},
    )
    parts["word/_rels/document.xml.rels"] = etree.tostring(rels)
    types = etree.fromstring(parts["[Content_Types].xml"])
    etree.SubElement(
        types,
        f"{{{types.nsmap[None]}}}Override",
        {"PartName": "/word/comments.xml", "ContentType": COMMENTS_CONTENT_TYPE},
    )
    parts["[Content_Types].xml"] = etree.tostring(types)
    return write_parts(parts)
