"""tests/comments 共享的 DOCX 包注入辅助函数。

用于构造带手工批注标记（commentRangeStart / commentRangeEnd）的 docx，
以覆盖 Word 可能产出、但 docxnote 公开 API 无法直接构造的 XML 形态
（跨段落范围、未闭合范围、缺失/非法 w:date 等）。
"""

import zipfile
from io import BytesIO

from docx import Document as PythonDocxDocument
from lxml import etree

from docxnote.namespaces import NS

COMMENTS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)

# 需要重写（而非原样复制）的包内条目
_REWRITTEN_PARTS = {
    "word/document.xml",
    "word/comments.xml",
    "word/_rels/document.xml.rels",
    "[Content_Types].xml",
}


def build_docx(paragraph_texts: list[str]) -> bytes:
    """用 python-docx 构造仅含若干顶层段落的 docx bytes。"""
    pd_doc = PythonDocxDocument()
    for text in paragraph_texts:
        pd_doc.add_paragraph(text)
    buffer = BytesIO()
    pd_doc.save(buffer)
    return buffer.getvalue()


def make_docx_with_comment_markers(
    paragraph_texts: list[str],
    markers: dict[int, list[tuple[str, int]]],
    *,
    comment_meta: dict[int, tuple[str, str, str | None]] | None = None,
) -> bytes:
    """构造带手工批注标记的 docx。

    Args:
        paragraph_texts: 顶层段落文本列表。
        markers: ``{段落索引: [(marker_tag, comment_id), ...]}``；
            marker_tag 为 ``"commentRangeStart"`` / ``"commentRangeEnd"``，
            标记按给定顺序插入到该段落第一个 run 之前（pPr 之后）。
        comment_meta: ``{comment_id: (text, author, w_date_or_None)}``；
            ``w_date_or_None`` 为写入 comments.xml 的原始 ``w:date``
            字符串，``None`` 表示不写该属性。
    """
    base = build_docx(paragraph_texts)
    meta = comment_meta or {}

    out = BytesIO()
    with zipfile.ZipFile(BytesIO(base), "r") as zin, zipfile.ZipFile(
        out, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for name in zin.namelist():
            if name not in _REWRITTEN_PARTS:
                zout.writestr(name, zin.read(name))

        zout.writestr(
            "word/document.xml",
            etree.tostring(
                _document_xml_with_markers(zin, markers),
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            ),
        )
        zout.writestr(
            "word/comments.xml",
            etree.tostring(
                _comments_xml(meta),
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            ),
        )
        zout.writestr(
            "word/_rels/document.xml.rels",
            etree.tostring(
                _rels_xml(zin),
                xml_declaration=True,
                encoding="UTF-8",
            ),
        )
        zout.writestr(
            "[Content_Types].xml",
            etree.tostring(
                _content_types_xml(zin),
                xml_declaration=True,
                encoding="UTF-8",
            ),
        )

    return out.getvalue()


def _document_xml_with_markers(
    zin: zipfile.ZipFile, markers: dict[int, list[tuple[str, int]]]
) -> etree._Element:
    doc_tree = etree.fromstring(zin.read("word/document.xml"))
    paras = doc_tree.findall(".//w:body/w:p", NS)
    for para_idx, marker_list in markers.items():
        p = paras[para_idx]
        insert_at = 0
        children = list(p)
        if children:
            first_tag = etree.QName(str(children[0].tag)).localname
            if first_tag == "pPr":
                insert_at = 1
        for offset, (tag, comment_id) in enumerate(marker_list):
            el = etree.Element(
                f"{{{NS['w']}}}{tag}",
                attrib={f"{{{NS['w']}}}id": str(comment_id)},
            )
            p.insert(insert_at + offset, el)
    return doc_tree


def _comments_xml(meta: dict[int, tuple[str, str, str | None]]) -> etree._Element:
    root = etree.Element(f"{{{NS['w']}}}comments", nsmap=NS)
    for comment_id, (text, author, date_str) in meta.items():
        attrib = {
            f"{{{NS['w']}}}id": str(comment_id),
            f"{{{NS['w']}}}author": author,
        }
        if date_str is not None:
            attrib[f"{{{NS['w']}}}date"] = date_str
        comment = etree.SubElement(root, f"{{{NS['w']}}}comment", attrib=attrib)
        p = etree.SubElement(comment, f"{{{NS['w']}}}p")
        r = etree.SubElement(p, f"{{{NS['w']}}}r")
        t = etree.SubElement(r, f"{{{NS['w']}}}t")
        t.text = text
    return root


def _rels_xml(zin: zipfile.ZipFile) -> etree._Element:
    rels_tree = etree.fromstring(zin.read("word/_rels/document.xml.rels"))
    max_id = 0
    for rel in rels_tree:
        rel_id = rel.get("Id", "")
        if rel_id.startswith("rId"):
            try:
                max_id = max(max_id, int(rel_id[3:]))
            except ValueError:
                pass
    etree.SubElement(
        rels_tree,
        "Relationship",
        attrib={
            "Id": f"rId{max_id + 1}",
            "Type": COMMENTS_REL_TYPE,
            "Target": "comments.xml",
        },
    )
    return rels_tree


def _content_types_xml(zin: zipfile.ZipFile) -> etree._Element:
    ct_tree = etree.fromstring(zin.read("[Content_Types].xml"))
    ns = ct_tree.nsmap.get(
        None, "http://schemas.openxmlformats.org/package/2006/content-types"
    )
    etree.SubElement(
        ct_tree,
        f"{{{ns}}}Override",
        attrib={
            "PartName": "/word/comments.xml",
            "ContentType": COMMENTS_CONTENT_TYPE,
        },
    )
    return ct_tree
