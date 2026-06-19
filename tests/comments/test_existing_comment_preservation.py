"""测试：对已有批注的 docx 再次批注时，不破坏原批注内容"""

import zipfile
from io import BytesIO

from lxml import etree
from docx import Document as PythonDocxDocument

from docxnote import DocxDocument, Paragraph
from docxnote.namespaces import NS


def _extract_comment_text(comment_el: etree._Element) -> str:
    """按 w:p 作为换行边界提取批注全文文本。"""
    parts: list[str] = []
    first_para = True
    for p in comment_el.findall(".//w:p", NS):
        if not first_para:
            parts.append("\n")
        first_para = False
        for run in p.findall(".//w:r", NS):
            for child in run:
                tag = etree.QName(child.tag).localname
                if tag == "t":
                    if child.text:
                        parts.append(child.text)
                elif tag == "br":
                    parts.append("\n")
                elif tag == "tab":
                    parts.append("\t")
    return "".join(parts)


def _make_docx_with_existing_multiline_comment() -> bytes:
    """
    生成一个带已有批注的 docx：
    - comments.xml 中 comment id=0 的内容为两段（用换行表示）
    - document.xml 中插入 commentRangeStart/End/commentReference
    - rels 与 content types 添加 comments 关系/override
    """
    pd_doc = PythonDocxDocument()
    pd_doc.add_paragraph("Hello world")

    buf = BytesIO()
    pd_doc.save(buf)
    buf.seek(0)
    base = buf.getvalue()

    comment_text_1 = "第一段批注"
    comment_text_2 = "第二段批注"

    out = BytesIO()
    with zipfile.ZipFile(BytesIO(base), "r") as zin, zipfile.ZipFile(
        out, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        # 先复制所有原始条目（将要改写的文件跳过，避免 zip 内出现重复条目）
        skip = {
            "word/document.xml",
            "word/comments.xml",
            "word/_rels/document.xml.rels",
            "[Content_Types].xml",
        }
        for name in zin.namelist():
            if name in skip:
                continue
            zout.writestr(name, zin.read(name))

        # 修改 document.xml：在第一个段落插入批注范围标记
        doc_xml = zin.read("word/document.xml")
        doc_tree = etree.fromstring(doc_xml)
        body = doc_tree.find(".//w:body", NS)
        p = body.find("./w:p", NS)
        first_run = p.find("./w:r", NS)

        crs = etree.Element(
            f"{{{NS['w']}}}commentRangeStart", attrib={f"{{{NS['w']}}}id": "0"}
        )
        cre = etree.Element(
            f"{{{NS['w']}}}commentRangeEnd", attrib={f"{{{NS['w']}}}id": "0"}
        )
        cref_run = etree.Element(f"{{{NS['w']}}}r")
        etree.SubElement(
            cref_run, f"{{{NS['w']}}}commentReference", attrib={f"{{{NS['w']}}}id": "0"}
        )

        children = list(p)
        run_pos = children.index(first_run)
        p.insert(run_pos, crs)
        p.insert(run_pos + 2, cre)
        p.insert(run_pos + 3, cref_run)

        zout.writestr(
            "word/document.xml",
            etree.tostring(
                doc_tree, xml_declaration=True, encoding="UTF-8", standalone=True
            ),
        )

        # 写入 comments.xml（两段）
        comments_root = etree.Element(f"{{{NS['w']}}}comments", nsmap=NS)
        c = etree.SubElement(
            comments_root,
            f"{{{NS['w']}}}comment",
            attrib={
                f"{{{NS['w']}}}id": "0",
                f"{{{NS['w']}}}author": "orig",
                f"{{{NS['w']}}}date": "2024-01-01T00:00:00Z",
                f"{{{NS['w']}}}initials": "O",
            },
        )
        # 第一段
        p1 = etree.SubElement(c, f"{{{NS['w']}}}p")
        r1 = etree.SubElement(p1, f"{{{NS['w']}}}r")
        t1 = etree.SubElement(r1, f"{{{NS['w']}}}t")
        t1.text = comment_text_1
        # 第二段
        p2 = etree.SubElement(c, f"{{{NS['w']}}}p")
        r2 = etree.SubElement(p2, f"{{{NS['w']}}}r")
        t2 = etree.SubElement(r2, f"{{{NS['w']}}}t")
        t2.text = comment_text_2

        zout.writestr(
            "word/comments.xml",
            etree.tostring(
                comments_root, xml_declaration=True, encoding="UTF-8", standalone=True
            ),
        )

        # 更新 rels：确保存在 comments 关系
        rels_path = "word/_rels/document.xml.rels"
        rels_xml = zin.read(rels_path)
        rels_tree = etree.fromstring(rels_xml)
        has_comments = False
        max_id = 0
        for rel in rels_tree:
            rid = rel.get("Id", "")
            if rid.startswith("rId"):
                try:
                    max_id = max(max_id, int(rid[3:]))
                except ValueError:
                    pass
            if (
                rel.get("Type")
                == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            ):
                has_comments = True
        if not has_comments:
            etree.SubElement(
                rels_tree,
                "Relationship",
                attrib={
                    "Id": f"rId{max_id + 1}",
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                    "Target": "comments.xml",
                },
            )
        zout.writestr(
            rels_path, etree.tostring(rels_tree, xml_declaration=True, encoding="UTF-8")
        )

        # 更新 content types：确保 Override /word/comments.xml
        ct_xml = zin.read("[Content_Types].xml")
        ct_tree = etree.fromstring(ct_xml)
        has_override = False
        for override in ct_tree:
            if override.get("PartName") == "/word/comments.xml":
                has_override = True
                break
        if not has_override:
            ns = ct_tree.nsmap.get(
                None, "http://schemas.openxmlformats.org/package/2006/content-types"
            )
            ct_tree.append(
                etree.Element(
                    f"{{{ns}}}Override",
                    attrib={
                        "PartName": "/word/comments.xml",
                        "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
                    },
                )
            )
        zout.writestr(
            "[Content_Types].xml",
            etree.tostring(ct_tree, xml_declaration=True, encoding="UTF-8"),
        )

    out.seek(0)
    # 把期望值塞到返回里（测试里会重新计算/断言）
    return out.getvalue()


def test_existing_multiline_comment_preserved_after_second_comment():
    docx_bytes = _make_docx_with_existing_multiline_comment()

    # 先确认原始 comments.xml 的全文确实是两段（避免测试构造错误）
    with zipfile.ZipFile(BytesIO(docx_bytes)) as z:
        comments_tree = etree.fromstring(z.read("word/comments.xml"))
        c0 = comments_tree.find("./w:comment[@w:id='0']", NS)
        assert c0 is not None
        assert _extract_comment_text(c0) == "第一段批注\n第二段批注"

    # 用 docxnote 二次批注并渲染
    dn = DocxDocument.parse(docx_bytes, keep_comments=True)
    p = next(b for b in dn.blocks() if isinstance(b, Paragraph) and b.text)
    p.comment("新增批注", start=0, end=5, author="new")
    out = dn.render()

    # 验证：原 comment id=0 的文本仍完整保留，且新增批注存在
    with zipfile.ZipFile(BytesIO(out)) as z:
        comments_tree = etree.fromstring(z.read("word/comments.xml"))

        c0 = comments_tree.find("./w:comment[@w:id='0']", NS)
        assert c0 is not None
        assert _extract_comment_text(c0) == "第一段批注\n第二段批注"

        # 新批注应当生成一个新的 id（1 或更大）
        ids = {
            c.get(f"{{{NS['w']}}}id") for c in comments_tree.findall("./w:comment", NS)
        }
        assert "0" in ids
        assert any(i != "0" for i in ids)


def test_strip_existing_comments_removes_stale_package_references():
    docx_bytes = _make_docx_with_existing_multiline_comment()

    dn = DocxDocument.parse(docx_bytes, keep_comments=False)
    out = dn.render()

    with zipfile.ZipFile(BytesIO(out)) as z:
        assert "word/comments.xml" not in z.namelist()

        rels_tree = etree.fromstring(z.read("word/_rels/document.xml.rels"))
        assert not any(
            rel.get("Type")
            == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
            for rel in rels_tree
        )

        ct_tree = etree.fromstring(z.read("[Content_Types].xml"))
        assert not any(
            override.get("PartName") == "/word/comments.xml" for override in ct_tree
        )


def test_keep_comments_preserves_existing_comment_xml_metadata():
    docx_bytes = _make_docx_with_existing_multiline_comment()

    out = BytesIO()
    with zipfile.ZipFile(BytesIO(docx_bytes), "r") as zin, zipfile.ZipFile(
        out, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for name in zin.namelist():
            if name != "word/comments.xml":
                zout.writestr(name, zin.read(name))

        comments_tree = etree.fromstring(zin.read("word/comments.xml"))
        c0 = comments_tree.find("./w:comment[@w:id='0']", NS)
        assert c0 is not None
        c0.set(f"{{{NS['w']}}}initials", "ZZ")
        c0.set("{http://schemas.microsoft.com/office/word/2012/wordml}done", "1")

        first_run = c0.find("./w:p/w:r", NS)
        assert first_run is not None
        rpr = etree.Element(f"{{{NS['w']}}}rPr")
        etree.SubElement(rpr, f"{{{NS['w']}}}b")
        first_run.insert(0, rpr)

        zout.writestr(
            "word/comments.xml",
            etree.tostring(
                comments_tree, xml_declaration=True, encoding="UTF-8", standalone=True
            ),
        )

    dn = DocxDocument.parse(out.getvalue(), keep_comments=True)
    rendered = dn.render()

    with zipfile.ZipFile(BytesIO(rendered)) as z:
        comments_tree = etree.fromstring(z.read("word/comments.xml"))
        c0 = comments_tree.find("./w:comment[@w:id='0']", NS)
        assert c0 is not None
        assert c0.get(f"{{{NS['w']}}}initials") == "ZZ"
        assert c0.get("{http://schemas.microsoft.com/office/word/2012/wordml}done") == "1"
        assert c0.find(".//w:rPr/w:b", NS) is not None
