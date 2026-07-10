"""DOCX 文档解析和渲染"""

import io
import threading
import zipfile
from copy import deepcopy
from datetime import datetime, timezone
from typing import Optional, Tuple

from lxml import etree

from .paragraph import Paragraph
from .table import Table, Cell
from .namespaces import NS
from .comments import Comment
from .paths import build_segment, comment_path, parse_path


def _parse_w_comment_date(value: str | None) -> datetime:
    """Parse w:date from comments.xml (ISO 8601, often UTC with Z)."""
    if not value or not str(value).strip():
        return datetime.now(timezone.utc)
    v = str(value).strip()
    if v.endswith("Z"):
        v = v[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(v)
    except ValueError:
        return datetime.now(timezone.utc)


def _format_w_comment_date(dt: datetime) -> str:
    """Serialize datetime to w:date (UTC, ``...Z`` suffix)."""
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    utc = dt.astimezone(timezone.utc)
    return utc.strftime("%Y-%m-%dT%H:%M:%SZ")


def _default_new_comment_date() -> datetime:
    """Default comment time: current system time (timezone-aware local)."""
    return datetime.now().astimezone()


class DocxDocument:
    """DOCX 文档对象

    同一 ``DocxDocument`` 实例可在多线程环境下安全使用（内部使用可重入锁串行化访问）。
    不同实例之间无共享可变状态，可并行使用。多进程请各自持有独立实例。
    """

    def __init__(self, zip_data: bytes):
        self._zip_data = zip_data
        self._zip = zipfile.ZipFile(io.BytesIO(zip_data))
        self._document_xml = None
        self._body = None
        self._comments: list[tuple[int, str, str, datetime]] = []
        self._comment_index: dict[int, tuple[str, str, datetime]] = {}
        self._existing_comment_elements: dict[int, etree._Element] = {}
        self._comments_root_template: etree._Element | None = None
        self._comment_id_counter = 0
        self._lock = threading.RLock()

    @classmethod
    def parse(cls, docx_bytes: bytes, *, keep_comments: bool = False) -> "DocxDocument":
        """解析 DOCX 并构建文档对象

        Args:
            keep_comments: 是否保留原有批注。默认 False（清空所有原有批注）。
        """
        doc = cls(docx_bytes)
        doc._load_document(keep_comments=keep_comments)
        return doc

    def _load_document(self, *, keep_comments: bool):
        """加载 document.xml，并按需保留/清空原有批注"""
        doc_xml = self._zip.read("word/document.xml")
        self._document_xml = etree.fromstring(doc_xml)
        self._body = self._document_xml.find(".//w:body", NS)

        if keep_comments:
            # 加载已有的批注
            self._load_existing_comments()
        else:
            # 默认不保留：清空 comments 列表，并移除 document.xml 中的批注标记
            self._comments = []
            self._comment_index = {}
            self._comment_id_counter = 0
            self._strip_all_comment_markers()

    def _strip_all_comment_markers(self) -> None:
        """移除 document.xml 中所有批注相关标记，避免残留引用。"""
        if self._document_xml is None:
            return

        # commentRangeStart / commentRangeEnd
        for tag in ("commentRangeStart", "commentRangeEnd"):
            for el in self._document_xml.findall(f".//w:{tag}", NS):
                parent = el.getparent()
                if parent is not None:
                    parent.remove(el)

        # commentReference 位于 w:r 内；移除后若 run 为空则一并移除
        for ref in self._document_xml.findall(".//w:commentReference", NS):
            run = ref.getparent()
            if run is None:
                continue
            run.remove(ref)
            if (
                len(run) == 0
                and (run.text is None)
                and (run.tail is None or run.tail == "")
            ):
                parent = run.getparent()
                if parent is not None:
                    parent.remove(run)

    def _load_existing_comments(self):
        """加载已有的批注"""
        try:
            comments_xml = self._zip.read("word/comments.xml")
            comments_tree = etree.fromstring(comments_xml)

            max_id = -1
            self._comments.clear()
            self._comment_index.clear()
            self._existing_comment_elements.clear()
            self._comments_root_template = deepcopy(comments_tree)

            for comment in comments_tree:
                comment_id_str = comment.get(f"{{{NS['w']}}}id")
                if not comment_id_str:
                    continue
                try:
                    comment_id = int(comment_id_str)
                except ValueError:
                    continue
                max_id = max(max_id, comment_id)

                # 提取批注内容
                author = comment.get(f"{{{NS['w']}}}author", "") or ""
                text = self._extract_comment_text(comment)
                date_str = comment.get(f"{{{NS['w']}}}date")
                date_val = _parse_w_comment_date(date_str)

                meta = (text, author, date_val)
                self._comments.append((comment_id, *meta))
                self._comment_index[comment_id] = meta
                self._existing_comment_elements[comment_id] = deepcopy(comment)

            # 设置下一个批注 ID
            self._comment_id_counter = max_id + 1
        except KeyError:
            # 没有 comments.xml 文件
            pass

    def _extract_comment_text(self, comment_element: etree._Element) -> str:
        """从 w:comment 中提取完整文本（按 w:p 插入换行）。"""
        parts: list[str] = []
        first_para = True

        # comments.xml 内部结构通常是多个 w:p
        for p in comment_element.findall(".//w:p", NS):
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

    def blocks(self) -> tuple[Paragraph | Table, ...]:
        """返回文档中的块级元素（元组），每个元素都带有 ``path``。"""
        with self._lock:
            if self._body is None:
                return ()

            blocks: list[Paragraph | Table] = []
            para_idx = 0
            table_idx = 0
            for child in self._body:
                tag = etree.QName(child.tag).localname
                if tag == "p":
                    blocks.append(
                        Paragraph(child, self, path=build_segment("p", para_idx))
                    )
                    para_idx += 1
                elif tag == "tbl":
                    blocks.append(
                        Table(child, self, path=build_segment("t", table_idx))
                    )
                    table_idx += 1
            return tuple(blocks)

    def iter_paragraphs(self):
        """按文档顺序迭代所有段落（包含表格与嵌套表格中的段落）。

        每个段落对象都带有可寻址路径 ``paragraph.path``。
        """
        with self._lock:
            paragraphs = tuple(_walk_paragraphs(self.blocks()))

        yield from paragraphs

    def add_comment(
        self,
        text: str,
        author: str = "docxnote",
        *,
        date: datetime | None = None,
    ) -> int:
        """添加批注并返回 ID。

        Args:
            date: 批注时间；默认 ``None`` 表示使用当前系统时间（带时区）。
        """
        with self._lock:
            when = date if date is not None else _default_new_comment_date()
            comment_id = self._comment_id_counter
            self._comment_id_counter += 1

            meta = (text, author, when)
            self._comments.append((comment_id, *meta))
            self._comment_index[comment_id] = meta
            return comment_id

    def _get_comment_meta(self, comment_id: int) -> Optional[Tuple[str, str, datetime]]:
        """根据 comment_id 返回批注的 (text, author, date)。"""
        return self._comment_index.get(comment_id)

    def render(self) -> bytes:
        """生成新的 DOCX 并返回 bytes"""
        with self._lock:
            return self._render_unlocked()

    def _render_unlocked(self) -> bytes:
        output = io.BytesIO()

        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as out_zip:
            include_comments = bool(self._comments)
            rels_data = self._prepare_rels(include_comments=include_comments)
            content_types_data = self._prepare_content_types(
                include_comments=include_comments
            )

            # 复制所有原始文件
            for item in self._zip.namelist():
                if item == "word/document.xml":
                    continue
                if item == "word/comments.xml":
                    continue
                if item == "word/_rels/document.xml.rels" and rels_data is not None:
                    continue
                if item == "[Content_Types].xml" and content_types_data is not None:
                    continue
                out_zip.writestr(item, self._zip.read(item))

            # 写入修改后的 document.xml
            doc_bytes = etree.tostring(
                self._document_xml,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            out_zip.writestr("word/document.xml", doc_bytes)

            # 写入 comments.xml、rels 和 content types
            if self._comments:
                comments_xml = self._build_comments_xml()
                out_zip.writestr("word/comments.xml", comments_xml)

            if rels_data is not None:
                out_zip.writestr("word/_rels/document.xml.rels", rels_data)

            if content_types_data is not None:
                out_zip.writestr("[Content_Types].xml", content_types_data)

        return output.getvalue()

    def _build_comments_xml(self) -> bytes:
        """构建 comments.xml"""
        if self._comments_root_template is not None:
            root = deepcopy(self._comments_root_template)
            for child in list(root):
                root.remove(child)
        else:
            root = etree.Element(f"{{{NS['w']}}}comments", nsmap=NS)

        for comment_id, text, author, date_val in self._comments:
            existing = self._existing_comment_elements.get(comment_id)
            if existing is not None:
                root.append(deepcopy(existing))
                continue

            root.append(
                self._build_new_comment_element(comment_id, text, author, date_val)
            )

        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    def _build_new_comment_element(
        self, comment_id: int, text: str, author: str, date_val: datetime
    ) -> etree._Element:
        comment = etree.Element(
            f"{{{NS['w']}}}comment",
            attrib={
                f"{{{NS['w']}}}id": str(comment_id),
                f"{{{NS['w']}}}author": author,
                f"{{{NS['w']}}}date": _format_w_comment_date(date_val),
                f"{{{NS['w']}}}initials": author[0].upper() if author else "D",
            },
        )

        lines = text.split("\n")
        if not lines:
            lines = [""]

        for line in lines:
            p = etree.SubElement(comment, f"{{{NS['w']}}}p")
            r = etree.SubElement(p, f"{{{NS['w']}}}r")

            if "\t" in line:
                buf: list[str] = []
                for ch in line:
                    if ch == "\t":
                        if buf:
                            t = etree.SubElement(r, f"{{{NS['w']}}}t")
                            seg = "".join(buf)
                            if seg[:1] == " " or seg[-1:] == " ":
                                t.set(
                                    "{http://www.w3.org/XML/1998/namespace}space",
                                    "preserve",
                                )
                            t.text = seg
                            buf.clear()
                        etree.SubElement(r, f"{{{NS['w']}}}tab")
                    else:
                        buf.append(ch)
                if buf or line == "":
                    t = etree.SubElement(r, f"{{{NS['w']}}}t")
                    seg = "".join(buf)
                    if seg[:1] == " " or seg[-1:] == " ":
                        t.set(
                            "{http://www.w3.org/XML/1998/namespace}space",
                            "preserve",
                        )
                    t.text = seg
            else:
                t = etree.SubElement(r, f"{{{NS['w']}}}t")
                if line[:1] == " " or line[-1:] == " ":
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t.text = line

        return comment

    def _prepare_rels(self, *, include_comments: bool) -> bytes | None:
        """准备 document.xml.rels 数据以包含 comments.xml 关系"""
        rels_path = "word/_rels/document.xml.rels"

        try:
            rels_data = self._zip.read(rels_path)
            rels_xml = etree.fromstring(rels_data)
        except KeyError:
            if not include_comments:
                return None
            rels_xml = etree.Element(
                "Relationships",
                nsmap={
                    None: "http://schemas.openxmlformats.org/package/2006/relationships"
                },
            )

        comment_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        for rel in list(rels_xml):
            if rel.get("Type") == comment_type:
                rels_xml.remove(rel)

        if include_comments:
            # 添加 comments 关系
            max_id = 0
            for rel in rels_xml:
                rel_id = rel.get("Id", "")
                if rel_id.startswith("rId"):
                    try:
                        num = int(rel_id[3:])
                        max_id = max(max_id, num)
                    except ValueError:
                        pass

            etree.SubElement(
                rels_xml,
                "Relationship",
                attrib={
                    "Id": f"rId{max_id + 1}",
                    "Type": comment_type,
                    "Target": "comments.xml",
                },
            )

        return etree.tostring(rels_xml, xml_declaration=True, encoding="UTF-8")

    def _prepare_content_types(self, *, include_comments: bool) -> bytes:
        """准备 [Content_Types].xml 数据以包含 comments.xml"""
        ct_data = self._zip.read("[Content_Types].xml")
        ct_xml = etree.fromstring(ct_data)

        # 获取命名空间
        ns = ct_xml.nsmap.get(
            None, "http://schemas.openxmlformats.org/package/2006/content-types"
        )

        for override in list(ct_xml):
            if override.get("PartName") == "/word/comments.xml":
                ct_xml.remove(override)

        if include_comments:
            # 添加 comments.xml 的 Override
            override_elem = etree.Element(
                f"{{{ns}}}Override",
                attrib={
                    "PartName": "/word/comments.xml",
                    "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
                },
            )
            ct_xml.append(override_elem)

        return etree.tostring(ct_xml, xml_declaration=True, encoding="UTF-8")

    def comments(self) -> tuple[Comment, ...]:
        """返回文档中所有批注（按文档遍历顺序）。"""
        with self._lock:
            result: list[Comment] = []
            for para in self.iter_paragraphs():
                result.extend(para.comments)
            return tuple(result)

    def resolve(self, path: str) -> Paragraph | Table | Cell | Comment:
        """根据路径字符串定位对应的对象。

        支持的路径形式见 :mod:`docxnote.paths`。例如：

        - ``"p:0"``                     顶层第 0 个段落
        - ``"t:0"``                     顶层第 0 个表格
        - ``"t:0/r:1/c:2"``             表格中的某个单元格
        - ``"t:0/r:1/c:2/p:0"``         单元格内的段落
        - ``"p:0#5"``                   段落 ``p:0`` 上 ``w:id=5`` 的批注

        Raises:
            ValueError: 路径格式非法
            LookupError: 路径结构合法但指向不存在的对象
        """
        segments, comment_id = parse_path(path)

        with self._lock:
            current = self._navigate_segments(segments)

            if comment_id is None:
                return current

            if not isinstance(current, Paragraph):
                raise ValueError(
                    "comment path must target a paragraph, got "
                    f"{type(current).__name__}"
                )
            canonical_path = comment_path(current.path, comment_id)
            for c in current.comments:
                if c.path == canonical_path:
                    return c
            raise LookupError(
                f"comment {comment_id} not found on paragraph {current.path!r}"
            )

    def _navigate_segments(
        self, segments: list[tuple[str, int]]
    ) -> Paragraph | Table | Cell:
        """按 segments 从文档根节点向下定位。"""
        if not segments:
            raise ValueError("empty segments")

        first_kind, first_idx = segments[0]
        blocks = self.blocks()

        if first_kind == "p":
            paras = [b for b in blocks if isinstance(b, Paragraph)]
            if first_idx >= len(paras):
                raise LookupError(f"paragraph index out of range: p:{first_idx}")
            current: Paragraph | Table | Cell = paras[first_idx]
        elif first_kind == "t":
            tables = [b for b in blocks if isinstance(b, Table)]
            if first_idx >= len(tables):
                raise LookupError(f"table index out of range: t:{first_idx}")
            current = tables[first_idx]
        else:
            raise ValueError(f"first segment must be 'p' or 't', got {first_kind!r}")

        i = 1
        while i < len(segments):
            kind, idx = segments[i]

            if isinstance(current, Table):
                if kind != "r":
                    raise ValueError(f"after table expected 'r:', got {kind!r} in path")
                if i + 1 >= len(segments):
                    raise ValueError("path has 'r:' without following 'c:'")
                kind2, idx2 = segments[i + 1]
                if kind2 != "c":
                    raise ValueError(f"after 'r:' expected 'c:', got {kind2!r}")
                rows, cols = current.shape()
                if idx >= rows or idx2 >= cols:
                    raise LookupError(
                        f"cell out of bounds: r:{idx}/c:{idx2} in {current.path!r}"
                    )
                current = current[idx, idx2]
                i += 2
                continue

            if isinstance(current, Cell):
                sub = current.blocks()
                if kind == "p":
                    ps = [b for b in sub if isinstance(b, Paragraph)]
                    if idx >= len(ps):
                        raise LookupError(
                            f"paragraph index out of range: p:{idx} in {current.path!r}"
                        )
                    current = ps[idx]
                elif kind == "t":
                    ts = [b for b in sub if isinstance(b, Table)]
                    if idx >= len(ts):
                        raise LookupError(
                            f"table index out of range: t:{idx} in {current.path!r}"
                        )
                    current = ts[idx]
                else:
                    raise ValueError(f"after cell expected 'p:' or 't:', got {kind!r}")
                i += 1
                continue

            if isinstance(current, Paragraph):
                raise ValueError(
                    "paragraph has no navigable children; use '#<id>' for comments"
                )

            raise ValueError(
                f"unexpected object in navigation: {type(current).__name__}"
            )

        return current


def _walk_paragraphs(blocks, _seen_cells: set[int] | None = None):
    """按文档顺序递归遍历 blocks，yield 每一个 Paragraph。

    对合并单元格按单元格身份去重，避免同一段落被多次 yield。
    """
    if _seen_cells is None:
        _seen_cells = set()

    for block in blocks:
        if isinstance(block, Paragraph):
            yield block
        elif isinstance(block, Table):
            rows, cols = block.shape()
            for r in range(rows):
                for c in range(cols):
                    cell = block[r, c]
                    cell_key = id(cell)
                    if cell_key in _seen_cells:
                        continue
                    _seen_cells.add(cell_key)
                    yield from _walk_paragraphs(cell.blocks(), _seen_cells)
