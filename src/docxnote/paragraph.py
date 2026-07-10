"""段落处理"""

from copy import deepcopy
from datetime import datetime
from typing import List, Tuple

from lxml import etree

from .namespaces import NS
from .comments import Comment
from .paths import comment_path


class Paragraph:
    """表示 Word 段落"""

    def __init__(self, element, document, path: str = ""):
        self._element = element
        self._document = document
        self._path = path
        self._text_cache = None

    @property
    def path(self) -> str:
        """段落的可寻址路径（例如 ``"p:0"`` 或 ``"t:0/r:1/c:2/p:0"``）。"""
        return self._path

    @property
    def text(self) -> str:
        """返回段落完整文本"""
        with self._document._lock:
            if self._text_cache is not None:
                return self._text_cache

            text_parts = []
            for run in self._element.findall(".//w:r", NS):
                # 遍历 run 的所有子元素，保持顺序
                for child in run:
                    tag = etree.QName(child.tag).localname
                    if tag == "t":
                        # 文本节点
                        if child.text:
                            text_parts.append(child.text)
                    elif tag == "br":
                        # 换行符
                        text_parts.append("\n")
                    elif tag == "tab":
                        # 制表符
                        text_parts.append("\t")

            self._text_cache = "".join(text_parts)
            return self._text_cache

    def comment(
        self,
        text: str,
        start: int = 0,
        end: int | None = None,
        *,
        author: str = "docxnote",
        date: datetime | None = None,
    ) -> Comment:
        """为段落文本范围添加批注，并返回对应的 :class:`Comment` 对象。

        Args:
            date: 批注时间；默认 ``None`` 表示使用当前系统时间（带时区）。

        Returns:
            新增批注对应的 :class:`Comment`，其 ``path`` 为
            ``"<paragraph.path>#<comment_id>"``。
        """
        with self._document._lock:
            para_len = len(self.text)
            safe_start, safe_end, _step = slice(start, end).indices(para_len)
            if safe_end < safe_start:
                safe_end = safe_start

            runs = list(self._element.findall(".//w:r", NS))
            if not runs:
                raise ValueError(
                    f"Cannot add comment to paragraph '{self._path}': "
                    "paragraph has no text runs"
                )

            comment_id = self._document.add_comment(text, author, date=date)
            self._insert_comment_markers(comment_id, safe_start, safe_end)

            meta = self._document._get_comment_meta(comment_id)
            if meta is None:
                stored_text, stored_author, stored_date = (
                    text,
                    author,
                    (date if date is not None else datetime.now().astimezone()),
                )
            else:
                stored_text, stored_author, stored_date = meta

            return Comment(
                paragraph=self,
                path=comment_path(self._path, comment_id),
                start=safe_start,
                end=safe_end,
                text=stored_text,
                author=stored_author,
                date=stored_date,
            )

    def _insert_comment_markers(self, comment_id: int, start: int, end: int):
        """在指定位置插入批注起止标记"""
        runs = list(self._element.findall(".//w:r", NS))
        if not runs:
            return

        # 先按结束位置、再按开始位置拆分，避免前一次拆分影响后续边界。
        self._split_run_at_boundary(end)
        self._split_run_at_boundary(start)

        comment_start = etree.Element(
            f"{{{NS['w']}}}commentRangeStart",
            attrib={f"{{{NS['w']}}}id": str(comment_id)},
        )
        comment_end = etree.Element(
            f"{{{NS['w']}}}commentRangeEnd",
            attrib={f"{{{NS['w']}}}id": str(comment_id)},
        )
        comment_ref_run = etree.Element(f"{{{NS['w']}}}r")
        etree.SubElement(
            comment_ref_run,
            f"{{{NS['w']}}}commentReference",
            attrib={f"{{{NS['w']}}}id": str(comment_id)},
        )

        start_parent, start_insert_pos = self._boundary_insert_location(
            self._element, start
        )
        start_parent.insert(start_insert_pos, comment_start)

        end_parent, end_insert_pos = self._boundary_insert_location(self._element, end)
        end_parent.insert(end_insert_pos, comment_end)
        end_parent.insert(end_insert_pos + 1, comment_ref_run)

    def _split_run_at_boundary(self, boundary: int):
        """在字符边界处拆分一个直接子 run。"""
        run_positions = self._run_positions(self._element)
        for run, parent, _idx, run_start, run_end in run_positions:
            if run_start < boundary < run_end:
                self._split_run(parent, run, boundary - run_start)
                return

    def _run_positions(
        self, container
    ) -> list[tuple[etree._Element, etree._Element, int, int, int]]:
        """返回容器内所有 run 的父节点、索引与字符范围。"""
        result: list[tuple[etree._Element, etree._Element, int, int, int]] = []
        current_pos = 0
        for run, parent, idx in self._iter_runs(container):
            run_len = self._run_text_length(run)
            result.append((run, parent, idx, current_pos, current_pos + run_len))
            current_pos += run_len
        return result

    def _iter_runs(self, container):
        """按文档顺序遍历容器内所有 run。"""
        for idx, child in enumerate(container):
            if etree.QName(child.tag).localname == "r":
                yield child, container, idx
            else:
                yield from self._iter_runs(child)

    def _node_text_length(self, node) -> int:
        """计算节点在段落文本视图中的字符长度。"""
        tag = etree.QName(node.tag).localname
        if tag == "r":
            return self._run_text_length(node)

        total = 0
        for child in node:
            total += self._node_text_length(child)
        return total

    def _boundary_insert_location(self, container, boundary: int, current_pos: int = 0):
        """返回指定字符边界应插入到的父节点和索引。"""
        for idx, child in enumerate(container):
            tag = etree.QName(child.tag).localname
            if tag in {"commentRangeStart", "commentRangeEnd"}:
                continue

            child_len = self._node_text_length(child)
            if child_len == 0:
                continue

            if tag == "r":
                if boundary <= current_pos:
                    return container, idx
                current_pos += child_len
                continue

            next_pos = current_pos + child_len
            if current_pos <= boundary <= next_pos:
                return self._boundary_insert_location(child, boundary, current_pos)
            current_pos = next_pos

        return container, len(container)

    def _run_text_length(self, run) -> int:
        """计算单个 run 在文本视图中的字符长度。"""
        run_len = 0
        for child in run:
            tag = etree.QName(child.tag).localname
            if tag == "t":
                run_len += len(child.text or "")
            elif tag in {"br", "tab"}:
                run_len += 1
        return run_len

    def _split_run(self, parent, run, split_offset: int):
        """按字符偏移拆分单个 run。"""
        run_len = self._run_text_length(run)
        if split_offset <= 0 or split_offset >= run_len:
            return

        before_run = etree.Element(run.tag, attrib=dict(run.attrib), nsmap=run.nsmap)
        after_run = etree.Element(run.tag, attrib=dict(run.attrib), nsmap=run.nsmap)
        current_pos = 0

        for child in run:
            tag = etree.QName(child.tag).localname

            if tag == "rPr":
                before_run.append(deepcopy(child))
                after_run.append(deepcopy(child))
                continue

            if tag == "t":
                text = child.text or ""
                next_pos = current_pos + len(text)

                if split_offset > current_pos:
                    before_text = text[: max(0, split_offset - current_pos)]
                    if before_text:
                        before_child = deepcopy(child)
                        before_child.text = before_text
                        before_run.append(before_child)

                if split_offset < next_pos:
                    after_text = text[max(0, split_offset - current_pos) :]
                    if after_text:
                        after_child = deepcopy(child)
                        after_child.text = after_text
                        after_run.append(after_child)

                current_pos = next_pos
                continue

            if tag in {"br", "tab"}:
                if current_pos < split_offset:
                    before_run.append(deepcopy(child))
                else:
                    after_run.append(deepcopy(child))
                current_pos += 1
                continue

            # Preserve non-text run content such as drawings, field markers,
            # symbols, and references instead of silently dropping it.
            if current_pos < split_offset:
                before_run.append(deepcopy(child))
            else:
                after_run.append(deepcopy(child))

        children = list(parent)
        run_idx = children.index(run)
        parent.remove(run)
        parent.insert(run_idx, before_run)
        parent.insert(run_idx + 1, after_run)

    def _iter_comment_ranges(self) -> List[Tuple[int, int, int]]:
        """按段落文本坐标返回本段落上的批注范围列表。

        返回列表元素为 ``(comment_id, start, end)``，其中 start/end 基于
        ``paragraph.text`` 的字符索引区间 [start, end)。
        """
        ranges: list[tuple[int, int, int]] = []
        open_starts: dict[int, int] = {}

        self._walk_comment_ranges(self._element, 0, open_starts, ranges)

        return ranges

    def _walk_comment_ranges(self, container, current_pos, open_starts, ranges):
        """递归遍历段落内容，收集批注范围。"""
        for child in container:
            tag = etree.QName(child.tag).localname

            if tag == "commentRangeStart":
                cid_attr = child.get(f"{{{NS['w']}}}id")
                if cid_attr is None:
                    continue
                try:
                    cid = int(cid_attr)
                except ValueError:
                    continue
                open_starts.setdefault(cid, current_pos)
                continue

            if tag == "commentRangeEnd":
                cid_attr = child.get(f"{{{NS['w']}}}id")
                if cid_attr is None:
                    continue
                try:
                    cid = int(cid_attr)
                except ValueError:
                    continue
                start_pos = open_starts.pop(cid, current_pos)
                if start_pos <= current_pos:
                    ranges.append((cid, start_pos, current_pos))
                continue

            if tag == "r":
                current_pos += self._run_text_length(child)
                continue

            current_pos = self._walk_comment_ranges(
                child, current_pos, open_starts, ranges
            )

        return current_pos

    @property
    def comments(self) -> tuple[Comment, ...]:
        """返回附着在本段落上的所有批注（按文档顺序）。"""
        with self._document._lock:
            ranges = self._iter_comment_ranges()
            result: list[Comment] = []

            for comment_id, start, end in ranges:
                meta = self._document._get_comment_meta(comment_id)
                if meta is None:
                    continue
                text, author, date_val = meta

                # 限制范围在合法区间内，避免异常 XML 造成越界
                para_len = len(self.text)
                safe_start = max(0, min(start, para_len))
                safe_end = max(safe_start, min(end, para_len))

                result.append(
                    Comment(
                        paragraph=self,
                        path=comment_path(self._path, comment_id),
                        start=safe_start,
                        end=safe_end,
                        text=text,
                        author=author,
                        date=date_val,
                    )
                )

            return tuple(result)
