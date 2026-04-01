"""段落处理"""

from datetime import datetime
from typing import List, Tuple

from lxml import etree

from .namespaces import NS
from .comments import Comment


class Paragraph:
    """表示 Word 段落"""

    def __init__(self, element, document):
        self._element = element
        self._document = document
        self._text_cache = None
        self._comments_cache: List[Comment] | None = None

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
    ):
        """为段落文本范围添加批注。

        Args:
            date: 批注时间；默认 ``None`` 表示使用当前系统时间（带时区）。
        """
        with self._document._lock:
            if end is None:
                end = len(self.text)

            # 获取批注 ID
            comment_id = self._document.add_comment(text, author, date=date)

            # 在段落中插入批注标记
            self._insert_comment_markers(comment_id, start, end)

            # 新增批注后，清空本段落的批注缓存
            self._comments_cache = None

    def _insert_comment_markers(self, comment_id: int, start: int, end: int):
        """在指定位置插入批注起止标记"""
        runs = list(self._element.findall(".//w:r", NS))
        if not runs:
            return

        # 计算字符位置到 run 的映射
        run_positions = []
        current_pos = 0

        for run in runs:
            run_start = current_pos
            run_text = ""
            for t in run.findall(".//w:t", NS):
                if t.text:
                    run_text += t.text
            run_end = current_pos + len(run_text)
            run_positions.append((run, run_start, run_end, run_text))
            current_pos = run_end

        # 找到需要分割的 run
        start_run_idx = None
        end_run_idx = None

        for idx, (run, run_start, run_end, run_text) in enumerate(run_positions):
            if start_run_idx is None and run_start <= start < run_end:
                start_run_idx = idx
            if end_run_idx is None and run_start < end <= run_end:
                end_run_idx = idx

        if start_run_idx is None or end_run_idx is None:
            return

        # 分割 run 并插入标记
        self._split_and_mark(
            run_positions, start_run_idx, end_run_idx, start, end, comment_id
        )

    def _split_and_mark(
        self, run_positions, start_idx, end_idx, start, end, comment_id
    ):
        """分割 run 并插入批注标记"""
        # 简化实现：在第一个 run 前插入开始标记，在最后一个 run 后插入结束标记
        start_run, start_pos, _, _ = run_positions[start_idx]
        end_run, _, end_pos, _ = run_positions[end_idx]

        # 创建批注范围开始标记
        comment_start = etree.Element(
            f"{{{NS['w']}}}commentRangeStart",
            attrib={f"{{{NS['w']}}}id": str(comment_id)},
        )

        # 创建批注范围结束标记
        comment_end = etree.Element(
            f"{{{NS['w']}}}commentRangeEnd",
            attrib={f"{{{NS['w']}}}id": str(comment_id)},
        )

        # 创建批注引用
        comment_ref_run = etree.Element(f"{{{NS['w']}}}r")
        etree.SubElement(
            comment_ref_run,
            f"{{{NS['w']}}}commentReference",
            attrib={f"{{{NS['w']}}}id": str(comment_id)},
        )

        # 插入标记
        parent = self._element

        # 查找 run 在父元素中的位置
        try:
            children = list(parent)
            start_run_pos = children.index(start_run)
            end_run_pos = children.index(end_run)
        except ValueError:
            # run 不是直接子元素，跳过
            return

        # 在开始 run 之前插入开始标记
        parent.insert(start_run_pos, comment_start)

        # 在结束 run 之后插入结束标记和引用（注意索引偏移）
        parent.insert(end_run_pos + 2, comment_end)
        parent.insert(end_run_pos + 3, comment_ref_run)

    def _iter_comment_ranges(self) -> List[Tuple[int, int, int]]:
        """按段落文本坐标返回本段落上的批注范围列表。

        返回列表元素为 ``(comment_id, start, end)``，其中 start/end 基于
        ``paragraph.text`` 的字符索引区间 [start, end)。
        """
        ranges: list[tuple[int, int, int]] = []
        current_pos = 0
        open_starts: dict[int, int] = {}

        for child in self._element:
            tag = etree.QName(child.tag).localname

            if tag == "commentRangeStart":
                cid_attr = child.get(f"{{{NS['w']}}}id")
                if cid_attr is None:
                    continue
                try:
                    cid = int(cid_attr)
                except ValueError:
                    continue
                # 如果同一 ID 已有开始位置，则不覆盖（保留最早的）
                open_starts.setdefault(cid, current_pos)
            elif tag == "commentRangeEnd":
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
            elif tag == "r":
                # run 内文本长度（与 Paragraph.text 计算方式保持一致）
                run_len = 0
                for r_child in child:
                    r_tag = etree.QName(r_child.tag).localname
                    if r_tag == "t":
                        if r_child.text:
                            run_len += len(r_child.text)
                    elif r_tag == "br":
                        run_len += 1
                    elif r_tag == "tab":
                        run_len += 1
                current_pos += run_len

        return ranges

    @property
    def comments(self) -> tuple[Comment, ...]:
        """返回附着在本段落上的所有批注（按文档顺序）。"""
        with self._document._lock:
            if self._comments_cache is not None:
                return tuple(self._comments_cache)

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
                        start=safe_start,
                        end=safe_end,
                        text=text,
                        author=author,
                        date=date_val,
                    )
                )

            self._comments_cache = result
            return tuple(result)
