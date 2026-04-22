"""可寻址单元：文档内的路径地址工具。

路径（Path）是一段形如 ``"t:0/r:1/c:2/p:0"`` 的字符串，用来在文档中唯一
定位一个段落 / 表格 / 单元格。每一段由 ``kind:index`` 组成：

- ``p:N``  顶层或单元格内第 N 个段落
- ``t:N``  顶层或单元格内第 N 个表格
- ``r:R``  表格的行号（与 ``c:C`` 成对，用来指向单元格原点坐标）
- ``c:C``  表格的列号

以 ``#<comment_id>`` 结尾时表示该段落上的某条批注，
例如 ``"p:0#3"`` 或 ``"t:0/r:0/c:0/p:1#5"``。

设计目标：
- 字符串形式便于 CLI / JSON / LLM 场景直接复用
- 每一条路径都可以通过 :meth:`DocxDocument.resolve` 回溯到对应对象
"""

from __future__ import annotations

from typing import List, Tuple

# 路径分隔符与批注后缀
SEP = "/"
COMMENT_SEP = "#"

# 合法的 segment kind
_SEGMENT_KINDS = ("p", "t", "r", "c")


def build_segment(kind: str, index: int) -> str:
    """构造单个路径段，如 ``build_segment("p", 3) == "p:3"``。"""
    if kind not in _SEGMENT_KINDS:
        raise ValueError(f"invalid segment kind: {kind!r}")
    if index < 0:
        raise ValueError(f"segment index must be >= 0, got {index}")
    return f"{kind}:{index}"


def join_path(*parts: str) -> str:
    """按 ``/`` 连接若干路径段或子路径，忽略空串。"""
    return SEP.join(p for p in parts if p)


def comment_path(paragraph_path: str, comment_id: int) -> str:
    """基于段落路径构造批注路径，如 ``"p:0#3"``。"""
    if comment_id < 0:
        raise ValueError(f"comment_id must be >= 0, got {comment_id}")
    return f"{paragraph_path}{COMMENT_SEP}{comment_id}"


def parse_path(path: str) -> Tuple[List[Tuple[str, int]], int | None]:
    """解析路径字符串。

    返回 ``(segments, comment_id)``：

    - ``segments`` 是 ``[(kind, index), ...]``
    - ``comment_id`` 若路径以 ``#N`` 结尾则为整数，否则为 ``None``
    """
    if not isinstance(path, str):
        raise TypeError("path must be a str")

    raw = path.strip()
    if not raw:
        raise ValueError("path is empty")

    comment_id: int | None = None
    if COMMENT_SEP in raw:
        base, cid_str = raw.split(COMMENT_SEP, 1)
        cid_str = cid_str.strip()
        if not cid_str:
            raise ValueError(f"missing comment id after '#': {path!r}")
        try:
            comment_id = int(cid_str)
        except ValueError as e:
            raise ValueError(f"invalid comment id: {cid_str!r}") from e
        if comment_id < 0:
            raise ValueError(f"comment_id must be >= 0, got {comment_id}")
        raw = base

    segments: list[tuple[str, int]] = []
    for part in raw.split(SEP):
        part = part.strip()
        if not part:
            continue
        if ":" not in part:
            raise ValueError(f"invalid path segment: {part!r}")
        kind, idx_str = part.split(":", 1)
        kind = kind.strip()
        idx_str = idx_str.strip()
        if kind not in _SEGMENT_KINDS:
            raise ValueError(f"invalid segment kind: {kind!r}")
        try:
            idx = int(idx_str)
        except ValueError as e:
            raise ValueError(f"invalid segment index: {idx_str!r}") from e
        if idx < 0:
            raise ValueError(f"segment index must be >= 0, got {idx}")
        segments.append((kind, idx))

    if not segments:
        raise ValueError(f"path has no segments: {path!r}")

    return segments, comment_id


__all__ = [
    "SEP",
    "COMMENT_SEP",
    "build_segment",
    "join_path",
    "comment_path",
    "parse_path",
]
