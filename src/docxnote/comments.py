"""公共批注对象模型（只暴露高层视图）"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .paragraph import Paragraph


class UnsupportedCommentRangeError(ValueError):
    """批注范围超出单段范围视图的支持能力。

    当批注的 ``commentRangeStart`` / ``commentRangeEnd`` 跨越段落边界，
    或范围未闭合（缺少配对标记）时，访问高层范围视图
    （``paragraph.comments``、``doc.comments()``、
    ``doc.resolve("p:0#N")``）会抛出本异常。

    解析（``DocxDocument.parse``）与渲染（``doc.render()``）不受影响：
    这类 XML 会被原样透传。
    """


@dataclass(frozen=True)
class Comment:
    """表示附着在段落上的批注视图。

    Attributes:
        paragraph: 所属的段落对象
        path: 批注的可寻址路径，形如 ``"t:0/r:0/c:0/p:0#3"``；
            ``#`` 后为 Word 内部 ``w:id``，在同一文档中稳定
        start: 在 ``paragraph.text`` 上的字符起点（含）
        end: 在 ``paragraph.text`` 上的字符终点（不含），``[start, end)``
        text: 批注内容（纯文本，多段以换行符分隔）
        author: 批注作者
        date: 批注时间（时区感知）；源 XML 的 ``w:date`` 缺失、空白或
            非法时为 ``None``
    """

    paragraph: "Paragraph"
    path: str
    start: int
    end: int
    text: str
    author: str
    date: datetime | None
