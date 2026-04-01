"""公共批注对象模型（只暴露高层视图）"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .paragraph import Paragraph


@dataclass(frozen=True)
class Comment:
    """表示附着在段落上的批注视图。

    - `paragraph`: 所属的段落对象
    - `start` / `end`: 在 `paragraph.text` 上的字符区间 [start, end)
    - `text`: 批注内容（纯文本，多段用换行符分隔）
    - `author`: 批注作者
    - `date`: 批注时间（时区感知）
    """

    paragraph: "Paragraph"
    start: int
    end: int
    text: str
    author: str
    date: datetime

