"""表格和单元格处理"""

from __future__ import annotations

from lxml import etree
from .namespaces import NS
from .paths import build_segment, join_path


class Table:
    """表示 Word 表格"""

    def __init__(self, element, document, path: str = ""):
        self._element = element
        self._document = document
        self._path = path
        self._grid = None
        self._build_grid()

    @property
    def path(self) -> str:
        """表格的可寻址路径（例如 ``"t:0"`` 或 ``"t:0/r:1/c:2/t:0"``）。"""
        return self._path

    def _build_grid(self):
        """构建表格网格，处理合并单元格"""
        # 只查找直接子行，不包括嵌套表格的行
        rows = self._element.findall("./w:tr", NS)
        if not rows:
            self._grid = []
            return

        # 构建一个“展开到坐标”的网格：同一合并区域的所有坐标都指向起始 Cell
        self._grid = []
        row_maps: list[dict[int, Cell]] = []
        active_vmerge: dict[
            int, Cell
        ] = {}  # col -> origin cell (for current/next rows)
        max_cols = 0

        for r_idx, row in enumerate(rows):
            # 只查找直接子单元格
            tcs = row.findall("./w:tc", NS)

            row_map: dict[int, Cell] = {}
            col_idx = 0

            for tc in tcs:
                # 跳过被上方 vMerge 占用的列位置
                while col_idx in row_map:
                    col_idx += 1

                colspan = 1
                vmerge_val: str | None = None

                tc_pr = tc.find("./w:tcPr", NS)
                if tc_pr is not None:
                    gridspan = tc_pr.find("./w:gridSpan", NS)
                    if gridspan is not None:
                        val = gridspan.get(f"{{{NS['w']}}}val")
                        if val:
                            colspan = int(val)

                    vmerge = tc_pr.find("./w:vMerge", NS)
                    if vmerge is not None:
                        vmerge_val = vmerge.get(f"{{{NS['w']}}}val")

                is_vmerge_continue = vmerge_val is None and (
                    tc_pr is not None and tc_pr.find("./w:vMerge", NS) is not None
                )
                if vmerge_val is not None:
                    is_vmerge_continue = vmerge_val != "restart"

                if is_vmerge_continue:
                    origin = active_vmerge.get(col_idx)
                    if origin is None:
                        origin = Cell(
                            tc,
                            self._document,
                            r_idx,
                            col_idx,
                            colspan,
                            path=self._cell_path(r_idx, col_idx),
                        )
                    else:
                        origin._grow_rowspan_to(r_idx + 1)
                else:
                    origin = Cell(
                        tc,
                        self._document,
                        r_idx,
                        col_idx,
                        colspan,
                        path=self._cell_path(r_idx, col_idx),
                    )
                    # 新单元格覆盖同列：意味着上方 vMerge 在该列结束
                    for i in range(colspan):
                        active_vmerge.pop(col_idx + i, None)

                    # 如果当前单元格是 vMerge restart，则开启纵向合并跟踪
                    if vmerge_val == "restart" or (
                        tc_pr is not None
                        and tc_pr.find("./w:vMerge", NS) is not None
                        and vmerge_val == "restart"
                    ):
                        for i in range(colspan):
                            active_vmerge[col_idx + i] = origin

                for i in range(colspan):
                    row_map[col_idx + i] = origin

                col_idx += colspan

            # 将本行未显式出现但仍在 vMerge 中的列补齐
            for c, origin in active_vmerge.items():
                if c not in row_map:
                    origin._grow_rowspan_to(r_idx + 1)
                    row_map[c] = origin

            if row_map:
                max_cols = max(max_cols, max(row_map.keys()) + 1)
            row_maps.append(row_map)

        # 生成最终 grid：每行是“实际出现过的 Cell（去重）”列表（用于 bounds/shape 辅助）
        for r_idx, row_map in enumerate(row_maps):
            seen: set[int] = set()
            grid_row: list[Cell] = []
            for c in range(max_cols):
                cell = row_map.get(c)
                if cell is None:
                    continue
                if id(cell) in seen:
                    continue
                seen.add(id(cell))
                grid_row.append(cell)
            self._grid.append(grid_row)

        # 另外保存一个坐标展开矩阵，供 __getitem__ 精确返回合并起点单元格
        self._matrix: list[list[Cell]] = []
        for r_idx, row_map in enumerate(row_maps):
            matrix_row: list[Cell] = []
            for c in range(max_cols):
                matrix_row.append(
                    row_map.get(c)
                    or Cell(
                        None,
                        self._document,
                        r_idx,
                        c,
                        1,
                        path=self._cell_path(r_idx, c),
                    )
                )
            self._matrix.append(matrix_row)

    def _cell_path(self, row: int, col: int) -> str:
        """根据本表格路径计算单元格路径（以单元格原点坐标为准）。"""
        return join_path(self._path, build_segment("r", row), build_segment("c", col))

    def shape(self) -> tuple[int, int]:
        """返回表格尺寸 (rows, cols)"""
        if not self._grid:
            return (0, 0)
        rows = len(self._grid)
        cols = (
            len(getattr(self, "_matrix", [[]])[0])
            if getattr(self, "_matrix", None)
            else 0
        )
        return (rows, cols)

    def __getitem__(self, key: tuple[int, int]) -> "Cell":
        """返回 Cell 对象"""
        row, col = key
        matrix = getattr(self, "_matrix", None)
        if (
            matrix is not None
            and 0 <= row < len(matrix)
            and 0 <= col < len(matrix[row])
        ):
            return matrix[row][col]
        return Cell(None, self._document, row, col, 1)


class Cell:
    """表示表格单元格"""

    def __init__(
        self,
        element,
        document,
        row: int,
        col: int,
        colspan: int = 1,
        path: str = "",
    ):
        self._element = element
        self._document = document
        self._row = row
        self._col = col
        self._colspan = colspan
        self._rowspan = 1
        self._path = path

    @property
    def path(self) -> str:
        """单元格的可寻址路径，基于合并原点坐标（例如 ``"t:0/r:1/c:2"``）。"""
        return self._path

    def _grow_rowspan_to(self, bottom_exclusive: int) -> None:
        """将 rowspan 扩展到指定 bottom（左闭右开）"""
        self._rowspan = max(self._rowspan, bottom_exclusive - self._row)

    def blocks(self) -> tuple:
        """返回单元格中的块级元素（元组）"""
        with self._document._lock:
            if self._element is None:
                return ()

            from .paragraph import Paragraph

            blocks: list = []
            para_idx = 0
            table_idx = 0
            for child in self._element:
                tag = etree.QName(child.tag).localname
                if tag == "p":
                    child_path = join_path(self._path, build_segment("p", para_idx))
                    blocks.append(Paragraph(child, self._document, path=child_path))
                    para_idx += 1
                elif tag == "tbl":
                    child_path = join_path(self._path, build_segment("t", table_idx))
                    blocks.append(Table(child, self._document, path=child_path))
                    table_idx += 1
            return tuple(blocks)

    def bounds(self) -> tuple[int, int, int, int]:
        """返回单元格边界 (top, left, bottom, right)"""
        if self._element is None:
            return (self._row, self._col, self._row + 1, self._col + 1)

        return (
            self._row,
            self._col,
            self._row + self._rowspan,
            self._col + self._colspan,
        )
