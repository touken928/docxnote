"""Public table and cell views over logical Word table coordinates."""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from ._navigation import collect_blocks
from ._paths import build_segment, join_path
from ._state import DocumentOwner
from ._xml.tables import CellRegion, parse_table_grid

if TYPE_CHECKING:
    from .paragraph import Paragraph


class Table:
    """A table whose merged coordinates refer to their origin Cell."""

    def __init__(
        self, element: etree._Element, document: DocumentOwner, path: str = ""
    ) -> None:
        self._state = document._state
        self._path = path
        with self._state.lock:
            regions = parse_table_grid(element)
            cells: dict[CellRegion, Cell] = {}
            matrix: list[tuple[Cell, ...]] = []
            for row in regions:
                row_cells: list[Cell] = []
                for region in row:
                    if region not in cells:
                        cell = Cell(
                            region.element,
                            self,
                            region.row,
                            region.col,
                            region.colspan,
                            path=join_path(
                                path,
                                build_segment("r", region.row),
                                build_segment("c", region.col),
                            ),
                        )
                        cell._rowspan = region.rowspan
                        cells[region] = cell
                    row_cells.append(cells[region])
                matrix.append(tuple(row_cells))
            self._matrix = tuple(matrix)

    @property
    def path(self) -> str:
        """The address of this table, such as ``t:0/r:1/c:2/t:0``."""
        return self._path

    def shape(self) -> tuple[int, int]:
        """Return the logical (rows, columns), including omitted cell lanes."""
        return len(self._matrix), len(self._matrix[0]) if self._matrix else 0

    def __getitem__(self, key: tuple[int, int]) -> Cell:
        row, col = key
        if not (0 <= row < len(self._matrix) and 0 <= col < len(self._matrix[row])):
            raise IndexError(
                f"cell ({row}, {col}) out of bounds for table {self.shape()}"
            )
        return self._matrix[row][col]


class Cell:
    """A merge-origin cell, or a synthetic empty cell for an omitted coordinate."""

    def __init__(
        self,
        element: etree._Element | None,
        document: DocumentOwner,
        row: int,
        col: int,
        colspan: int = 1,
        path: str = "",
    ) -> None:
        self._element = element
        self._state = document._state
        self._row = row
        self._col = col
        self._colspan = colspan
        self._rowspan = 1
        self._path = path

    @property
    def path(self) -> str:
        """The address of the merge origin, such as ``t:0/r:1/c:2``."""
        return self._path

    def blocks(self) -> tuple[Paragraph | Table, ...]:
        """Return this cell's paragraphs and tables in document order."""
        with self._state.lock:
            return collect_blocks(self._element, self, self._path)

    def bounds(self) -> tuple[int, int, int, int]:
        """Return half-open (top, left, bottom, right) logical coordinates."""
        if self._element is None:
            return self._row, self._col, self._row + 1, self._col + 1
        return (
            self._row,
            self._col,
            self._row + self._rowspan,
            self._col + self._colspan,
        )
