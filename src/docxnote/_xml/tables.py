"""Parse logical table coordinates without constructing public Cell objects."""

from dataclasses import dataclass
from typing import Literal

from lxml import etree

from .namespaces import NS, W


@dataclass(frozen=True)
class CellSpec:
    element: etree._Element
    colspan: int
    merge: Literal["restart", "continue"] | None


@dataclass(frozen=True)
class RowSpec:
    before: int
    after: int
    cells: tuple[CellSpec, ...]

    @property
    def width(self) -> int:
        return self.before + sum(cell.colspan for cell in self.cells) + self.after


@dataclass(eq=False)
class CellRegion:
    """One merge origin shared by all covered coordinates during grid parsing."""

    element: etree._Element | None
    row: int
    col: int
    colspan: int = 1
    rowspan: int = 1

    def extend_to(self, bottom: int) -> None:
        self.rowspan = max(self.rowspan, bottom - self.row)


def _integer_property(element: etree._Element | None, path: str, default: int) -> int:
    if element is None:
        return default
    property_element = element.find(path, NS)
    value = property_element.get(W + "val") if property_element is not None else None
    return int(value) if value else default


def _parse_row(row: etree._Element) -> RowSpec:
    properties = row.find("./w:trPr", NS)
    cells: list[CellSpec] = []
    for element in row.findall("./w:tc", NS):
        cell_properties = element.find("./w:tcPr", NS)
        merge_element = (
            cell_properties.find("./w:vMerge", NS)
            if cell_properties is not None
            else None
        )
        merge: Literal["restart", "continue"] | None = None
        if merge_element is not None:
            merge = (
                "restart" if merge_element.get(W + "val") == "restart" else "continue"
            )
        cells.append(
            CellSpec(
                element, _integer_property(cell_properties, "./w:gridSpan", 1), merge
            )
        )
    return RowSpec(
        _integer_property(properties, "./w:gridBefore", 0),
        _integer_property(properties, "./w:gridAfter", 0),
        tuple(cells),
    )


def parse_table_grid(table: etree._Element) -> tuple[tuple[CellRegion, ...], ...]:
    """Return a rectangular matrix, sharing regions across merged coordinates.

    tblGrid defines logical width when present. Missing row coordinates are
    synthetic one-cell regions; merges cannot continue across omitted lanes.
    """
    rows = tuple(_parse_row(row) for row in table.findall("./w:tr", NS))
    if not rows:
        return ()
    grid = table.find("./w:tblGrid", NS)
    width = (
        len(grid.findall("./w:gridCol", NS))
        if grid is not None
        else max(row.width for row in rows)
    )
    active_merges: dict[int, CellRegion] = {}
    row_maps: list[dict[int, CellRegion]] = []
    for row_index, row in enumerate(rows):
        row_end = width - row.after
        active_merges = {
            col: origin
            for col, origin in active_merges.items()
            if row.before <= col < row_end
        }
        row_map: dict[int, CellRegion] = {}
        col = row.before
        for cell in row.cells:
            while col in row_map:
                col += 1
            origin = active_merges.get(col) if cell.merge == "continue" else None
            if origin is None:
                origin = CellRegion(cell.element, row_index, col, cell.colspan)
            else:
                origin.extend_to(row_index + 1)
            if cell.merge != "continue":
                for lane in range(col, col + cell.colspan):
                    active_merges.pop(lane, None)
                    if cell.merge == "restart":
                        active_merges[lane] = origin
            for lane in range(col, col + cell.colspan):
                row_map[lane] = origin
            col += cell.colspan
        # Preserve the existing fallback for a merge lane with no explicit
        # continuation cell, but never fill gridBefore/gridAfter coordinates.
        for col, origin in active_merges.items():
            if row.before <= col < row_end and col not in row_map:
                origin.extend_to(row_index + 1)
                row_map[col] = origin
        if grid is None and row_map:
            width = max(width, max(row_map) + 1)
        row_maps.append(row_map)
    return tuple(
        tuple(
            row_map[col] if col in row_map else CellRegion(None, row_index, col)
            for col in range(width)
        )
        for row_index, row_map in enumerate(row_maps)
    )
