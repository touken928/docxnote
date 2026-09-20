"""Omitted grid lanes and vertical merges retain logical coordinates."""

from docx import Document as PythonDocxDocument
from lxml import etree

from docxnote import Cell, DocxDocument, Paragraph, Table
from tests.support.docx import save_docx, read_parts, write_parts, xml_part

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _document_with_grid_lanes(*, before=0, after=0, grid_width=5, rows=2):
    document = PythonDocxDocument()
    table = document.add_table(rows=rows, cols=3)
    for row_index, row in enumerate(table.rows):
        for cell_index, cell in enumerate(row.cells):
            cell.text = f"R{row_index}C{cell_index}"
    parts = read_parts(save_docx(document))
    root = etree.fromstring(parts["word/document.xml"])
    grid = root.find(f".//{W}tbl/{W}tblGrid")
    assert grid is not None
    grid.clear()
    for _ in range(grid_width):
        etree.SubElement(grid, W + "gridCol")
    parts["word/document.xml"] = etree.tostring(root)
    return _rewrite_table(write_parts(parts), [(before, after)] * rows)


def _rewrite_table(data, row_specs, mutate=None):
    parts = read_parts(data)
    root = etree.fromstring(parts["word/document.xml"])
    table = root.find(f".//{W}tbl")
    assert table is not None
    for row, (before, after) in zip(table.findall(f"./{W}tr"), row_specs, strict=True):
        properties = row.find(f"./{W}trPr")
        if properties is None:
            properties = etree.Element(W + "trPr")
            row.insert(0, properties)
        for tag, value in (("gridBefore", before), ("gridAfter", after)):
            if value:
                etree.SubElement(properties, W + tag).set(W + "val", str(value))
    if mutate is not None:
        mutate(table)
    parts["word/document.xml"] = etree.tostring(root)
    return write_parts(parts)


def _cell_text(cell):
    return "\n".join(
        block.text for block in cell.blocks() if isinstance(block, Paragraph)
    )


def test_standard_vertical_merge_preserves_following_physical_cells():
    document = PythonDocxDocument()
    table = document.add_table(rows=2, cols=3)
    for row_index, row in enumerate(table.rows):
        for cell_index, cell in enumerate(row.cells):
            cell.text = f"R{row_index}C{cell_index}"
    table.cell(0, 0).merge(table.cell(1, 0)).text = "MERGED"

    parsed = DocxDocument.parse(save_docx(document))
    table = next(block for block in parsed.blocks() if isinstance(block, Table))

    assert table.shape() == (2, 3)
    assert _cell_text(table[0, 0]) == "MERGED"
    assert table[0, 0] is table[1, 0]
    assert table[0, 0].bounds() == (0, 0, 2, 1)
    assert _cell_text(table[0, 1]) == "R0C1"
    assert _cell_text(table[0, 2]) == "R0C2"
    assert _cell_text(table[1, 1]) == "R1C1"
    assert _cell_text(table[1, 2]) == "R1C2"
    assert table[0, 1].bounds() == (0, 1, 1, 2)
    assert table[1, 1].bounds() == (1, 1, 2, 2)
    assert table[1, 2].bounds() == (1, 2, 2, 3)


def test_grid_before_only_keeps_leading_synthetic_lane():
    document = DocxDocument.parse(_document_with_grid_lanes(before=1, grid_width=4))
    table = next(block for block in document.blocks() if isinstance(block, Table))

    assert table.shape() == (2, 4)
    assert _cell_text(table[0, 0]) == ""
    assert _cell_text(table[0, 1]) == "R0C0"
    assert table[0, 0].blocks() == ()
    assert table[0, 0].bounds() == (0, 0, 1, 1)
    assert table[0, 0].path == "t:0/r:0/c:0"
    assert document.resolve(table[0, 0].path).path == table[0, 0].path


def test_grid_after_only_keeps_trailing_synthetic_lane():
    document = DocxDocument.parse(_document_with_grid_lanes(after=1, grid_width=4))
    table = next(block for block in document.blocks() if isinstance(block, Table))

    assert table.shape() == (2, 4)
    assert _cell_text(table[0, 2]) == "R0C2"
    assert _cell_text(table[0, 3]) == ""
    assert table[0, 3].bounds() == (0, 3, 1, 4)
    assert table[0, 3].path == "t:0/r:0/c:3"
    assert document.resolve(table[0, 3].path).path == table[0, 3].path


def test_uneven_rows_keep_each_row_logical_coordinates():
    data = _document_with_grid_lanes(grid_width=5)

    def remove_last_cell(tbl):
        tbl.findall(f"./{W}tr")[0].remove(
            tbl.findall(f"./{W}tr")[0].findall(f"./{W}tc")[-1]
        )

    document = DocxDocument.parse(
        _rewrite_table(data, [(1, 2), (0, 2)], remove_last_cell)
    )
    table = next(block for block in document.blocks() if isinstance(block, Table))

    assert table.shape() == (2, 5)
    assert _cell_text(table[0, 0]) == ""
    assert _cell_text(table[0, 1]) == "R0C0"
    assert _cell_text(table[0, 2]) == "R0C1"
    assert _cell_text(table[0, 3]) == ""
    assert _cell_text(table[1, 2]) == "R1C2"
    assert _cell_text(table[1, 4]) == ""
    assert table[0, 3].bounds() == (0, 3, 1, 4)
    assert document.resolve(table[0, 3].path).path == table[0, 3].path


def test_independent_grid_before_and_after_keep_synthetic_cells():
    data = _document_with_grid_lanes(before=1, after=1, grid_width=5)
    document = DocxDocument.parse(data)
    table = next(block for block in document.blocks() if isinstance(block, Table))

    assert table.shape() == (2, 5)
    assert _cell_text(table[0, 0]) == ""
    assert _cell_text(table[0, 1]) == "R0C0"
    assert _cell_text(table[1, 3]) == "R1C2"
    assert _cell_text(table[1, 4]) == ""
    assert table[0, 0].bounds() == (0, 0, 1, 1)
    assert table[0, 0].path == "t:0/r:0/c:0"
    assert document.resolve(table[0, 0].path).path == table[0, 0].path


def test_spans_and_vertical_merges_follow_leading_offset():
    data = _document_with_grid_lanes(before=1, after=1, grid_width=6)
    xml = xml_part(data)
    rows = xml.findall(f".//{W}tbl/{W}tr")

    first = rows[0].findall(f"./{W}tc")[0]
    tc_pr = first.find(f"./{W}tcPr")
    assert tc_pr is not None
    etree.SubElement(tc_pr, f"{W}gridSpan").set(f"{W}val", "2")
    etree.SubElement(tc_pr, f"{W}vMerge").set(f"{W}val", "restart")
    continuation = rows[1].findall(f"./{W}tc")[0]
    continuation_pr = continuation.find(f"./{W}tcPr")
    assert continuation_pr is not None
    etree.SubElement(continuation_pr, f"{W}gridSpan").set(f"{W}val", "2")
    etree.SubElement(continuation_pr, f"{W}vMerge")

    files = read_parts(data)
    files["word/document.xml"] = etree.tostring(xml)
    document = DocxDocument.parse(write_parts(files))
    table = next(block for block in document.blocks() if isinstance(block, Table))
    merged = table[0, 1]

    assert table.shape() == (2, 6)
    assert table[0, 1] is table[0, 2] is table[1, 1] is table[1, 2]
    assert _cell_text(merged) == "R0C0"
    assert merged.bounds() == (0, 1, 2, 3)
    assert merged.path == "t:0/r:0/c:1"
    assert document.resolve(merged.path).path == merged.path
    assert isinstance(table[0, 0], Cell)
    assert _cell_text(table[0, 0]) == ""


def test_vertical_merge_terminates_on_shifted_row_boundary():
    data = _document_with_grid_lanes(grid_width=4, rows=3)

    def add_merge(tbl):
        rows = tbl.findall(f"./{W}tr")
        first = rows[0].findall(f"./{W}tc")[0].find(f"./{W}tcPr")
        etree.SubElement(first, f"{W}vMerge").set(f"{W}val", "restart")
        continuation = rows[1].findall(f"./{W}tc")[1].find(f"./{W}tcPr")
        etree.SubElement(continuation, f"{W}vMerge")

    document = DocxDocument.parse(
        _rewrite_table(data, [(1, 0), (0, 1), (1, 0)], add_merge)
    )
    table = next(block for block in document.blocks() if isinstance(block, Table))

    merged = table[0, 1]
    assert table.shape() == (3, 4)
    assert table[0, 1] is table[1, 1]
    assert table[2, 1] is not merged
    assert merged.bounds() == (0, 1, 2, 2)
    assert table[2, 1].bounds() == (2, 1, 3, 2)
    assert merged.path == "t:0/r:0/c:1"
    assert document.resolve(merged.path).path == merged.path


def test_vertical_merge_does_not_cross_omitted_lane():
    data = _document_with_grid_lanes(grid_width=4, rows=3)

    def add_merge(tbl):
        rows = tbl.findall(f"./{W}tr")
        first = rows[0].findall(f"./{W}tc")[0].find(f"./{W}tcPr")
        etree.SubElement(first, f"{W}vMerge").set(f"{W}val", "restart")
        continuation = rows[1].findall(f"./{W}tc")[0].find(f"./{W}tcPr")
        etree.SubElement(continuation, f"{W}vMerge")

    document = DocxDocument.parse(
        _rewrite_table(data, [(1, 0), (2, 0), (0, 1)], add_merge)
    )
    table = next(block for block in document.blocks() if isinstance(block, Table))

    merged = table[0, 1]
    assert table.shape() == (3, 4)
    assert merged.bounds() == (0, 1, 1, 2)
    assert table[1, 1].path == "t:0/r:1/c:1"
    assert table[1, 1] is not merged
    assert table[2, 1].path == "t:0/r:2/c:1"
