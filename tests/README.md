# Development and tests

[Project overview](../README.md) · [Python API](../docs/API.md) · [Maintenance rules](../AGENTS.md)

Run commands from the repository root. Dependencies are managed with `uv`.
All test documents are generated with `python-docx`; no checked-in DOCX fixtures
or Word installation is required.

## Set up

```bash
uv sync --dev
uv run pre-commit install
```

## Run checks

Before committing:

```bash
uv run pre-commit run --all-files
uv run pytest
```

Pre-commit checks the lockfile, Ruff lint/format, static types with ty, and tests.
Its first run may download hook environments and tools.

During development, run the relevant domain or regression file:

```bash
uv run pytest tests/comments/ -q
uv run pytest tests/comments/test_comment_part_paths.py -v
uv run pytest tests/xml/test_xml_validity.py
```

## Coverage

```bash
uv run pytest --cov=docxnote --cov-report=term-missing
```

## Test layout

| Directory | Coverage | Regression examples |
| --- | --- | --- |
| [comments/](comments/) | Writing, reading, ranges, dates, whitespace, preservation | [Part paths and relationships](comments/test_comment_part_paths.py), [range ordering](comments/test_nested_comment_order.py), [unsupported ranges](comments/test_unsupported_ranges.py), [unanchored records](comments/test_unanchored_comments.py) |
| [document/](document/) | Paths, structure, addressable objects, thread safety | [Content controls](document/test_content_controls.py), [concurrent access](document/test_thread_safety.py) |
| [tables/](tables/) | Cell contents, nesting, merges, logical grid | [Grid lanes](tables/test_grid_lanes.py), [nested tables](tables/test_nested_tables.py) |
| [text/](text/) | Paragraph text and embedded-content scope | [Text parity](text/test_paragraph_text.py), [text-box scope](text/test_textbox_scope.py), [shared coordinates](text/test_shared_text_coordinates.py) |
| [xml/](xml/) | XML structure and DOCX package consistency | [Package validity](xml/test_xml_validity.py) |
| [shell/](shell/) | Commands, filters, paging, comments, concurrency | [Shell behavior](shell/test_shell.py) |

Shared fixtures live in [conftest.py](conftest.py). Comment XML/package builders
live in [comments/_helpers.py](comments/_helpers.py).

## Add a regression

Assert public views and rendered/reparsed packages rather than private storage layouts.
Place new tests in the closest domain. Reproduce the failure before changing the
implementation, then cover the affected public behavior and render/reparse cycle
where relevant. Package changes should check relationships and content types as
well as visible comments.

Keep English and Chinese behavior references synchronized; the ownership and
update rules are in [AGENTS.md](../AGENTS.md#文档分工).
