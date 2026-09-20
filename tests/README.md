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
uv run pytest tests/package/test_part_paths.py -v
uv run pytest tests/package/test_roundtrip.py
```

## Coverage

```bash
uv run pytest --cov=docxnote --cov-report=term-missing
```

The current suite contains 177 behavior tests. The suite is intentionally
smaller than the historical test layout: repeated setup, duplicate roundtrip
assertions, and tests that only checked that rendering returned bytes have been
removed. Each remaining test should assert a public result, a rendered package
invariant, or a synchronization guarantee.

## Test layout

| Directory | Coverage | Regression examples |
| --- | --- | --- |
| [comments/](comments/) | Creation, ranges, dates, unsupported anchors, unanchored records | [Range matrix](comments/test_ranges.py), [source dates](comments/test_dates.py), [unsupported ranges](comments/test_unsupported_ranges.py) |
| [document/](document/) | Navigation, paths, content controls, thread safety | [Navigation](document/test_navigation.py), [concurrent access](document/test_thread_safety.py) |
| [package/](package/) | Parts, relationships, content types, source XML preservation | [Part paths](package/test_part_paths.py), [roundtrip](package/test_roundtrip.py) |
| [tables/](tables/) | Cell content, nested tables, merges, omitted grid lanes | [Grid lanes](tables/test_grid_lanes.py), [nested tables](tables/test_nested.py) |
| [text/](text/) | Text views, inline containers, text boxes, whitespace, coordinates | [Text scope](text/test_text.py), [inline content](text/test_inline_content.py) |
| [shell/](shell/) | Commands, filters, paging, comments, concurrency | [Shell behavior](shell/test_shell.py) |

Shared generators live in [support/docx.py](support/docx.py) and
[support/comments.py](support/comments.py). They create source documents and
package variants without calling private docxnote implementation code.

## Add a regression

Assert public views and rendered/reparsed packages rather than private storage layouts.
Place new tests in the closest domain. Reproduce the failure before changing the
implementation, then cover the affected public behavior and render/reparse cycle
where relevant. Package changes should check relationships and content types as
well as visible comments.

Do not add a fixture for one test. Add a small generator to `tests/support/` when
the same document or package mutation is needed by multiple suites. Keep source
document generation separate from package-part mutation, and make generated
documents deterministic so failure output identifies the behavior under test.

Keep English and Chinese behavior references synchronized; the ownership and
update rules are in [AGENTS.md](../AGENTS.md).
