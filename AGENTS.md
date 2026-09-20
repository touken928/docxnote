# Repository Guidelines

## Project Structure & Module Organization

- `src/docxnote/`: document/package handling, paragraphs, tables, comments, paths, and the in-memory shell.
- `tests/`: pytest suites grouped into `comments/`, `document/`, `tables/`, `text/`, `xml/`, and `shell/`. Shared fixtures live in `conftest.py`; DOCX inputs are generated with `python-docx`, not stored as assets.
- `docs/`: authoritative English/Chinese API and Shell references. README files provide installation and minimal examples.

## Build, Test, and Development Commands

Use Python 3.12+ and `uv`; keep `pyproject.toml` and `uv.lock` synchronized.

```bash
uv sync --dev                           # Install development dependencies
uv run pre-commit install               # Enable commit hooks
uv run pytest tests/comments/ -q        # Run an affected domain
uv run pytest                          # Run the complete suite
uv run pre-commit run --all-files       # Check lockfile, Ruff, ty, and tests
uv run pytest --cov=docxnote --cov-report=term-missing
uv build                               # Build wheel and source distribution
```

Run both full pytest and pre-commit checks before committing.

## Coding Style & Architecture

Use four-space indentation, `snake_case` functions/modules, and `PascalCase` classes. Follow Ruff lint/format and ty checks configured in `.pre-commit-config.yaml`.

Keep patches focused. Expose text views and high-level objects; encapsulate Word runs and XML internally. Hold the document's `threading.RLock` when accessing shared XML or comment state. Comment offsets must remain `[start, end)` in `paragraph.text`. With `keep_comments=True`, preserve existing comment metadata, part paths, and package relationships.

## Testing Guidelines

Name files `test_*.py` and functions `test_*`. Place tests in the closest domain, not the test root. New features and behavior changes require coverage; reproduce bugs with failing regression tests before fixing them.

Prioritize partial ranges, hyperlinks, content controls, text boxes, merged/nested tables, and comment preservation/stripping. Verify render/reparse behavior and package relationships/content types where relevant. See `tests/README.md` for coverage and commands.

## Documentation Maintenance

Update both README languages and `docs/API.md` / `docs/API_zh.md` for public API or behavior changes. Update `docs/SHELL.md` / `docs/SHELL_zh.md` for Shell changes and `tests/README.md` for test-layout/command changes. Keep README examples minimal, API methods under their objects, and framework integration after direct Shell usage. Preserve navigation links and moved-section anchors. Retain README badges and centered H1 titles.

## Commit & Pull Request Guidelines

Follow history's concise imperative subjects, e.g. `Fix comment preservation`; no mandatory prefix. PRs should explain the problem, resulting behavior, regression coverage, validation results, and relevant issues. Include synchronized documentation.

For releases, update the project version and push a matching `v*` tag; `.github/workflows/publish.yml` publishes through PyPI Trusted Publishing.
