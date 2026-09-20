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

### Module Boundaries

Keep the public API stable during internal refactors, including names, argument defaults, properties, path syntax, comment policies, and Shell commands. Underscore-prefixed modules and attributes are internal.

| Module | Required responsibility |
| --- | --- |
| `document.py` | Coordinate public operations, traversal snapshots, and serialization; delegate package and XML algorithms. |
| `paragraph.py` | Expose paragraph text and anchored comment views; delegate run splitting and anchor parsing. |
| `table.py` | Construct public `Table` / `Cell` views and maintain one immutable coordinate matrix. |
| `comments.py` | Define immutable `Comment` snapshots and export the public range exception; do not store mutable records here. |
| `paths.py` | Parse and construct string addresses; leave structural validation and object lookup to navigation. |
| `shell.py` | Parse commands, filter records, enforce output budgets, and track session comments; use high-level document operations. |
| `_state.py` | Own shared XML roots, package, comment store, preservation policy, and the single document lock. |
| `_package.py` | Handle ZIP parts, comment part paths, relationships, and content types; do not interpret paragraph text or anchors. |
| `_comment_store.py` | Own IDs, ordered records, metadata, and preserved source comment XML; do not manage anchors or ZIP paths. |
| `_navigation.py` | Centralize view construction, block traversal, merged-cell deduplication, and path resolution. |
| `_xml/` | Implement namespaces, block traversal, text coordinates, anchors, and grid parsing independently of document state, locks, Shell, and public view objects. |

Keep `__init__.py` limited to public exports. Import internal dependencies from their defining modules, never through the package entry point. Preserve the `namespaces.NS` compatibility import. Use `TYPE_CHECKING` for type-only view references. Keep deferred concrete-view imports at the navigation construction boundary so `Cell.blocks()` can construct nested views without an import-time cycle. Pass shared state through `DocumentOwner`; do not make child views call back into the public document facade.

### State, Locking, and Caches

- Use one `DocumentState` and one reentrant lock per document, shared by every view and bound Shell session. Hold it for package access, shared XML access, record creation, and anchor changes. XML helpers and the comment store rely on their callers' operation boundaries; do not introduce independent locks.
- Hold the document lock for an entire Shell command, including validation, pipeline evaluation, writes, result collection, and session bookkeeping. For `DocxDocument.iter_paragraphs()`, snapshot wrappers under the lock and release it before yielding. Wrappers remain live views, not complete XML snapshots.
- Construct table matrices and cell bounds under the lock and keep them unchanged after publication.
- Cache paragraph text only while supported mutations preserve visible text. Read comment ranges afresh so older wrappers observe comments added through other wrappers. Before introducing text editing, implement shared cache invalidation.

### Text, Comments, Tables, and Paths

- Centralize direct run-child text contributions (`w:t`, `w:tab`, `w:br`) and inline-wrapper traversal in `_xml/text.py`. Text reads, run spans, boundary lookup, splitting, and range parsing must share these rules. Treat runs as traversal leaves so embedded text boxes and their markers cannot affect host-paragraph coordinates. Retain the separate descendant-paragraph extraction policy for comment bodies.
- Keep `Paragraph.comment()` slice normalization, run validation, record creation, and anchor insertion within one locked operation. Validate the presence of runs before allocating an ID. Preserve the distinct contract of `DocxDocument.add_comment()`: store an unanchored body and return its ID; serialize it without exposing it in anchored comment views.
- Use named records such as `RunSpan`, `CommentRange`, `InsertionPoint`, and `CommentRecord` instead of opaque positional tuples. Keep the ordered comment records authoritative and use the ID index only for lookup. Preserve original XML alongside interpreted metadata; render existing comments from source XML rather than reconstructing them from the text view. Missing or invalid dates must remain `None` in views without changing their source attributes.
- Reject cross-paragraph or unpaired anchors only when reading range views. Parsing and rendering must remain able to preserve such XML. Preserve resolved comment part paths and relationships when keeping comments; remove the old comment part's own relationships when stripping them.
- Parse table properties into specifications and `CellRegion` data before constructing public cells. Share one `Cell` per merged origin across the coordinate matrix; represent omitted coordinates as synthetic empty cells. Preserve declared grid widths, merge termination, and fallback semantics.
- Use one block factory for document bodies and cells. Count paragraph and table indices independently at each level, expand content controls transparently, traverse merged cells once, and address cells by their merge origins.

### Names and Source Documentation

Use English source docstrings and comments. Explain contracts and reasons for non-obvious preservation, traversal, and locking behavior; avoid comments that merely restate code. Reserve `iter_*` for iterator-returning functions and use `collect_*` or `parse_*` for materialized results. Initialize internal state explicitly rather than masking incomplete initialization with `getattr` fallbacks.

## Testing Guidelines

Name files `test_*.py` and functions `test_*`. Place tests in the closest domain, not the test root. New features and behavior changes require coverage; reproduce bugs with failing regression tests before fixing them.

Prioritize partial ranges, hyperlinks, content controls, text boxes, merged/nested tables, and comment preservation/stripping. Verify render/reparse behavior and package relationships/content types where relevant. See `tests/README.md` for coverage and commands.

Verify compatibility through public views and rendered/reparsed packages rather than private storage layouts. Include unanchored ID allocation, shared coordinates across wrappers, and concurrent operations when changing state ownership or XML algorithms.

## Documentation Maintenance

Update both README languages and `docs/API.md` / `docs/API_zh.md` for public API or behavior changes. Update `docs/SHELL.md` / `docs/SHELL_zh.md` for Shell changes and `tests/README.md` for test-layout/command changes. Keep README examples minimal, API methods under their objects, and framework integration after direct Shell usage. Preserve navigation links and moved-section anchors. Retain README badges and centered H1 titles.

Maintain internal architecture as enforceable rules in this `AGENTS.md`, not as separate architecture documents under `docs/`. Keep user-facing references focused on public interfaces and behavior.

## Commit & Pull Request Guidelines

Follow history's concise imperative subjects, e.g. `Fix comment preservation`; no mandatory prefix. PRs should explain the problem, resulting behavior, regression coverage, validation results, and relevant issues. Include synchronized documentation.

For releases, update the project version and push a matching `v*` tag; `.github/workflows/publish.yml` publishes through PyPI Trusted Publishing.
