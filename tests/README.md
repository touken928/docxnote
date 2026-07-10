# Tests

## Install dependencies

```bash
uv sync --dev
```

## Run tests

```bash
# Run all tests
uv run pytest

# Verbose output
uv run pytest -v

# Run one file
uv run pytest tests/xml/test_xml_validity.py
```

## Coverage

```bash
uv run pytest --cov=docxnote --cov-report=term-missing
```

## Test layout

- **cli/**
  - **test_cli.py** - CLI command behavior
- **comments/**
  - **test_comment_writing.py** - comment creation and render behavior
  - **test_comment_reading.py** - comment reading, precise ranges, and `keep_comments`
  - **test_comment_conflict_ranges.py** - overlapping comments and range/run splitting
  - **test_existing_comment_preservation.py** - existing comment preservation and `comments.xml` metadata
- **document/**
  - **test_addressable_units.py** - addressable paths and paragraph traversal
  - **test_paths.py** - path parsing and building helpers
  - **test_structure_comparison.py** - structure parity with `python-docx`
  - **test_thread_safety.py** - shared-document thread safety
- **tables/**
  - **test_cell_content.py** - cell text extraction
  - **test_nested_tables.py** - nested table traversal
  - **test_table_shape.py** - table shape and merged-cell behavior
- **text/**
  - **test_paragraph_text.py** - paragraph text extraction
- **xml/**
  - **test_xml_validity.py** - package and XML validity

All test documents are generated dynamically with `python-docx`; there are no checked-in DOCX fixtures.
