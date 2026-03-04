---
type: documentation
entity: project-overview
version: 1.0
---

# pivoteer

## Purpose

pivoteer injects pandas DataFrames into existing Excel templates by editing OpenXML
parts directly inside `.xlsx` ZIP archives, preserving PivotTable wiring and template
formatting while updating table-backed source data.

## Architecture

The project is a Python library centered around a single package (`src/pivoteer`) with
an orchestration layer (`Pivoteer` -> `TemplateEngine`) and lower-level XML/ZIP handlers
(`XmlEngine`, `TableResizer`, `pivot_cache_updater`).

Typical flow:
- Load workbook metadata (worksheets, tables, pivot caches) from workbook XML relations.
- Inject DataFrame values into target worksheet cells.
- Resize ListObject/table XML references to match injected rows/columns.
- Mark pivot caches for refresh on workbook open.
- Optionally append missing pivot cache field definitions for newly added table columns.
- Repack modified XML parts into a new output workbook.

### System Diagram

```text
Caller code
   |
   v
Pivoteer (core.py)
   |
   v
TemplateEngine (template_engine.py)
   |------------------------------.
   |                              |
   v                              v
XmlEngine (xml_engine.py)     TableResizer (table_resizer.py)
   |                              |
   '-------> modified XML parts <-'
                   |
                   v
   (optional) pivot_cache_updater.sync_cache_fields()
                   |
                   v
         write new .xlsx archive
```

### Tech Stack

- Language: Python 3.10+
- Core libs: `pandas`, `lxml`
- Packaging/build: `hatchling`
- Dev/test: `pytest`, `pytest-cov`, `ruff`, `xlsxwriter` (fixture/template generation)
- CI/CD: GitHub Actions (`.github/workflows/ci.yml`, `.github/workflows/release.yml`)

## Modules

| Module | Description | Documentation |
|--------|-------------|---------------|
| pivoteer-library | Runtime package for workbook mapping, XML injection, table resizing, and pivot metadata updates. | [Detail](modules/pivoteer-library.md) |

## Key Features

| Feature | Description | Documentation |
|---------|-------------|---------------|
| dataframe-injection | Injects DataFrame values directly into worksheet XML cells using numeric and inline string representations. | [Detail](features/dataframe-injection.md) |
| table-range-resizing | Recomputes Excel table (`ref`) ranges after data updates to keep ListObject boundaries aligned. | [Detail](features/table-range-resizing.md) |
| pivot-refresh-on-open | Sets `refreshOnLoad="1"` in pivot cache definitions so Excel refreshes pivots when opening the output file. | [Detail](features/pivot-refresh-on-open.md) |
| pivot-cache-field-sync | Optionally appends missing pivot cache fields when new table columns are introduced. | [Detail](features/pivot-cache-field-sync.md) |

## Development

### Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .[dev]
```

### Build & Run

```bash
python -m build
python demo_runner.py
```

### Testing

```bash
pytest
pytest --cov=pivoteer --cov-report=term-missing --cov-fail-under=70
ruff check src/ tests/
ruff format --check src/ tests/
```

Testing strategy is mixed unit + integration, with synthetic and real pivot fixtures in
`tests/` and `tests/fixtures/`.

## References

- Project README: `README.md`
- Changelog/history: `CHANGELOG.md`
- Public package config: `pyproject.toml`
- CI workflow: `.github/workflows/ci.yml`
- Release workflow: `.github/workflows/release.yml`
