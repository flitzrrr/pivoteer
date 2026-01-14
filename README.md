# pivoteer

pivoteer injects pandas DataFrames into existing Excel templates by editing the underlying XML. It resizes Excel Tables (ListObjects) and forces PivotTables to refresh on open without corrupting pivot caches.

## Status

Phase 1 scaffold with a template generator for test inputs.

## Development

Install dependencies:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .[dev]
```

Generate a dummy template:

```bash
python tests/generate_dummy_template.py --output ./dummy_template.xlsx
```
