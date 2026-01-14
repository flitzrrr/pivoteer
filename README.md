# pivoteer

[![CI](https://github.com/flitzrrr/pivoteer/actions/workflows/ci.yml/badge.svg)](https://github.com/flitzrrr/pivoteer/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/pivoteer.svg)](https://pypi.org/project/pivoteer/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

pivoteer injects pandas DataFrames into existing Excel templates by editing the
underlying XML. It resizes Excel Tables (ListObjects) and forces PivotTables to
refresh on open without corrupting pivot caches.

## Why pivoteer

Most Python Excel libraries rewrite workbooks, which can break PivotTables,
filters, and formatting in real-world templates. pivoteer is designed for
enterprise reporting workflows where templates are authored in Excel and must
remain intact. It surgically updates only the table data and table metadata so
PivotTables remain connected and refresh correctly.

## Installation

```bash
pip install pivoteer
```

## Quick Start

```python
from pathlib import Path
import pandas as pd

from pivoteer.core import Pivoteer

pivoteer = Pivoteer(Path("template.xlsx"))

df = pd.DataFrame(
    {
        "Category": ["Hardware", "Software"],
        "Region": ["North", "South"],
        "Amount": [120.0, 250.0],
        "Date": ["2024-01-01", "2024-01-02"],
    }
)

pivoteer.apply_dataframe("DataSource", df)
pivoteer.save("report_output.xlsx")
```

## Architecture Overview

- Input/output: `.xlsx` files are ZIP archives containing OpenXML parts.
- Data injection: updates `xl/worksheets/sheetN.xml` row data using inline
  strings to avoid touching sharedStrings.xml.
- Table resizing: updates `xl/tables/tableN.xml` by recalculating the `ref`
  range based on the DataFrame shape.
- Pivot refresh: sets `refreshOnLoad="1"` in
  `xl/pivotCache/pivotCacheDefinitionN.xml` when present.

## Features

- Surgical Data Injection: updates worksheet XML without touching sharedStrings.
- Table Resizing: recalculates ListObject ranges to match injected data.
- Pivot Preservation: sets pivot caches to refresh on load when present.
- Minimal IO: stream-based ZIP copy-and-replace for stability.

## Usage Patterns

### Multiple table updates

```python
from pivoteer.core import Pivoteer
import pandas as pd

p = Pivoteer("template.xlsx")
p.apply_dataframe("SalesData", pd.read_csv("sales.csv"))
p.apply_dataframe("CostData", pd.read_csv("costs.csv"))
p.save("report_output.xlsx")
```

### Large datasets

pivoteer is optimized for replacing table data without rewriting the entire
workbook. It is a good fit for large tables where preserving PivotTables and
filters matters more than Excel formatting for each row.

## Limitations

- The generated test template uses inline strings and does not create pivot
  caches when the installed xlsxwriter lacks pivot table support.
- Date formatting is injected as inline text; apply Excel formatting if needed.
- Shared strings are not modified in Phase 1.
- PivotTables are refreshed on open via `refreshOnLoad`, but pivoteer does not
  recalculate pivot caches or modify pivot layout.

## Compatibility

- Python: 3.10+
- Excel: Desktop Excel (Windows/macOS) supports `refreshOnLoad` for PivotTables.
- Templates: Must include Excel Tables (ListObjects) with stable names.

## Troubleshooting

- "Table not found": Ensure the Excel Table name matches exactly.
- "Pivot cache not found": The template may not include a PivotTable; this is
  expected for synthetic templates.
- "DataFrame is empty": pivoteer refuses empty payloads to protect templates.

## Security

If you discover a vulnerability, please read `SECURITY.md` for reporting
instructions.

## Development

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .[dev]
pytest
```
