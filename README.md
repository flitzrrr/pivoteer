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
- Pivot cache field sync (opt-in): appends missing cache field entries for table
  columns so new headers appear in existing PivotTables.

## Features

- Surgical Data Injection: updates worksheet XML without touching sharedStrings.
- Table Resizing: recalculates ListObject ranges to match injected data.
- Pivot Preservation: sets pivot caches to refresh on load when present.
- Optional Pivot Cache Field Sync: appends missing cache field metadata for new
  table columns without touching PivotTable layouts.
- Minimal IO: stream-based ZIP copy-and-replace for stability.

## Pivot Cache Field Sync

When new columns are added to an Excel Table, existing PivotTables often fail to
show the new fields until the PivotCache metadata is updated. pivoteer can
synchronize PivotCache field definitions so new table columns appear in the
PivotTable field list.

What pivoteer does:

- Syncs PivotCache field metadata for the target table.
- Appends missing cache fields so new columns are visible in the PivotTable UI.

What pivoteer does not do:

- Does not create PivotTables.
- Does not modify PivotTable layouts or filters.
- Does not touch slicers or formatting.

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

### Opt-in pivot cache field sync

```python
from pivoteer.core import Pivoteer
import pandas as pd

p = Pivoteer("template.xlsx", enable_pivot_field_sync=True)
p.apply_dataframe("RawData", pd.read_csv("usage.csv"))
p.save("report_output.xlsx")
```

This flag is optional; when it is not set, pivoteer behaves exactly as before.

### Advanced usage with TemplateEngine

```python
from pathlib import Path
import pandas as pd

from pivoteer.template_engine import TemplateEngine

engine = TemplateEngine(Path("template.xlsx"))
engine.apply_dataframe("RawData", pd.read_csv("usage.csv"))
engine.sync_pivot_cache_fields()
engine.ensure_pivot_refresh_on_load()
parts = engine.get_modified_parts()
```

### Large datasets

pivoteer is optimized for replacing table data without rewriting the entire
workbook. It is a good fit for large tables where preserving PivotTables and
filters matters more than Excel formatting for each row.

## Safety Guarantees

- Opt-in only: the feature is disabled unless explicitly enabled.
- Only missing cache fields are appended.
- Existing cache field order is preserved.
- PivotTable definitions are not modified.

## Limitations

- The PivotCache source must reference the named Excel Table.
- The template must already contain PivotTables and pivot caches.
- The structured table must exist and be the PivotTable cache source.
- pivoteer does not auto-refresh the Excel UI; Excel recalculates pivots on open.

## Compatibility

- Python: 3.10+
- Excel: Desktop Excel (Windows/macOS) supports `refreshOnLoad` for PivotTables.
- Templates: Must include Excel Tables (ListObjects) with stable names.

## Troubleshooting

- "Table not found": Ensure the Excel Table name matches exactly.
- "Pivot cache not found": The template may not include a PivotTable; this is
  expected for synthetic templates.
- "DataFrame is empty": pivoteer refuses empty payloads to protect templates.

## Support and Requests

- Bugs: open a GitHub issue using the Bug Report template.
- Feature requests: open a GitHub issue using the Feature Request template.
- Security: follow the reporting process in `SECURITY.md`.

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
