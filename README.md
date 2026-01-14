# pivoteer

pivoteer injects pandas DataFrames into existing Excel templates by editing the
underlying XML. It resizes Excel Tables (ListObjects) and forces PivotTables to
refresh on open without corrupting pivot caches.

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

## Features

- Surgical Data Injection: updates worksheet XML without touching sharedStrings.
- Table Resizing: recalculates ListObject ranges to match injected data.
- Pivot Preservation: sets pivot caches to refresh on load when present.
- Minimal IO: stream-based ZIP copy-and-replace for stability.

## Limitations

- The generated test template uses inline strings and does not create pivot
  caches when the installed xlsxwriter lacks pivot table support.
- Date formatting is injected as inline text; apply Excel formatting if needed.
- Shared strings are not modified in Phase 1.

## Development

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .[dev]
pytest
```
