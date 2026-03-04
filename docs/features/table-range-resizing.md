---
type: documentation
entity: feature
feature: "table-range-resizing"
version: 1.0
---

# Feature: table-range-resizing

> Part of [pivoteer](../overview.md)

## Summary

This feature recalculates and updates the Excel table (`ListObject`) `ref` range to
match the injected DataFrame dimensions so table boundaries remain valid for filters
and PivotTable sources.

## How It Works

After rows are injected, pivoteer parses the original A1 range, preserves the table
start coordinates and header row, computes new end row/column from DataFrame shape,
and writes the updated `ref` back into table XML.

### User Flow

1. User calls `apply_dataframe` with table data.
2. User saves workbook.
3. Opened output workbook shows table range expanded/contracted to new data size.

### Technical Flow

1. `TemplateEngine.apply_dataframe` computes `row_count` and `col_count` from DataFrame.
2. Existing table range is read from `TableRef.ref` and parsed with `parse_a1_range`.
3. `TableResizer.resize_table` computes `updated_ref` and mutates `<table ref="...">`.
4. Updated table XML is staged for output serialization.

## Implementation

| Module | Symbols | Role |
|--------|---------|------|
| [pivoteer-library](../modules/pivoteer-library.md) | `TemplateEngine.apply_dataframe` (`src/pivoteer/template_engine.py:37`) | Invokes table resizing after worksheet injection. |
| [pivoteer-library](../modules/pivoteer-library.md) | `TableResizer.resize_table` (`src/pivoteer/table_resizer.py:24`) | Core range recomputation logic and XML mutation. |
| [pivoteer-library](../modules/pivoteer-library.md) | `parse_a1_range` (`src/pivoteer/utils.py:45`) | Parses existing A1 table range into coordinates. |
| [pivoteer-library](../modules/pivoteer-library.md) | `build_a1_range` (`src/pivoteer/utils.py:62`) | Builds updated A1 range string for table `ref`. |

## Configuration

- No feature flag; runs automatically during `apply_dataframe`.
- Requires table metadata (`TableRef.ref`) to be present and valid A1 notation.

## Edge Cases & Limitations

- Table XML missing `ref` raises `XmlStructureError`.
- Zero-column payloads are invalid and rejected before resize.
- Header row count is fixed to one row in current implementation.

## Related Features

- [dataframe-injection](dataframe-injection.md)
- [pivot-cache-field-sync](pivot-cache-field-sync.md)
