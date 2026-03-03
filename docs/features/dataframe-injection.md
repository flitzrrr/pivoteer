---
type: documentation
entity: feature
feature: "dataframe-injection"
version: 1.0
---

# Feature: dataframe-injection

> Part of [pivoteer](../overview.md)

## Summary

This feature replaces table-backed worksheet data in an Excel template with rows from
a pandas DataFrame while preserving workbook structure and avoiding a full workbook
rewrite.

## How It Works

DataFrames are validated for non-empty rows and columns, converted to row lists, then
injected into worksheet XML cells starting at the table's first data row. Values are
encoded as numeric (`<v>`), inline strings (`<is><t>`), or empty cells.

### User Flow

1. User constructs `Pivoteer(template_path)`.
2. User calls `apply_dataframe(table_name, df)` with a target table and DataFrame.
3. User calls `save(output_path)` to write a new workbook.
4. Output workbook contains updated source data in the target worksheet table area.

### Technical Flow

1. `Pivoteer.apply_dataframe` delegates to `TemplateEngine.apply_dataframe`.
2. `TemplateEngine` validates table existence and DataFrame shape.
3. Table `ref` start row/column is parsed to compute injection start row.
4. `XmlEngine.inject_rows_inline_strings` creates/updates row and cell nodes.
5. Updated worksheet XML tree is staged in `TemplateEngine._modified_trees`.

## Implementation

| Module | Symbols | Role |
|--------|---------|------|
| [pivoteer-library](../modules/pivoteer-library.md) | `Pivoteer.apply_dataframe` (`src/pivoteer/core.py:26`) | Public entry point for DataFrame application. |
| [pivoteer-library](../modules/pivoteer-library.md) | `TemplateEngine.apply_dataframe` (`src/pivoteer/template_engine.py:37`) | Validates inputs, computes coordinates, and stages worksheet/table updates. |
| [pivoteer-library](../modules/pivoteer-library.md) | `XmlEngine.inject_rows_inline_strings` (`src/pivoteer/xml_engine.py:86`) | Performs cell-level XML mutation with type-aware value encoding. |
| [pivoteer-library](../modules/pivoteer-library.md) | `XmlEngine._set_cell_value_inline` (`src/pivoteer/xml_engine.py:245`) | Encodes numeric, text, date-like, and missing values into OpenXML cells. |

## Configuration

- No dedicated feature flag; feature is always active when `apply_dataframe` is called.
- Behavior depends on DataFrame values and target table metadata from workbook map.

## Edge Cases & Limitations

- Empty DataFrames are rejected (`InvalidDataError`).
- Missing table names are rejected (`TableNotFoundError`).
- Data is injected as raw values; pivoteer does not apply custom cell formatting.
- Date/datetime values are written as ISO strings (inline strings), not Excel serial dates.

## Related Features

- [table-range-resizing](table-range-resizing.md)
- [pivot-refresh-on-open](pivot-refresh-on-open.md)
