---
type: documentation
entity: feature
feature: "pivot-refresh-on-open"
version: 1.0
---

# Feature: pivot-refresh-on-open

> Part of [pivoteer](../overview.md)

## Summary

This feature ensures existing PivotTables recalculate against updated source data by
setting `refreshOnLoad="1"` on discovered pivot cache definition XML parts.

## How It Works

During save flow, pivoteer enumerates pivot cache definition paths from workbook
relationships, loads each XML part, sets the refresh flag attribute, and writes the
modified parts into the output workbook.

### User Flow

1. User updates one or more tables and saves output workbook.
2. User opens output workbook in Excel.
3. Excel refreshes pivot caches marked with `refreshOnLoad`, reflecting new source data.

### Technical Flow

1. `XmlEngine.build_workbook_map` collects pivot cache paths from workbook rels.
2. `TemplateEngine.ensure_pivot_refresh_on_load` iterates pivot cache parts.
3. For each part, root attribute `refreshOnLoad` is set to `1`.
4. Modified cache XML trees are staged and serialized during save.

## Implementation

| Module | Symbols | Role |
|--------|---------|------|
| [pivoteer-library](../modules/pivoteer-library.md) | `XmlEngine.build_workbook_map` (`src/pivoteer/xml_engine.py:54`) | Populates pivot cache definition part map. |
| [pivoteer-library](../modules/pivoteer-library.md) | `XmlEngine._parse_pivot_caches` (`src/pivoteer/xml_engine.py:186`) | Detects pivot cache relationship targets. |
| [pivoteer-library](../modules/pivoteer-library.md) | `TemplateEngine.ensure_pivot_refresh_on_load` (`src/pivoteer/template_engine.py:81`) | Applies refresh flag mutation on each pivot cache definition. |
| [pivoteer-library](../modules/pivoteer-library.md) | `Pivoteer.save` (`src/pivoteer/core.py:30`) | Ensures pivot refresh processing occurs before archive write. |

## Configuration

- No user-facing flag; refresh-flag update runs on every `save()`.
- If no pivot caches are present in template, method exits without changes.

## Edge Cases & Limitations

- Feature does not create pivot caches; it only mutates existing cache definitions.
- Excel desktop behavior determines actual refresh timing/UI after file open.
- Templates with nonstandard pivot relationship layouts may be skipped if unresolved.

## Related Features

- [dataframe-injection](dataframe-injection.md)
- [pivot-cache-field-sync](pivot-cache-field-sync.md)
