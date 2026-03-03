---
type: documentation
entity: feature
feature: "pivot-cache-field-sync"
version: 1.0
---

# Feature: pivot-cache-field-sync

> Part of [pivoteer](../overview.md)

## Summary

This optional feature appends missing pivot cache field definitions when table columns
change, helping newly added source columns appear in PivotTable field lists.

## How It Works

When enabled, pivoteer checks updated tables after data injection, locates pivot cache
definitions sourcing those tables, compares table column names to existing cache field
names, and appends missing `<cacheField>` nodes while preserving existing order.

### User Flow

1. User initializes `Pivoteer(..., enable_pivot_field_sync=True)`.
2. User applies DataFrame that may include new columns for target table.
3. User saves workbook.
4. Opened workbook exposes newly added fields in PivotTable field chooser.

### Technical Flow

1. `Pivoteer.save` checks `enable_pivot_field_sync` and calls sync path.
2. `TemplateEngine.sync_pivot_cache_fields` iterates updated table names.
3. `sync_cache_fields` loads table + pivot cache XML parts and filters caches by
   matching `worksheetSource name`.
4. `_append_missing_cache_fields` appends missing field nodes and updates `count`.
5. Modified pivot cache parts are staged and written into output workbook.

## Implementation

| Module | Symbols | Role |
|--------|---------|------|
| [pivoteer-library](../modules/pivoteer-library.md) | `Pivoteer.__init__` (`src/pivoteer/core.py:19`) | Captures `enable_pivot_field_sync` flag from caller. |
| [pivoteer-library](../modules/pivoteer-library.md) | `Pivoteer.save` (`src/pivoteer/core.py:30`) | Conditionally invokes field-sync workflow before writing archive. |
| [pivoteer-library](../modules/pivoteer-library.md) | `TemplateEngine.sync_pivot_cache_fields` (`src/pivoteer/template_engine.py:94`) | Triggers synchronization for each table updated in current session. |
| [pivoteer-library](../modules/pivoteer-library.md) | `sync_cache_fields` (`src/pivoteer/pivot_cache_updater.py:16`) | Main reconciliation routine between table schema and pivot cache schema. |
| [pivoteer-library](../modules/pivoteer-library.md) | `_append_missing_cache_fields` (`src/pivoteer/pivot_cache_updater.py:75`) | Performs append-only cache field updates and count recalculation. |

## Configuration

- Controlled by constructor flag: `enable_pivot_field_sync` (default `False`).
- Runs only for tables updated via `apply_dataframe` in the current engine instance.

## Edge Cases & Limitations

- Missing table columns metadata raises `PivotCacheError`.
- Missing `<cacheFields>` element raises `PivotCacheError`.
- Only append-only reconciliation is implemented; existing cache fields are not removed
  or reordered.
- Pivot cache source matching uses `worksheetSource name`; mismatched naming skips updates.

## Related Features

- [pivot-refresh-on-open](pivot-refresh-on-open.md)
- [table-range-resizing](table-range-resizing.md)
