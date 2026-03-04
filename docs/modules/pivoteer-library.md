---
type: documentation
entity: module
module: "pivoteer-library"
version: 1.0
---

# Module: pivoteer-library

> Part of [pivoteer](../overview.md)

## Overview

This module is the runtime library under `src/pivoteer`. It provides a high-level API
for updating Excel template tables from DataFrames while preserving pivot behavior,
plus the low-level XML parsing/injection logic used to perform targeted modifications.

### Responsibility

- In scope: workbook relationship discovery, worksheet cell injection, table range
  recalculation, pivot refresh flags, optional pivot cache field synchronization, and
  serialization of modified XML parts.
- Out of scope: creating Excel templates, authoring pivot layouts, UI automation,
  spreadsheet formula management, and direct Excel process control.

### Dependencies

| Dependency | Type | Purpose |
|-----------|------|---------|
| `pandas` | library | Provides DataFrame input model and missing-value detection (`pd.isna`). |
| `lxml.etree` | library | Parses and mutates OpenXML documents while preserving XML structure. |
| `zipfile` | library | Reads and rewrites `.xlsx` archives as ZIP containers. |
| `pathlib.Path` | library | Normalized file path handling for templates and outputs. |
| `re` | library | Validates A1 notation for utility conversions/parsing. |
| `posixpath` | library | Resolves workbook relationship targets to normalized ZIP part paths. |

## Structure

| Path | Type | Purpose |
|------|------|---------|
| `src/pivoteer/` | dir | Package root for runtime library code. |
| `src/pivoteer/__init__.py` | file | Package metadata and exported version string. |
| `src/pivoteer/core.py` | file | Public entry API (`Pivoteer`) that coordinates write flow and archive output. |
| `src/pivoteer/template_engine.py` | file | Main orchestrator combining XML injection, resizing, pivot updates, and part serialization. |
| `src/pivoteer/xml_engine.py` | file | Low-level workbook map building and worksheet XML cell/row manipulation primitives. |
| `src/pivoteer/table_resizer.py` | file | Calculates and applies updated table `ref` ranges from data shape. |
| `src/pivoteer/pivot_cache_updater.py` | file | Optional pivot cache field synchronization for updated tables. |
| `src/pivoteer/models.py` | file | Immutable workbook/table metadata dataclasses shared across orchestration code. |
| `src/pivoteer/utils.py` | file | A1 reference conversion and parsing helpers. |
| `src/pivoteer/exceptions.py` | file | Domain-specific exception hierarchy used across the package. |

## Key Symbols

| Symbol | Kind | Visibility | Location | Purpose |
|--------|------|------------|----------|---------|
| `__version__` | const | public | `src/pivoteer/__init__.py:7` | Exposes package version metadata. |
| `__all__` | const | public | `src/pivoteer/__init__.py:3` | Declares explicit export list for package metadata. |
| `LOGGER` | const | internal | `src/pivoteer/core.py:13` | Emits save-level runtime logs from public API. |
| `Pivoteer` | class | public | `src/pivoteer/core.py:16` | Main user-facing facade for template update workflow. |
| `Pivoteer.__init__` | function | public | `src/pivoteer/core.py:19` | Creates `TemplateEngine` and stores pivot sync feature flag. |
| `Pivoteer.apply_dataframe` | function | public | `src/pivoteer/core.py:26` | Delegates table-targeted DataFrame injection to engine. |
| `Pivoteer.save` | function | public | `src/pivoteer/core.py:30` | Applies pivot updates and writes modified/new XML parts to output workbook. |
| `PivoteerError` | class | public | `src/pivoteer/exceptions.py:4` | Base exception for package-specific failures. |
| `TemplateNotFoundError` | class | public | `src/pivoteer/exceptions.py:8` | Signals missing/unopenable workbook template path. |
| `TableNotFoundError` | class | public | `src/pivoteer/exceptions.py:12` | Signals unresolved target table name in workbook map. |
| `XmlStructureError` | class | public | `src/pivoteer/exceptions.py:16` | Signals malformed/missing XML structures or parts. |
| `InvalidDataError` | class | public | `src/pivoteer/exceptions.py:20` | Signals invalid data payloads or coordinate inputs. |
| `WriteError` | class | public | `src/pivoteer/exceptions.py:24` | Reserved write failure domain exception type. |
| `PivotCacheError` | class | public | `src/pivoteer/exceptions.py:28` | Signals pivot cache metadata update failures. |
| `TableRef` | class | public | `src/pivoteer/models.py:10` | Immutable descriptor linking table name, paths, and current range ref. |
| `WorksheetInfo` | class | public | `src/pivoteer/models.py:21` | Immutable worksheet metadata extracted from workbook relations. |
| `WorkbookMap` | class | public | `src/pivoteer/models.py:31` | Aggregated mapping of worksheets, tables, and pivot cache definition paths. |
| `_NS_MAIN` | const | internal | `src/pivoteer/pivot_cache_updater.py:13` | Default SpreadsheetML namespace fallback for pivot cache parsing. |
| `sync_cache_fields` | function | public | `src/pivoteer/pivot_cache_updater.py:16` | Finds matching pivot caches for a table and appends missing cache fields. |
| `_extract_table_columns` | function | internal | `src/pivoteer/pivot_cache_updater.py:47` | Reads table column names from table XML metadata. |
| `_cache_source_table_name` | function | internal | `src/pivoteer/pivot_cache_updater.py:65` | Resolves pivot cache source table name from `worksheetSource`. |
| `_append_missing_cache_fields` | function | internal | `src/pivoteer/pivot_cache_updater.py:75` | Appends missing `<cacheField>` nodes and updates cache field counts. |
| `_get_main_namespace` | function | internal | `src/pivoteer/pivot_cache_updater.py:107` | Resolves active/default namespace from XML root map. |
| `TableResizeResult` | class | public | `src/pivoteer/table_resizer.py:14` | Captures before/after table range values after resize operation. |
| `TableResizer` | class | public | `src/pivoteer/table_resizer.py:21` | Encapsulates table `ref` recomputation and mutation logic. |
| `TableResizer.resize_table` | function | public | `src/pivoteer/table_resizer.py:24` | Recalculates end row/column from data shape and writes new `ref`. |
| `LOGGER` | const | internal | `src/pivoteer/template_engine.py:19` | Emits orchestration-level diagnostics. |
| `TemplateEngine` | class | public | `src/pivoteer/template_engine.py:22` | Coordinates workbook map resolution, XML updates, and staged modifications. |
| `TemplateEngine.__init__` | function | public | `src/pivoteer/template_engine.py:25` | Instantiates XML/resizer services and initial workbook/table state. |
| `TemplateEngine.template_path` | function | public | `src/pivoteer/template_engine.py:34` | Returns current template workbook path. |
| `TemplateEngine.apply_dataframe` | function | public | `src/pivoteer/template_engine.py:37` | Injects rows into sheet XML and resizes matching table reference. |
| `TemplateEngine.ensure_pivot_refresh_on_load` | function | public | `src/pivoteer/template_engine.py:81` | Sets `refreshOnLoad` on all discovered pivot cache definitions. |
| `TemplateEngine.sync_pivot_cache_fields` | function | public | `src/pivoteer/template_engine.py:94` | Applies optional pivot cache schema alignment for updated tables only. |
| `TemplateEngine.get_modified_parts` | function | public | `src/pivoteer/template_engine.py:104` | Serializes staged XML trees into ZIP-part byte payloads. |
| `TemplateEngine._read_xml_part` | function | internal | `src/pivoteer/template_engine.py:113` | Reads XML part from staged cache or archive source. |
| `_A1_RE` | const | internal | `src/pivoteer/utils.py:7` | Regex for validating single-cell A1 notation. |
| `_A1_RANGE_RE` | const | internal | `src/pivoteer/utils.py:8` | Regex for validating A1 range notation. |
| `column_index_to_letter` | function | public | `src/pivoteer/utils.py:11` | Converts 1-based numeric column index to Excel letters. |
| `column_letter_to_index` | function | public | `src/pivoteer/utils.py:24` | Converts Excel column letters to 1-based numeric index. |
| `parse_a1_cell` | function | public | `src/pivoteer/utils.py:35` | Parses A1 cell into `(row, col)` tuple coordinates. |
| `parse_a1_range` | function | public | `src/pivoteer/utils.py:45` | Parses A1 range into start/end coordinate tuples. |
| `build_a1_cell` | function | public | `src/pivoteer/utils.py:55` | Builds A1 cell reference from 1-based row/column coordinates. |
| `build_a1_range` | function | public | `src/pivoteer/utils.py:62` | Builds A1 range string from start/end coordinate pairs. |
| `LOGGER` | const | internal | `src/pivoteer/xml_engine.py:21` | Emits low-level XML injection and validation warnings. |
| `_NS_MAIN` | const | internal | `src/pivoteer/xml_engine.py:23` | SpreadsheetML namespace constant. |
| `_NS_REL` | const | internal | `src/pivoteer/xml_engine.py:24` | Office document relationship namespace constant. |
| `_NS_PKG_REL` | const | internal | `src/pivoteer/xml_engine.py:25` | Package relationship namespace constant. |
| `_NSMAP_MAIN` | const | internal | `src/pivoteer/xml_engine.py:27` | Namespace map for worksheet/table XPath queries. |
| `_NSMAP_REL` | const | internal | `src/pivoteer/xml_engine.py:28` | Namespace map for office relationships. |
| `_NSMAP_PKG` | const | internal | `src/pivoteer/xml_engine.py:29` | Namespace map for package relationship parts. |
| `read_xml_part` | function | public | `src/pivoteer/xml_engine.py:32` | Shared helper to read and parse XML ZIP part by path. |
| `XmlEngine` | class | public | `src/pivoteer/xml_engine.py:42` | Core low-level service for workbook map extraction and XML edits. |
| `XmlEngine.__init__` | function | public | `src/pivoteer/xml_engine.py:45` | Validates template existence and stores template path. |
| `XmlEngine.template_path` | function | public | `src/pivoteer/xml_engine.py:51` | Exposes template path for callers. |
| `XmlEngine.build_workbook_map` | function | public | `src/pivoteer/xml_engine.py:54` | Builds worksheet/table/pivot cache lookup model from workbook parts. |
| `XmlEngine.read_sheet_xml` | function | public | `src/pivoteer/xml_engine.py:71` | Reads worksheet XML tree by worksheet part path. |
| `XmlEngine.write_sheet_xml` | function | public | `src/pivoteer/xml_engine.py:77` | Serializes and writes worksheet XML back into archive. |
| `XmlEngine.inject_rows_inline_strings` | function | public | `src/pivoteer/xml_engine.py:86` | Mutates sheet rows/cells with typed value encoding rules. |
| `XmlEngine._parse_worksheets` | function | internal | `src/pivoteer/xml_engine.py:116` | Resolves workbook sheet metadata and worksheet paths via rel IDs. |
| `XmlEngine._parse_tables` | function | internal | `src/pivoteer/xml_engine.py:143` | Resolves table parts and constructs table lookup entries. |
| `XmlEngine._parse_pivot_caches` | function | internal | `src/pivoteer/xml_engine.py:186` | Detects pivot cache definition targets from workbook relationships. |
| `XmlEngine.read_xml` | function | public | `src/pivoteer/xml_engine.py:194` | Public method wrapper over shared `read_xml_part`. |
| `XmlEngine._read_xml` | function | internal | `src/pivoteer/xml_engine.py:198` | Backward-compatible alias for `read_xml`. |
| `XmlEngine._write_xml` | function | internal | `src/pivoteer/xml_engine.py:200` | Writes XML bytes into ZIP archive path. |
| `XmlEngine._parse_relationships` | function | internal | `src/pivoteer/xml_engine.py:203` | Parses package relationship entries into ID-target map. |
| `XmlEngine._sheet_rels_path` | function | internal | `src/pivoteer/xml_engine.py:213` | Computes worksheet relationship part path from worksheet path. |
| `XmlEngine._normalize_rel_target` | function | internal | `src/pivoteer/xml_engine.py:218` | Normalizes relative relationship targets to workbook ZIP paths. |
| `XmlEngine._find_or_create_row` | function | internal | `src/pivoteer/xml_engine.py:225` | Returns existing row node or creates one for target row index. |
| `XmlEngine._find_or_create_cell` | function | internal | `src/pivoteer/xml_engine.py:235` | Returns existing cell node or creates one for target cell reference. |
| `XmlEngine._set_cell_value_inline` | function | internal | `src/pivoteer/xml_engine.py:245` | Encodes values as empty/numeric/inline string cell representations. |
| `XmlEngine._is_missing` | function | internal | `src/pivoteer/xml_engine.py:267` | Detects missing-like values via pandas NA semantics. |
| `XmlEngine._sort_rows` | function | internal | `src/pivoteer/xml_engine.py:273` | Reorders `<row>` nodes numerically by row index after mutations. |

## Data Flow

1. Caller invokes `Pivoteer.apply_dataframe(table_name, df)` and `Pivoteer.save(...)`.
2. `TemplateEngine.apply_dataframe` validates input and resolves table metadata from
   `WorkbookMap.tables`.
3. `XmlEngine.inject_rows_inline_strings` mutates worksheet `sheetData` rows/cells.
4. `TableResizer.resize_table` recalculates and mutates table `ref` range.
5. `TemplateEngine.ensure_pivot_refresh_on_load` marks pivot cache definitions with
   `refreshOnLoad="1"`.
6. Optional `TemplateEngine.sync_pivot_cache_fields` appends missing cache fields for
   updated tables.
7. `TemplateEngine.get_modified_parts` serializes changed XML trees to bytes.
8. `Pivoteer.save` copies all ZIP parts from template, replacing only modified parts.

## Configuration

- Runtime flag: `enable_pivot_field_sync` in `Pivoteer.__init__` controls whether
  pivot cache field synchronization runs before save.
- No environment variables are used by runtime module code.
- Package/runtime compatibility and dependencies are defined in `pyproject.toml`.

## Inventory Notes

- **Coverage**: full
- **Notes**: Inventory covers all runtime source files under `src/pivoteer` and all
  top-level module symbols plus class methods/constants that participate in external
  behavior or internal data flow. Auto-generated cache directories are intentionally
  excluded.
