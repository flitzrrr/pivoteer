# Changelog

## [0.2.1] - 2026-02-18

### Fixed

- Narrow exception handling in `_is_missing()` from bare `except Exception` to
  `except (TypeError, ValueError)` to avoid silencing critical errors

### Changed

- Extract `read_xml_part()` as public module-level function in `xml_engine.py`
  for reuse across modules
- Make `XmlEngine.read_xml()` public (backward-compatible `_read_xml` alias
  kept)
- Remove duplicated `_read_xml` from `pivot_cache_updater.py` in favour of
  shared `read_xml_part()`
- Remove unused imports across `core.py`, `template_engine.py`, and
  `xml_engine.py`

### Removed

- Delete obsolete `MANIFEST.in` (unused by hatchling build backend)

### Added

- Comprehensive unit tests for `utils.py` covering column index conversion,
  A1 cell/range parsing, round-trip consistency, and edge cases (columns 1, 26,
  27, 702, 703, 16384/XFD)
- Unit tests for `xml_engine.py` covering cell value injection for all data
  types (int, float, str, None, NaN, NaT, date, datetime), multi-row/column
  injection, row sorting, and `read_xml_part()`

## [0.2.0] - 2026-02-01

- Add opt-in pivot cache field synchronization for new table columns

## [0.1.1] - 2026-01-15

- Update PyPI project URLs
- Add tag-based PyPI release workflow (Trusted Publishing)
- Bump version to 0.1.1

## [0.1.0] - 2026-01-14

- Initial release
- Core XML Engine
- Table Resizer
