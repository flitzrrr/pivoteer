"""Data models for workbook metadata and table references."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class TableRef:
    """Represents a table (ListObject) and its XML references."""

    name: str
    sheet_name: str
    table_path: str
    worksheet_path: str
    ref: str


@dataclass(frozen=True)
class WorksheetInfo:
    """Represents worksheet metadata from workbook relationships."""

    name: str
    sheet_id: str
    path: str
    rel_id: str


@dataclass(frozen=True)
class WorkbookMap:
    """Holds resolved workbook references used for XML updates."""

    template_path: Path
    worksheets: dict[str, WorksheetInfo]
    tables: dict[str, TableRef]
    pivot_cache_definition_paths: dict[str, str]
    shared_strings_path: str | None = None
