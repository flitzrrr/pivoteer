"""High-level orchestration for XML updates."""

from __future__ import annotations

import logging
import zipfile
from pathlib import Path

import pandas as pd
from lxml import etree

from pivoteer.exceptions import InvalidDataError, TableNotFoundError
from pivoteer.models import TableRef, WorkbookMap
from pivoteer.pivot_cache_updater import sync_cache_fields
from pivoteer.table_resizer import TableResizer
from pivoteer.utils import parse_a1_range
from pivoteer.xml_engine import XmlEngine

LOGGER = logging.getLogger(__name__)


class TemplateEngine:
    """Coordinates XmlEngine and TableResizer for template updates."""

    def __init__(self, template_path: Path) -> None:
        self._xml_engine = XmlEngine(template_path)
        self._table_resizer = TableResizer()
        self._workbook_map: WorkbookMap = self._xml_engine.build_workbook_map()
        self._tables: dict[str, TableRef] = dict(self._workbook_map.tables)
        self._modified_trees: dict[str, etree._ElementTree] = {}
        self._updated_tables: set[str] = set()

    @property
    def template_path(self) -> Path:
        return self._xml_engine.template_path

    def apply_dataframe(self, table_name: str, df: pd.DataFrame) -> None:
        """Inject a DataFrame into the target table and resize it."""
        table_ref = self._tables.get(table_name)
        if not table_ref:
            raise TableNotFoundError(f"Table not found: {table_name}")
        if df.empty:
            raise InvalidDataError(
                f"Table '{table_name}' requires data, but DataFrame was empty."
            )
        if df.columns.empty:
            raise InvalidDataError(
                f"Table '{table_name}' requires columns, but DataFrame has none."
            )

        data_rows = df.itertuples(index=False, name=None)
        rows = [list(row) for row in data_rows]
        row_count = len(rows)
        col_count = len(df.columns)

        (start_row, start_col), _ = parse_a1_range(table_ref.ref)
        data_start_row = start_row + 1

        with zipfile.ZipFile(self.template_path, "r") as archive:
            sheet_tree = self._read_xml_part(archive, table_ref.worksheet_path)
            self._xml_engine.inject_rows_inline_strings(
                sheet_tree, data_start_row, start_col, rows
            )
            self._modified_trees[table_ref.worksheet_path] = sheet_tree

            table_tree = self._read_xml_part(archive, table_ref.table_path)
            resize_result = self._table_resizer.resize_table(
                table_tree, data_rows=row_count, data_cols=col_count
            )
            self._modified_trees[table_ref.table_path] = table_tree

        self._tables[table_name] = TableRef(
            name=table_ref.name,
            sheet_name=table_ref.sheet_name,
            table_path=table_ref.table_path,
            worksheet_path=table_ref.worksheet_path,
            ref=resize_result.updated_ref,
        )
        self._updated_tables.add(table_name)

    def ensure_pivot_refresh_on_load(self) -> None:
        """Set refreshOnLoad=1 for all pivot cache definitions."""
        pivot_paths = self._workbook_map.pivot_cache_definition_paths.values()
        if not pivot_paths:
            return

        with zipfile.ZipFile(self.template_path, "r") as archive:
            for path in pivot_paths:
                tree = self._read_xml_part(archive, path)
                root = tree.getroot()
                root.set("refreshOnLoad", "1")
                self._modified_trees[path] = tree

    def sync_pivot_cache_fields(self) -> None:
        """Append missing pivot cache fields for updated tables."""
        if not self._updated_tables:
            return

        for table_name in sorted(self._updated_tables):
            updated_parts = sync_cache_fields(self._workbook_map, table_name)
            for path, tree in updated_parts.items():
                self._modified_trees[path] = tree

    def get_modified_parts(self) -> dict[str, bytes]:
        """Serialize modified XML trees to bytes for writing."""
        parts: dict[str, bytes] = {}
        for path, tree in self._modified_trees.items():
            parts[path] = etree.tostring(
                tree, encoding="UTF-8", xml_declaration=True, standalone="yes"
            )
        return parts

    def _read_xml_part(self, archive: zipfile.ZipFile, path: str) -> etree._ElementTree:
        cached = self._modified_trees.get(path)
        if cached is not None:
            return cached
        return self._xml_engine.read_xml(archive, path)
