"""Low-level XML and ZIP handling for pivoteer."""

from __future__ import annotations

import logging
import posixpath
import zipfile
from pathlib import Path

import pandas as pd
from lxml import etree

from pivoteer.exceptions import (
    InvalidDataError,
    TemplateNotFoundError,
    XmlStructureError,
)
from pivoteer.models import TableRef, WorkbookMap, WorksheetInfo
from pivoteer.utils import build_a1_cell

LOGGER = logging.getLogger(__name__)

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

_NSMAP_MAIN = {"main": _NS_MAIN}
_NSMAP_REL = {"rel": _NS_REL}
_NSMAP_PKG = {"rel": _NS_PKG_REL}


def read_xml_part(archive: zipfile.ZipFile, path: str) -> etree._ElementTree:
    """Read an XML part from a ZIP archive and return an lxml ElementTree."""
    try:
        data = archive.read(path)
    except KeyError as exc:
        raise XmlStructureError(f"Missing XML part: {path}") from exc
    parser = etree.XMLParser(remove_blank_text=False)
    return etree.fromstring(data, parser).getroottree()


class XmlEngine:
    """Provides ZIP IO and XML manipulation for Excel workbooks."""

    def __init__(self, template_path: Path) -> None:
        if not template_path.exists():
            raise TemplateNotFoundError(f"Template not found: {template_path}")
        self._template_path = template_path

    @property
    def template_path(self) -> Path:
        return self._template_path

    def build_workbook_map(self) -> WorkbookMap:
        """Build a map of worksheets, tables, and pivot caches."""
        with zipfile.ZipFile(self._template_path, "r") as archive:
            workbook_tree = self._read_xml(archive, "xl/workbook.xml")
            rels_tree = self._read_xml(archive, "xl/_rels/workbook.xml.rels")

            worksheets = self._parse_worksheets(workbook_tree, rels_tree)
            tables = self._parse_tables(archive, worksheets)
            pivot_cache_paths = self._parse_pivot_caches(rels_tree)

        return WorkbookMap(
            template_path=self._template_path,
            worksheets=worksheets,
            tables=tables,
            pivot_cache_definition_paths=pivot_cache_paths,
        )

    def read_sheet_xml(
        self, archive: zipfile.ZipFile, worksheet_path: str
    ) -> etree._ElementTree:
        """Read worksheet XML as an lxml tree."""
        return self._read_xml(archive, worksheet_path)

    def write_sheet_xml(
        self, archive: zipfile.ZipFile, worksheet_path: str, tree: etree._ElementTree
    ) -> None:
        """Write worksheet XML back into the archive."""
        xml_bytes = etree.tostring(
            tree, encoding="UTF-8", xml_declaration=True, standalone="yes"
        )
        self._write_xml(archive, worksheet_path, xml_bytes)

    def inject_rows_inline_strings(
        self,
        tree: etree._ElementTree,
        start_row: int,
        start_col: int,
        rows: list[list[object]],
    ) -> None:
        """Inject data rows into sheetData using inline strings for text."""
        if start_row < 1 or start_col < 1:
            raise InvalidDataError("Start row/col must be >= 1.")
        if not rows:
            LOGGER.warning("No rows provided for injection; worksheet left unchanged.")
            return

        sheet_data = tree.find(".//main:sheetData", namespaces=_NSMAP_MAIN)
        if sheet_data is None:
            raise XmlStructureError("sheetData element not found.")

        for row_offset, row_values in enumerate(rows):
            row_idx = start_row + row_offset
            row_element = self._find_or_create_row(sheet_data, row_idx)

            for col_offset, value in enumerate(row_values):
                col_idx = start_col + col_offset
                cell_ref = build_a1_cell(row_idx, col_idx)
                cell = self._find_or_create_cell(row_element, cell_ref)
                self._set_cell_value_inline(cell, value)

        self._sort_rows(sheet_data)

    def _parse_worksheets(
        self,
        workbook_tree: etree._ElementTree,
        rels_tree: etree._ElementTree,
    ) -> dict[str, WorksheetInfo]:
        sheet_nodes = workbook_tree.findall(".//main:sheets/main:sheet", _NSMAP_MAIN)
        rel_map = self._parse_relationships(rels_tree)

        worksheets: dict[str, WorksheetInfo] = {}
        for sheet in sheet_nodes:
            name = sheet.get("name")
            sheet_id = sheet.get("sheetId")
            rel_id = sheet.get(f"{{{_NS_REL}}}id")
            if not name or not sheet_id or not rel_id:
                raise XmlStructureError("Worksheet metadata is incomplete.")

            target = rel_map.get(rel_id)
            if not target:
                raise XmlStructureError(f"Missing relationship for sheet {name}.")
            path = f"xl/{target}"

            worksheets[name] = WorksheetInfo(
                name=name, sheet_id=sheet_id, path=path, rel_id=rel_id
            )

        return worksheets

    def _parse_tables(
        self,
        archive: zipfile.ZipFile,
        worksheets: dict[str, WorksheetInfo],
    ) -> dict[str, TableRef]:
        tables: dict[str, TableRef] = {}
        for worksheet in worksheets.values():
            sheet_tree = self._read_xml(archive, worksheet.path)
            table_parts = sheet_tree.findall(
                ".//main:tableParts/main:tablePart", _NSMAP_MAIN
            )
            if not table_parts:
                continue

            rels_path = self._sheet_rels_path(worksheet.path)
            rels_tree = self._read_xml(archive, rels_path)
            rel_map = self._parse_relationships(rels_tree)

            for table_part in table_parts:
                rel_id = table_part.get(f"{{{_NS_REL}}}id")
                if not rel_id:
                    continue
                target = rel_map.get(rel_id)
                if not target:
                    continue
                table_path = self._normalize_rel_target(worksheet.path, target)
                table_tree = self._read_xml(archive, table_path)
                table_node = table_tree.getroot()
                name = table_node.get("name")
                ref = table_node.get("ref")
                if not name or not ref:
                    raise XmlStructureError("Table definition missing name or ref.")

                tables[name] = TableRef(
                    name=name,
                    sheet_name=worksheet.name,
                    table_path=table_path,
                    worksheet_path=worksheet.path,
                    ref=ref,
                )

        return tables

    def _parse_pivot_caches(self, rels_tree: etree._ElementTree) -> dict[str, str]:
        rel_map = self._parse_relationships(rels_tree)
        cache_paths: dict[str, str] = {}
        for rel_id, target in rel_map.items():
            if "pivotCache" in target:
                cache_paths[rel_id] = f"xl/{target}"
        return cache_paths

    def read_xml(self, archive: zipfile.ZipFile, path: str) -> etree._ElementTree:
        return read_xml_part(archive, path)

    # Backward-compatible alias
    _read_xml = read_xml

    def _write_xml(self, archive: zipfile.ZipFile, path: str, data: bytes) -> None:
        archive.writestr(path, data)

    def _parse_relationships(self, rels_tree: etree._ElementTree) -> dict[str, str]:
        rels: dict[str, str] = {}
        rel_nodes = rels_tree.findall(".//rel:Relationship", _NSMAP_PKG)
        for rel in rel_nodes:
            rel_id = rel.get("Id")
            target = rel.get("Target")
            if rel_id and target:
                rels[rel_id] = target
        return rels

    def _sheet_rels_path(self, worksheet_path: str) -> str:
        filename = Path(worksheet_path).name
        parent = Path(worksheet_path).parent
        return str(parent / "_rels" / f"{filename}.rels")

    def _normalize_rel_target(self, worksheet_path: str, target: str) -> str:
        base_dir = posixpath.dirname(worksheet_path)
        normalized = posixpath.normpath(posixpath.join(base_dir, target))
        if not normalized.startswith("xl/"):
            normalized = f"xl/{normalized.lstrip('./')}"
        return normalized

    def _find_or_create_row(
        self, sheet_data: etree._Element, row_idx: int
    ) -> etree._Element:
        row = sheet_data.find(f"main:row[@r='{row_idx}']", namespaces=_NSMAP_MAIN)
        if row is not None:
            return row
        row = etree.SubElement(sheet_data, f"{{{_NS_MAIN}}}row")
        row.set("r", str(row_idx))
        return row

    def _find_or_create_cell(
        self, row: etree._Element, cell_ref: str
    ) -> etree._Element:
        cell = row.find(f"main:c[@r='{cell_ref}']", namespaces=_NSMAP_MAIN)
        if cell is not None:
            return cell
        cell = etree.SubElement(row, f"{{{_NS_MAIN}}}c")
        cell.set("r", cell_ref)
        return cell

    def _set_cell_value_inline(self, cell: etree._Element, value: object) -> None:
        for child in list(cell):
            cell.remove(child)

        if value is None or self._is_missing(value):
            cell.attrib.pop("t", None)
            cell.attrib.pop("s", None)
            return

        if isinstance(value, (int, float)):
            cell.attrib.pop("t", None)
            v = etree.SubElement(cell, f"{{{_NS_MAIN}}}v")
            v.text = str(value)
            return

        text_value = value.isoformat() if hasattr(value, "isoformat") else str(value)

        cell.set("t", "inlineStr")
        inline = etree.SubElement(cell, f"{{{_NS_MAIN}}}is")
        text = etree.SubElement(inline, f"{{{_NS_MAIN}}}t")
        text.text = text_value

    def _is_missing(self, value: object) -> bool:
        try:
            return bool(pd.isna(value))
        except (TypeError, ValueError):
            return False

    def _sort_rows(self, sheet_data: etree._Element) -> None:
        rows = sheet_data.findall("main:row", namespaces=_NSMAP_MAIN)
        rows_sorted = sorted(rows, key=lambda elem: int(elem.get("r", "0")))
        for row in rows:
            sheet_data.remove(row)
        for row in rows_sorted:
            sheet_data.append(row)
