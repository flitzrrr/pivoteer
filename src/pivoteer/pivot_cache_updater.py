"""Pivot cache field synchronization for Excel tables."""

from __future__ import annotations

import zipfile

from lxml import etree

from pivoteer.exceptions import PivotCacheError, TableNotFoundError
from pivoteer.models import WorkbookMap
from pivoteer.xml_engine import read_xml_part

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def sync_cache_fields(
    workbook_map: WorkbookMap, table_name: str
) -> dict[str, etree._ElementTree]:
    """Sync pivot cache field names with the specified table's columns.

    Returns a mapping of modified pivot cache definition paths to XML trees.
    """
    table_ref = workbook_map.tables.get(table_name)
    if not table_ref:
        raise TableNotFoundError(f"Table not found: {table_name}")

    cache_paths = list(workbook_map.pivot_cache_definition_paths.values())
    if not cache_paths:
        return {}

    with zipfile.ZipFile(workbook_map.template_path, "r") as archive:
        table_tree = read_xml_part(archive, table_ref.table_path)
        table_columns = _extract_table_columns(table_tree)

        updated_parts: dict[str, etree._ElementTree] = {}
        for cache_path in cache_paths:
            cache_tree = read_xml_part(archive, cache_path)
            if _cache_source_table_name(cache_tree) != table_name:
                continue
            updated = _append_missing_cache_fields(cache_tree, table_columns)
            if updated:
                updated_parts[cache_path] = cache_tree

        return updated_parts


def _extract_table_columns(table_tree: etree._ElementTree) -> list[str]:
    ns = _get_main_namespace(table_tree)
    table_columns = table_tree.find(f".//{{{ns}}}tableColumns")
    if table_columns is None:
        raise PivotCacheError("Table columns element not found.")

    columns: list[str] = []
    for column in table_columns.findall(f"{{{ns}}}tableColumn"):
        name = column.get("name")
        if name:
            columns.append(name)

    if not columns:
        msg = "Table columns are empty; cannot sync pivot cache fields."
        raise PivotCacheError(msg)
    return columns


def _cache_source_table_name(cache_tree: etree._ElementTree) -> str | None:
    ns = _get_main_namespace(cache_tree)
    worksheet_source = cache_tree.find(
        f".//{{{ns}}}cacheSource/{{{ns}}}worksheetSource"
    )
    if worksheet_source is None:
        return None
    return worksheet_source.get("name")


def _append_missing_cache_fields(
    cache_tree: etree._ElementTree, table_columns: list[str]
) -> bool:
    ns = _get_main_namespace(cache_tree)
    cache_fields = cache_tree.find(f".//{{{ns}}}cacheFields")
    if cache_fields is None:
        msg = "cacheFields element not found in pivot cache definition."
        raise PivotCacheError(msg)

    existing_fields = [
        field.get("name")
        for field in cache_fields.findall(f"{{{ns}}}cacheField")
        if field.get("name")
    ]

    missing = [name for name in table_columns if name not in existing_fields]
    if not missing:
        return False

    for name in missing:
        new_field = etree.SubElement(cache_fields, f"{{{ns}}}cacheField")
        new_field.set("name", name)
        shared_items = etree.SubElement(new_field, f"{{{ns}}}sharedItems")
        shared_items.set("count", "0")

    cache_fields.set(
        "count",
        str(len(cache_fields.findall(f"{{{ns}}}cacheField"))),
    )
    return True


def _get_main_namespace(tree: etree._ElementTree) -> str:
    root = tree.getroot()
    ns = root.nsmap.get(None) or root.nsmap.get("main")
    return ns or _NS_MAIN
