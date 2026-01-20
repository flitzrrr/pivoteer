"""Unit tests for pivot cache field synchronization."""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import List

from lxml import etree

from pivoteer.pivot_cache_updater import sync_cache_fields
from pivoteer.utils import column_index_to_letter
from pivoteer.xml_engine import XmlEngine


_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _write_minimal_xlsx(
    path: Path, table_columns: List[str], cache_fields: List[str]
) -> None:
    ref = f"A1:{column_index_to_letter(len(table_columns))}2"
    table_columns_xml = "".join(
        f'<tableColumn id="{idx}" name="{name}"/>'
        for idx, name in enumerate(table_columns, start=1)
    )
    cache_fields_xml = "".join(
        f'<cacheField name="{name}"/>'
        for name in cache_fields
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{_NS_MAIN}" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>"
        '<sheet name="Data" sheetId="1" r:id="rId1"/>'
        "</sheets>"
        "</workbook>"
    )

    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_NS_PKG_REL}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" '
        'Target="pivotCache/pivotCacheDefinition1.xml"/>'
        "</Relationships>"
    )

    worksheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{_NS_MAIN}" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheetData/>"
        '<tableParts count="1">'
        '<tablePart r:id="rId1"/>'
        "</tableParts>"
        "</worksheet>"
    )

    worksheet_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_NS_PKG_REL}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" '
        'Target="../tables/table1.xml"/>'
        "</Relationships>"
    )

    table_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<table xmlns="{_NS_MAIN}" id="1" name="DataSource" '
        f'displayName="DataSource" ref="{ref}" totalsRowShown="0">'
        f'<autoFilter ref="{ref}"/>'
        f'<tableColumns count="{len(table_columns)}">'
        f"{table_columns_xml}"
        "</tableColumns>"
        "</table>"
    )

    pivot_cache_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<pivotCacheDefinition xmlns="{_NS_MAIN}" refreshOnLoad="1">'
        '<cacheSource type="worksheet">'
        f'<worksheetSource name="DataSource" sheet="Data" ref="{ref}"/>'
        "</cacheSource>"
        f'<cacheFields count="{len(cache_fields)}">'
        f"{cache_fields_xml}"
        "</cacheFields>"
        "</pivotCacheDefinition>"
    )

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        archive.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
        archive.writestr("xl/worksheets/_rels/sheet1.xml.rels", worksheet_rels_xml)
        archive.writestr("xl/tables/table1.xml", table_xml)
        archive.writestr(
            "xl/pivotCache/pivotCacheDefinition1.xml", pivot_cache_xml
        )


def _main_namespace(tree: etree._ElementTree) -> str:
    root = tree.getroot()
    return root.nsmap.get(None) or root.nsmap.get("main") or _NS_MAIN


def _cache_fields(tree: etree._ElementTree) -> List[etree._Element]:
    ns = _main_namespace(tree)
    cache_fields = tree.find(f".//{{{ns}}}cacheFields")
    if cache_fields is None:
        return []
    return cache_fields.findall(f"{{{ns}}}cacheField")


def _cache_field_names(tree: etree._ElementTree) -> List[str]:
    return [
        node.get("name")
        for node in _cache_fields(tree)
        if node.get("name")
    ]


def test_cache_fields_unchanged_when_schema_matches(tmp_path: Path) -> None:
    fixture = tmp_path / "baseline.xlsx"
    columns = ["Category", "Region", "Amount"]
    _write_minimal_xlsx(fixture, columns, columns)

    workbook_map = XmlEngine(fixture).build_workbook_map()
    updates = sync_cache_fields(workbook_map, "DataSource")

    assert updates == {}


def test_cache_fields_appended_for_new_columns(tmp_path: Path) -> None:
    fixture = tmp_path / "new_column.xlsx"
    table_columns = ["Category", "Region", "Amount", "call_type"]
    cache_fields = ["Category", "Region", "Amount"]
    _write_minimal_xlsx(fixture, table_columns, cache_fields)

    workbook_map = XmlEngine(fixture).build_workbook_map()
    updates = sync_cache_fields(workbook_map, "DataSource")

    assert "xl/pivotCache/pivotCacheDefinition1.xml" in updates
    tree = updates["xl/pivotCache/pivotCacheDefinition1.xml"]
    cache_field_elements = _cache_fields(tree)
    assert _cache_field_names(tree) == table_columns
    assert len(cache_field_elements) == len(table_columns)
    ns = _main_namespace(tree)
    cache_fields_node = tree.find(f".//{{{ns}}}cacheFields")
    assert cache_fields_node is not None
    assert cache_fields_node.get("count") == str(len(table_columns))
