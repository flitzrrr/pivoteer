"""Integration test to verify resized table range in output XLSX."""

from __future__ import annotations

import posixpath
import zipfile
from pathlib import Path
from typing import Dict

from lxml import etree

from pivoteer.utils import parse_a1_range
from pivoteer.core import Pivoteer


_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

_NSMAP_MAIN = {"main": _NS_MAIN}
_NSMAP_PKG = {"rel": _NS_PKG_REL}


def _read_xml(archive: zipfile.ZipFile, path: str) -> etree._ElementTree:
    data = archive.read(path)
    parser = etree.XMLParser(remove_blank_text=False)
    return etree.fromstring(data, parser).getroottree()


def _parse_relationships(rels_tree: etree._ElementTree) -> Dict[str, str]:
    rels: Dict[str, str] = {}
    rel_nodes = rels_tree.findall(".//rel:Relationship", namespaces=_NSMAP_PKG)
    for rel in rel_nodes:
        rel_id = rel.get("Id")
        target = rel.get("Target")
        if rel_id and target:
            rels[rel_id] = target
    return rels


def _resolve_table_path(archive: zipfile.ZipFile, table_name: str) -> str:
    workbook_tree = _read_xml(archive, "xl/workbook.xml")
    rels_tree = _read_xml(archive, "xl/_rels/workbook.xml.rels")
    rel_map = _parse_relationships(rels_tree)

    sheet_nodes = workbook_tree.findall(".//main:sheets/main:sheet", _NSMAP_MAIN)
    for sheet in sheet_nodes:
        rel_id = sheet.get(f"{{{_NS_REL}}}id")
        if not rel_id:
            continue
        worksheet_target = rel_map.get(rel_id)
        if not worksheet_target:
            continue

        worksheet_path = f"xl/{worksheet_target}"
        rels_path = Path(worksheet_path).parent / "_rels" / f"{Path(worksheet_path).name}.rels"
        rels_tree = _read_xml(archive, str(rels_path))
        worksheet_rel_map = _parse_relationships(rels_tree)

        sheet_tree = _read_xml(archive, worksheet_path)
        table_parts = sheet_tree.findall(
            ".//main:tableParts/main:tablePart", _NSMAP_MAIN
        )
        for table_part in table_parts:
            table_rel_id = table_part.get(f"{{{_NS_REL}}}id")
            if not table_rel_id:
                continue
            target = worksheet_rel_map.get(table_rel_id)
            if not target:
                continue

            base_dir = posixpath.dirname(worksheet_path)
            table_path = posixpath.normpath(posixpath.join(base_dir, target))
            if not table_path.startswith("xl/"):
                table_path = f"xl/{table_path.lstrip('./')}"

            table_tree = _read_xml(archive, table_path)
            table_node = table_tree.getroot()
            name = table_node.get("name")
            if name == table_name:
                return table_path

    raise RuntimeError(f"Table {table_name!r} not found.")


def _build_output_report(template_path: Path, output_path: Path) -> None:
    import pandas as pd
    from datetime import date, timedelta

    data = []
    base = date(2024, 1, 1)
    for i in range(500):
        data.append(
            {
                "Category": "Hardware",
                "Region": "North",
                "Amount": 100.0 + (i * 10.0),
                "Date": base + timedelta(days=i),
            }
        )

    df = pd.DataFrame(data)
    pivoteer = Pivoteer(template_path)
    pivoteer.apply_dataframe("DataSource", df)
    pivoteer.save(output_path)


def test_table_range_resized(template_path: Path, tmp_path: Path) -> None:
    report_path = tmp_path / "report_output.xlsx"
    _build_output_report(template_path, report_path)

    with zipfile.ZipFile(report_path, "r") as archive:
        table_path = _resolve_table_path(archive, "DataSource")
        table_tree = _read_xml(archive, table_path)
        table_node = table_tree.getroot()
        ref = table_node.get("ref")
        if not ref:
            raise RuntimeError("Table ref attribute missing.")

    (start_row, _), (end_row, _) = parse_a1_range(ref)
    row_count = end_row - start_row + 1

    assert row_count >= 501, (
        f"Expected at least 501 rows (header + 500 data). Got {row_count}."
    )


def test_pivot_refresh_flag_optional(template_path: Path, tmp_path: Path) -> None:
    report_path = tmp_path / "report_output.xlsx"
    _build_output_report(template_path, report_path)

    pivot_refresh_enabled = False
    pivot_cache_found = False
    with zipfile.ZipFile(report_path, "r") as archive:
        for filename in archive.namelist():
            if filename.startswith(
                "xl/pivotCache/pivotCacheDefinition"
            ) and filename.endswith(".xml"):
                pivot_cache_found = True
                pivot_tree = _read_xml(archive, filename)
                root = pivot_tree.getroot()
                if root.get("refreshOnLoad") == "1":
                    pivot_refresh_enabled = True
                    break

    if pivot_cache_found:
        assert pivot_refresh_enabled, "Pivot cache refreshOnLoad is not enabled."
