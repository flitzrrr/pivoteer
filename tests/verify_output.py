"""Verify that report_output.xlsx contains resized table range."""

from __future__ import annotations

import logging
import posixpath
import zipfile
from pathlib import Path
from typing import Dict

from lxml import etree

from pivoteer.utils import parse_a1_range


LOGGER = logging.getLogger(__name__)

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


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    report_path = Path("report_output.xlsx")

    if not report_path.exists():
        raise FileNotFoundError("report_output.xlsx not found. Run demo_runner.py first.")

    with zipfile.ZipFile(report_path, "r") as archive:
        table_path = _resolve_table_path(archive, "DataSource")
        table_tree = _read_xml(archive, table_path)
        table_node = table_tree.getroot()
        ref = table_node.get("ref")
        if not ref:
            raise RuntimeError("Table ref attribute missing.")

    (start_row, _), (end_row, _) = parse_a1_range(ref)
    row_count = end_row - start_row + 1

    print(f"Table ref: {ref}")
    assert row_count >= 501, (
        f"Expected at least 501 rows (header + 500 data). Got {row_count}."
    )
    LOGGER.info("Verified table range includes injected data (%s rows).", row_count)

    pivot_refresh_enabled = False
    pivot_cache_found = False
    with zipfile.ZipFile(report_path, "r") as archive:
        for filename in archive.namelist():
            if filename.startswith("xl/pivotCache/pivotCacheDefinition") and filename.endswith(".xml"):
                pivot_cache_found = True
                pivot_tree = _read_xml(archive, filename)
                root = pivot_tree.getroot()
                if root.get("refreshOnLoad") == "1":
                    pivot_refresh_enabled = True
                    break

    if not pivot_cache_found:
        LOGGER.warning(
            "Skipping Pivot Refresh check (no cache found in generated template)."
        )
    elif pivot_refresh_enabled:
        print("Pivot Cache Refresh: ENABLED")
    else:
        raise RuntimeError("Pivot cache refreshOnLoad is not enabled.")


if __name__ == "__main__":
    main()
