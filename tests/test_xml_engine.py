"""Unit tests for XmlEngine cell value injection and row handling."""

from __future__ import annotations

from datetime import date, datetime
from math import nan

import pandas as pd
import pytest
from lxml import etree

from pivoteer.exceptions import InvalidDataError, XmlStructureError
from pivoteer.xml_engine import XmlEngine, read_xml_part

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NSMAP = {"main": _NS_MAIN}


def _make_sheet_tree() -> etree._ElementTree:
    """Create a minimal worksheet tree with an empty sheetData."""
    xml = f'<worksheet xmlns="{_NS_MAIN}"><sheetData/></worksheet>'
    return etree.fromstring(xml.encode()).getroottree()


def _get_cell(tree: etree._ElementTree, cell_ref: str) -> etree._Element | None:
    return tree.find(f".//main:row/main:c[@r='{cell_ref}']", namespaces=_NSMAP)


class TestInjectRowsInlineStrings:
    def _make_engine(self, tmp_path) -> XmlEngine:
        # Create a minimal xlsx for XmlEngine init
        import zipfile

        path = tmp_path / "minimal.xlsx"
        workbook_xml = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<workbook xmlns="{_NS_MAIN}" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            "<sheets>"
            '<sheet name="Data" sheetId="1" r:id="rId1"/>'
            "</sheets>"
            "</workbook>"
        )
        rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/>'
            "</Relationships>"
        )
        worksheet_xml = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="{_NS_MAIN}"><sheetData/></worksheet>'
        )
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("xl/workbook.xml", workbook_xml)
            archive.writestr("xl/_rels/workbook.xml.rels", rels_xml)
            archive.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
        return XmlEngine(path)

    def test_integer_value(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        engine.inject_rows_inline_strings(tree, 1, 1, [[42]])
        cell = _get_cell(tree, "A1")
        assert cell is not None
        assert cell.get("t") is None  # numeric, no type attr
        v = cell.find(f"{{{_NS_MAIN}}}v")
        assert v is not None
        assert v.text == "42"

    def test_float_value(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        engine.inject_rows_inline_strings(tree, 1, 1, [[3.14]])
        cell = _get_cell(tree, "A1")
        assert cell is not None
        v = cell.find(f"{{{_NS_MAIN}}}v")
        assert v is not None
        assert v.text == "3.14"

    def test_string_value(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        engine.inject_rows_inline_strings(tree, 1, 1, [["hello"]])
        cell = _get_cell(tree, "A1")
        assert cell is not None
        assert cell.get("t") == "inlineStr"
        t = cell.find(f".//{{{_NS_MAIN}}}t")
        assert t is not None
        assert t.text == "hello"

    def test_none_value(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        engine.inject_rows_inline_strings(tree, 1, 1, [[None]])
        cell = _get_cell(tree, "A1")
        assert cell is not None
        assert cell.get("t") is None
        assert len(list(cell)) == 0  # no children

    def test_nan_value(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        engine.inject_rows_inline_strings(tree, 1, 1, [[nan]])
        cell = _get_cell(tree, "A1")
        assert cell is not None
        assert cell.get("t") is None
        assert len(list(cell)) == 0  # treated as missing

    def test_pd_nat_value(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        engine.inject_rows_inline_strings(tree, 1, 1, [[pd.NaT]])
        cell = _get_cell(tree, "A1")
        assert cell is not None
        assert len(list(cell)) == 0

    def test_date_value(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        engine.inject_rows_inline_strings(tree, 1, 1, [[date(2024, 6, 15)]])
        cell = _get_cell(tree, "A1")
        assert cell is not None
        assert cell.get("t") == "inlineStr"
        t = cell.find(f".//{{{_NS_MAIN}}}t")
        assert t is not None
        assert t.text == "2024-06-15"

    def test_datetime_value(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        engine.inject_rows_inline_strings(tree, 1, 1, [[datetime(2024, 6, 15, 10, 30)]])
        cell = _get_cell(tree, "A1")
        assert cell is not None
        t = cell.find(f".//{{{_NS_MAIN}}}t")
        assert t is not None
        assert "2024-06-15" in t.text

    def test_multiple_columns(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        engine.inject_rows_inline_strings(tree, 1, 1, [[1, "two", 3.0]])
        assert _get_cell(tree, "A1") is not None
        assert _get_cell(tree, "B1") is not None
        assert _get_cell(tree, "C1") is not None

    def test_multiple_rows(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        engine.inject_rows_inline_strings(tree, 2, 1, [["a"], ["b"], ["c"]])
        assert _get_cell(tree, "A2") is not None
        assert _get_cell(tree, "A3") is not None
        assert _get_cell(tree, "A4") is not None

    def test_zero_start_row_raises(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        with pytest.raises(InvalidDataError):
            engine.inject_rows_inline_strings(tree, 0, 1, [["x"]])

    def test_empty_rows_no_change(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        engine.inject_rows_inline_strings(tree, 1, 1, [])
        rows = tree.findall(".//main:row", namespaces=_NSMAP)
        assert len(rows) == 0

    def test_rows_are_sorted(self, tmp_path) -> None:
        engine = self._make_engine(tmp_path)
        tree = _make_sheet_tree()
        # Inject at row 5 then row 2 to test sorting
        engine.inject_rows_inline_strings(tree, 5, 1, [["late"]])
        engine.inject_rows_inline_strings(tree, 2, 1, [["early"]])
        rows = tree.findall(".//main:row", namespaces=_NSMAP)
        row_indices = [int(r.get("r", "0")) for r in rows]
        assert row_indices == sorted(row_indices)


class TestReadXmlPart:
    def test_missing_path_raises(self, tmp_path) -> None:
        import zipfile

        path = tmp_path / "empty.zip"
        with zipfile.ZipFile(path, "w"):
            pass

        with zipfile.ZipFile(path, "r") as archive:
            with pytest.raises(XmlStructureError, match="Missing XML part"):
                read_xml_part(archive, "nonexistent.xml")

    def test_reads_valid_xml(self, tmp_path) -> None:
        import zipfile

        path = tmp_path / "test.zip"
        xml_content = '<?xml version="1.0"?><root><child/></root>'
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("test.xml", xml_content)

        with zipfile.ZipFile(path, "r") as archive:
            tree = read_xml_part(archive, "test.xml")
            assert tree.getroot().tag == "root"
