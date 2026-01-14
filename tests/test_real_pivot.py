"""Integration test for real pivot cache preservation."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd
import pytest
from lxml import etree

from pivoteer.core import Pivoteer


def test_real_pivot_refresh_flag(tmp_path: Path) -> None:
    fixture_path = Path("tests/fixtures/real_pivot.xlsx")
    if not fixture_path.exists():
        pytest.skip("Fixture not found: tests/fixtures/real_pivot.xlsx")

    df = pd.DataFrame(
        {
            "Spalte1": ["A", "B", "C", "D", "E"],
            "Spalte2": [1, 2, 3, 4, 5],
        }
    )

    output_path = tmp_path / "real_pivot_output.xlsx"
    pivoteer = Pivoteer(fixture_path)
    pivoteer.apply_dataframe("TestTable", df)
    pivoteer.save(output_path)

    pivot_definition_path = None
    refresh_flag = None
    with zipfile.ZipFile(output_path, "r") as archive:
        for filename in archive.namelist():
            if "pivotCacheDefinition" in filename and filename.endswith(".xml"):
                pivot_definition_path = filename
                tree = etree.fromstring(archive.read(filename)).getroottree()
                root = tree.getroot()
                refresh_flag = root.get("refreshOnLoad")
                break

    print(f"Pivot cache path: {pivot_definition_path}")
    print(f"refreshOnLoad: {refresh_flag}")

    assert pivot_definition_path is not None, "No pivot cache definition found."
    assert refresh_flag == "1", "Pivot cache refreshOnLoad was not set to '1'."
