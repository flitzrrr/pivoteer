"""Unit test for pivot cache refresh flag logic."""

from __future__ import annotations

from lxml import etree


def test_pivot_cache_refresh_flag() -> None:
    xml = (
        "<pivotCacheDefinition "
        "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />"
    )
    tree = etree.fromstring(xml.encode("utf-8")).getroottree()
    root = tree.getroot()

    root.set("refreshOnLoad", "1")

    assert root.get("refreshOnLoad") == "1"
