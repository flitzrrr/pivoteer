"""Table resizing logic for Excel ListObjects."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Tuple

from lxml import etree

from pivoteer.exceptions import XmlStructureError
from pivoteer.utils import build_a1_range, parse_a1_range


@dataclass(frozen=True)
class TableResizeResult:
    """Result of a table resize operation."""

    original_ref: str
    updated_ref: str


class TableResizer:
    """Calculates and applies table range updates."""

    def resize_table(
        self,
        table_tree: etree._ElementTree,
        data_rows: int,
        data_cols: int,
    ) -> TableResizeResult:
        if data_rows < 0 or data_cols < 1:
            raise ValueError("Data rows must be >= 0 and data cols >= 1.")

        table = table_tree.getroot()
        ref = table.get("ref")
        if not ref:
            raise XmlStructureError("Table ref attribute missing.")

        (start_row, start_col), _ = parse_a1_range(ref)
        header_rows = 1
        end_row = start_row + header_rows + data_rows - 1
        end_col = start_col + data_cols - 1

        updated_ref = build_a1_range(start_row, start_col, end_row, end_col)
        table.set("ref", updated_ref)
        return TableResizeResult(original_ref=ref, updated_ref=updated_ref)
