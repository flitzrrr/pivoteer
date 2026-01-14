"""Generate a dummy Excel template with a table, pivot table, and slicer."""

from __future__ import annotations

import argparse
import logging
from datetime import date, timedelta
from pathlib import Path
from typing import Iterable, List, Optional

import xlsxwriter


LOGGER = logging.getLogger(__name__)


def _seed_rows() -> List[List[object]]:
    categories = ["Hardware", "Software", "Services"]
    regions = ["North", "South", "East", "West"]
    base = date(2024, 1, 1)
    rows: List[List[object]] = []
    amount = 100.0

    for i in range(12):
        rows.append(
            [
                categories[i % len(categories)],
                regions[i % len(regions)],
                amount + (i * 25.0),
                base + timedelta(days=i),
            ]
        )
    return rows


def _write_table(
    worksheet: xlsxwriter.worksheet.Worksheet,
    headers: List[str],
    rows: List[List[object]],
) -> str:
    worksheet.write_row(0, 0, headers)
    for row_idx, row in enumerate(rows, start=1):
        worksheet.write_row(row_idx, 0, row)

    last_row = len(rows)
    last_col = len(headers) - 1

    table_range = xlsxwriter.utility.xl_range(0, 0, last_row, last_col)
    worksheet.add_table(
        table_range,
        {
            "name": "DataSource",
            "columns": [{"header": h} for h in headers],
        },
    )
    return table_range


def _add_pivot_table(
    worksheet: xlsxwriter.worksheet.Worksheet, data_range: str
) -> Optional[str]:
    pivot_table_name = "PivotTable1"
    if hasattr(worksheet, "add_pivot_table"):
        worksheet.add_pivot_table(
            {
                "data": data_range,
                "name": pivot_table_name,
                "rows": [{"data": "Category"}],
                "columns": [{"data": "Region"}],
                "values": [{"data": "Amount", "subtotal": "sum"}],
            }
        )
        return pivot_table_name

    LOGGER.warning("Pivot table API not available in installed xlsxwriter.")
    return None


def _add_slicer(
    worksheet: xlsxwriter.worksheet.Worksheet,
    pivot_table_name: Optional[str],
    field_name: str,
) -> None:
    if not pivot_table_name:
        return
    options = {"pivot_table": pivot_table_name, "field": field_name}

    if hasattr(worksheet, "insert_slicer"):
        try:
            worksheet.insert_slicer("H2", options)
            return
        except TypeError:
            pass

    if hasattr(worksheet, "add_slicer"):
        try:
            worksheet.add_slicer(options)
            return
        except TypeError:
            pass

    LOGGER.warning("Slicer API not available in installed xlsxwriter.")


def generate_template(output_path: Path) -> None:
    headers = ["Category", "Region", "Amount", "Date"]
    rows = _seed_rows()

    with xlsxwriter.Workbook(output_path) as workbook:
        data_sheet = workbook.add_worksheet("Data")
        pivot_sheet = workbook.add_worksheet("Pivot")

        data_range = _write_table(data_sheet, headers, rows)
        pivot_table_name = _add_pivot_table(pivot_sheet, data_range)
        _add_slicer(pivot_sheet, pivot_table_name, "Region")

    LOGGER.info("Template generated at %s", output_path)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate a dummy Excel template with table, pivot, slicer."
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("dummy_template.xlsx"),
        help="Output .xlsx path",
    )
    return parser.parse_args()


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    args = _parse_args()
    generate_template(args.output)


if __name__ == "__main__":
    main()
