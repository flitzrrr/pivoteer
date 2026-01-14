"""Generate a template, inject data, and save an updated report."""

from __future__ import annotations

import logging
from datetime import date, timedelta
from pathlib import Path
from typing import List

import pandas as pd

from pivoteer.core import Pivoteer
from tests.generate_dummy_template import generate_template


LOGGER = logging.getLogger(__name__)


def _build_dataframe(rows: int) -> pd.DataFrame:
    categories = ["Hardware", "Software", "Services"]
    regions = ["North", "South", "East", "West"]
    base = date(2024, 1, 1)

    data: List[dict[str, object]] = []
    for i in range(rows):
        data.append(
            {
                "Category": categories[i % len(categories)],
                "Region": regions[i % len(regions)],
                "Amount": 100.0 + (i * 10.0),
                "Date": base + timedelta(days=i),
            }
        )

    return pd.DataFrame(data)


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    template_path = Path("dummy_template.xlsx")
    output_path = Path("report_output.xlsx")

    generate_template(template_path)

    df = _build_dataframe(500)
    pivoteer = Pivoteer(template_path)
    pivoteer.apply_dataframe("DataSource", df)
    pivoteer.save(output_path)

    LOGGER.info("Report saved at %s", output_path)


if __name__ == "__main__":
    main()
