"""Public API for pivoteer."""

from __future__ import annotations

import logging
import zipfile
from pathlib import Path
from typing import Dict

import pandas as pd

from pivoteer.template_engine import TemplateEngine


LOGGER = logging.getLogger(__name__)


class Pivoteer:
    """Public entry point for applying DataFrames to Excel templates."""

    def __init__(
        self, template_path: str | Path, *, enable_pivot_field_sync: bool = False
    ) -> None:
        """Initialize with optional pivot cache field synchronization."""
        self._template_engine = TemplateEngine(Path(template_path))
        self._enable_pivot_field_sync = enable_pivot_field_sync

    def apply_dataframe(self, table_name: str, df: pd.DataFrame) -> None:
        """Apply a DataFrame to the specified table."""
        self._template_engine.apply_dataframe(table_name, df)

    def save(self, output_path: str | Path) -> Path:
        """Write the modified template to a new file."""
        output_path = Path(output_path)
        if self._enable_pivot_field_sync:
            self._template_engine.sync_pivot_cache_fields()
        self._template_engine.ensure_pivot_refresh_on_load()
        modified_parts = self._template_engine.get_modified_parts()

        with zipfile.ZipFile(self._template_engine.template_path, "r") as src:
            with zipfile.ZipFile(
                output_path, "w", compression=zipfile.ZIP_DEFLATED
            ) as dest:
                for info in src.infolist():
                    filename = info.filename
                    if filename in modified_parts:
                        dest.writestr(info, modified_parts[filename])
                    else:
                        dest.writestr(info, src.read(filename))

        LOGGER.info("Saved output to %s", output_path)
        return output_path
