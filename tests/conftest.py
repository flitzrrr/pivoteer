"""Pytest fixtures for pivoteer tests."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))


@pytest.fixture()
def template_path(tmp_path: Path) -> Path:
    from tests.generate_dummy_template import generate_template

    template = tmp_path / "dummy_template.xlsx"
    generate_template(template)
    return template
