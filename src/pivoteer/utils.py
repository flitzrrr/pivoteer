"""Utility helpers for Excel XML manipulation."""

from __future__ import annotations

import re
from typing import Tuple


_A1_RE = re.compile(r"^([A-Z]+)([0-9]+)$")
_A1_RANGE_RE = re.compile(r"^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$")


def column_index_to_letter(index: int) -> str:
    """Convert a 1-based column index to Excel column letters."""
    if index < 1:
        raise ValueError("Column index must be >= 1.")

    letters = []
    current = index
    while current:
        current, remainder = divmod(current - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def column_letter_to_index(letters: str) -> int:
    """Convert Excel column letters to a 1-based column index."""
    if not letters or not letters.isalpha() or not letters.isupper():
        raise ValueError(f"Invalid column letters: {letters!r}")

    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - 64)
    return index


def parse_a1_cell(cell: str) -> Tuple[int, int]:
    """Parse an A1 cell reference into (row, col) 1-based coordinates."""
    match = _A1_RE.match(cell)
    if not match:
        raise ValueError(f"Invalid A1 cell reference: {cell!r}")

    col_letters, row_str = match.groups()
    return int(row_str), column_letter_to_index(col_letters)


def parse_a1_range(a1_range: str) -> Tuple[Tuple[int, int], Tuple[int, int]]:
    """Parse an A1 range like A1:D10 into row/col tuples."""
    match = _A1_RANGE_RE.match(a1_range)
    if not match:
        raise ValueError(f"Invalid A1 range: {a1_range!r}")

    start_cell, end_cell = match.groups()
    return parse_a1_cell(start_cell), parse_a1_cell(end_cell)


def build_a1_cell(row: int, col: int) -> str:
    """Build an A1 cell reference from 1-based row/col."""
    if row < 1 or col < 1:
        raise ValueError("Row and column must be >= 1.")
    return f"{column_index_to_letter(col)}{row}"


def build_a1_range(
    start_row: int, start_col: int, end_row: int, end_col: int
) -> str:
    """Build an A1 range from 1-based coordinates."""
    if end_row < start_row or end_col < start_col:
        raise ValueError("Range end must be >= range start.")
    return f"{build_a1_cell(start_row, start_col)}:{build_a1_cell(end_row, end_col)}"
