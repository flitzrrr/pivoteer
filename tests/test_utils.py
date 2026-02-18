"""Unit tests for utility helpers."""

from __future__ import annotations

import pytest

from pivoteer.utils import (
    build_a1_cell,
    build_a1_range,
    column_index_to_letter,
    column_letter_to_index,
    parse_a1_cell,
    parse_a1_range,
)


class TestColumnIndexToLetter:
    def test_single_letters(self) -> None:
        assert column_index_to_letter(1) == "A"
        assert column_index_to_letter(26) == "Z"

    def test_double_letters(self) -> None:
        assert column_index_to_letter(27) == "AA"
        assert column_index_to_letter(28) == "AB"
        assert column_index_to_letter(52) == "AZ"
        assert column_index_to_letter(53) == "BA"
        assert column_index_to_letter(702) == "ZZ"

    def test_triple_letters(self) -> None:
        assert column_index_to_letter(703) == "AAA"

    def test_excel_max_column(self) -> None:
        assert column_index_to_letter(16384) == "XFD"

    def test_zero_raises(self) -> None:
        with pytest.raises(ValueError):
            column_index_to_letter(0)

    def test_negative_raises(self) -> None:
        with pytest.raises(ValueError):
            column_index_to_letter(-1)


class TestColumnLetterToIndex:
    def test_single_letters(self) -> None:
        assert column_letter_to_index("A") == 1
        assert column_letter_to_index("Z") == 26

    def test_double_letters(self) -> None:
        assert column_letter_to_index("AA") == 27
        assert column_letter_to_index("AZ") == 52
        assert column_letter_to_index("ZZ") == 702

    def test_triple_letters(self) -> None:
        assert column_letter_to_index("AAA") == 703
        assert column_letter_to_index("XFD") == 16384

    def test_lowercase_raises(self) -> None:
        with pytest.raises(ValueError):
            column_letter_to_index("a")

    def test_empty_raises(self) -> None:
        with pytest.raises(ValueError):
            column_letter_to_index("")

    def test_digits_raises(self) -> None:
        with pytest.raises(ValueError):
            column_letter_to_index("A1")


class TestRoundTrip:
    @pytest.mark.parametrize("index", [1, 13, 26, 27, 52, 256, 702, 703, 16384])
    def test_column_round_trip(self, index: int) -> None:
        letters = column_index_to_letter(index)
        assert column_letter_to_index(letters) == index


class TestParseA1Cell:
    def test_simple(self) -> None:
        assert parse_a1_cell("A1") == (1, 1)

    def test_double_letter(self) -> None:
        assert parse_a1_cell("AA100") == (100, 27)

    def test_invalid_raises(self) -> None:
        with pytest.raises(ValueError):
            parse_a1_cell("1A")

    def test_empty_raises(self) -> None:
        with pytest.raises(ValueError):
            parse_a1_cell("")


class TestBuildA1Cell:
    def test_simple(self) -> None:
        assert build_a1_cell(1, 1) == "A1"
        assert build_a1_cell(10, 3) == "C10"

    def test_zero_row_raises(self) -> None:
        with pytest.raises(ValueError):
            build_a1_cell(0, 1)

    def test_zero_col_raises(self) -> None:
        with pytest.raises(ValueError):
            build_a1_cell(1, 0)


class TestParseA1Range:
    def test_simple(self) -> None:
        result = parse_a1_range("A1:D10")
        assert result == ((1, 1), (10, 4))

    def test_invalid_raises(self) -> None:
        with pytest.raises(ValueError):
            parse_a1_range("A1")

    def test_single_cell_range(self) -> None:
        result = parse_a1_range("B2:B2")
        assert result == ((2, 2), (2, 2))


class TestBuildA1Range:
    def test_simple(self) -> None:
        assert build_a1_range(1, 1, 10, 4) == "A1:D10"

    def test_end_before_start_raises(self) -> None:
        with pytest.raises(ValueError):
            build_a1_range(10, 1, 5, 4)

    def test_round_trip(self) -> None:
        original = "B3:F20"
        parsed = parse_a1_range(original)
        rebuilt = build_a1_range(parsed[0][0], parsed[0][1], parsed[1][0], parsed[1][1])
        assert rebuilt == original
