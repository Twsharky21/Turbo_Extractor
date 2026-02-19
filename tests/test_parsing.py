\
import pytest

from core.errors import AppError, BAD_SPEC
from core.parsing import (
    col_letters_to_index,
    col_index_to_letters,
    parse_columns,
    parse_rows,
)


def test_col_letters_to_index_basic():
    assert col_letters_to_index("A") == 1
    assert col_letters_to_index("Z") == 26
    assert col_letters_to_index("AA") == 27
    assert col_letters_to_index("AZ") == 52
    assert col_letters_to_index("BA") == 53


def test_col_index_to_letters_basic():
    assert col_index_to_letters(1) == "A"
    assert col_index_to_letters(26) == "Z"
    assert col_index_to_letters(27) == "AA"
    assert col_index_to_letters(52) == "AZ"
    assert col_index_to_letters(53) == "BA"


def test_col_roundtrip_property_small():
    for n in range(1, 200):
        assert col_letters_to_index(col_index_to_letters(n)) == n


def test_parse_columns_blank_means_all_by_caller():
    assert parse_columns("") == []
    assert parse_columns("   ") == []


def test_parse_columns_singletons_and_ranges_unique_sorted():
    assert parse_columns("A,C") == [0, 2]
    assert parse_columns("C,A") == [0, 2]
    assert parse_columns("A-C") == [0, 1, 2]
    assert parse_columns("A-C,C,A") == [0, 1, 2]


def test_parse_columns_whitespace_and_trailing_commas():
    assert parse_columns(" A, C, ") == [0, 2]
    cols = parse_columns("A, C, AC-AE,")
    assert cols[0] == 0
    assert 2 in cols


def test_parse_columns_reverse_range_normalizes():
    assert parse_columns("D-B") == [1, 2, 3]


def test_parse_columns_rejects_bad_tokens():
    with pytest.raises(AppError) as ei:
        parse_columns("A,??,C")
    assert ei.value.code == BAD_SPEC


def test_parse_rows_blank_means_all_by_caller():
    assert parse_rows("") == []
    assert parse_rows("   ") == []


def test_parse_rows_singletons_and_ranges_unique_sorted():
    assert parse_rows("1,3") == [0, 2]
    assert parse_rows("3,1") == [0, 2]
    assert parse_rows("1-3") == [0, 1, 2]
    assert parse_rows("1-3,3-5,10,12-13") == [0, 1, 2, 3, 4, 9, 11, 12]


def test_parse_rows_whitespace_and_trailing_commas():
    assert parse_rows(" 1 , 2-3, ") == [0, 1, 2]


def test_parse_rows_reverse_range_normalizes():
    assert parse_rows("6-4") == [3, 4, 5]


def test_parse_rows_rejects_bad_tokens():
    with pytest.raises(AppError) as ei:
        parse_rows("1,2,nope")
    assert ei.value.code == BAD_SPEC


def test_parse_rows_rejects_nonpositive():
    with pytest.raises(AppError):
        parse_rows("0")
    with pytest.raises(AppError):
        parse_rows("-1")
    with pytest.raises(AppError):
        parse_rows("1-0")
