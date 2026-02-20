"""Tests for core.io — occupancy, normalization, used-range detection, CSV/XLSX loading."""
from core.io import is_occupied, normalize_table, compute_used_range


def test_is_occupied_none_and_empty_string():
    assert not is_occupied(None)
    assert not is_occupied("")


def test_is_occupied_truthy_values():
    assert is_occupied(" ")
    assert is_occupied(0)
    assert is_occupied("text")


def test_normalize_table_pads_ragged_rows():
    rows = [[1, 2], [3]]
    norm = normalize_table(rows)
    assert len(norm[1]) == 2
    assert norm[1][1] is None


def test_compute_used_range_basic():
    rows = [
        [None, None],
        [None, 5],
        [None, None],
    ]
    h, w = compute_used_range(rows)
    assert h == 2
    assert w == 2
