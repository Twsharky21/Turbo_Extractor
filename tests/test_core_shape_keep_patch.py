"""
test_core_shape_keep_patch.py — Replacement tests for shape_keep behavior.

These replace the two outdated shape_keep tests in test_core.py:
  - test_shape_keep_preserves_gaps_between_selected_rows_and_cols
  - test_shape_keep_bounding_box_height_and_width
"""
from __future__ import annotations

from core.transform import shape_keep


def test_shape_keep_preserves_col_gaps_compresses_rows():
    """
    Keep Format: rows are compressed (no empty rows between selected rows),
    but column gaps within the bounding box are preserved as None.

    Source rows 0 and 2 selected, cols 0 and 2 selected:
      - Output has 2 rows (only selected rows, no gap row for row 1).
      - Output has 3 columns: col0 (data), col1 (None gap), col2 (data).
    """
    original = [
        ["A1", "B1", "C1"],
        ["A2", "B2", "C2"],
        ["A3", "B3", "C3"],
    ]
    result = shape_keep(original, [0, 2], [0, 2])
    assert len(result) == 2          # 2 selected rows, no empty row gap
    assert len(result[0]) == 3       # bounding box cols 0-2
    assert result[0][0] == "A1"
    assert result[0][1] is None      # col gap preserved
    assert result[0][2] == "C1"
    assert result[1][0] == "A3"      # row 2 immediately follows row 0
    assert result[1][1] is None      # col gap preserved
    assert result[1][2] == "C3"


def test_shape_keep_empty_col_indices_uses_all_cols():
    original = [["a", "b", "c"], ["d", "e", "f"]]
    result = shape_keep(original, [0, 1], [])
    assert result == original


def test_shape_keep_empty_row_indices_uses_all_rows():
    original = [["a", "b"], ["c", "d"]]
    result = shape_keep(original, [], [0, 1])
    assert result == original


def test_shape_keep_row_count_equals_selected_rows():
    """
    Keep Format output height = number of selected rows (not bounding box height).
    Column width = bounding box width (max_col - min_col + 1).
    """
    original = [
        ["a", "b", "c", "d", "e"],
        ["f", "g", "h", "i", "j"],
        ["k", "l", "m", "n", "o"],
        ["p", "q", "r", "s", "t"],
        ["u", "v", "w", "x", "y"],
    ]
    result = shape_keep(original, [0, 2, 4], [0, 2, 4])
    assert len(result) == 3      # 3 selected rows, not 5
    assert len(result[0]) == 5   # cols 0-4 bounding box width preserved
    assert result[0][0] == "a"
    assert result[0][1] is None  # col gap
    assert result[0][2] == "c"
    assert result[0][3] is None  # col gap
    assert result[0][4] == "e"
    assert result[1][0] == "k"   # row index 2, immediately after row index 0
    assert result[2][0] == "u"   # row index 4
