\
from core.transform import (
    apply_row_selection,
    apply_column_selection,
    shape_pack,
    shape_keep,
)


def sample_rows():
    return [
        ["A1", "B1", "C1", "D1"],
        ["A2", "B2", "C2", "D2"],
        ["A3", "B3", "C3", "D3"],
        ["A4", "B4", "C4", "D4"],
    ]


def test_row_selection_basic():
    rows = sample_rows()
    result = apply_row_selection(rows, [0, 2])
    assert len(result) == 2
    assert result[1][0] == "A3"


def test_column_selection_basic():
    rows = sample_rows()
    result = apply_column_selection(rows, [0, 2])
    assert result[0] == ["A1", "C1"]


def test_pack_mode_identity():
    rows = sample_rows()
    packed = shape_pack(rows)
    assert packed == rows


def test_keep_mode_preserves_spacing():
    rows = sample_rows()
    selected_rows = [0, 2]
    selected_cols = [0, 2]

    shaped = shape_keep(rows, selected_rows, selected_cols)

    # bounding box: rows 0-2, cols 0-2
    assert len(shaped) == 3
    assert len(shaped[0]) == 3

    # row 1 should be blank (not selected)
    assert shaped[1] == [None, None, None]

    # selected positions preserved
    assert shaped[0][0] == "A1"
    assert shaped[0][2] == "C1"
    assert shaped[2][0] == "A3"
    assert shaped[2][2] == "C3"
