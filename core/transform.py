\
from __future__ import annotations

from typing import List, Any


def apply_row_selection(rows: List[List[Any]], row_indices: List[int]) -> List[List[Any]]:
    """
    If row_indices is empty -> caller interprets as ALL rows.
    Otherwise select only those 0-based indices.
    """
    if not row_indices:
        return rows
    return [rows[i] for i in row_indices if 0 <= i < len(rows)]


def apply_column_selection(
    rows: List[List[Any]],
    col_indices: List[int],
) -> List[List[Any]]:
    """
    If col_indices is empty -> caller interprets as ALL columns.
    Otherwise select those 0-based indices.
    """
    if not rows:
        return []

    if not col_indices:
        return rows

    selected = []
    for row in rows:
        new_row = []
        for c in col_indices:
            if 0 <= c < len(row):
                new_row.append(row[c])
            else:
                new_row.append(None)
        selected.append(new_row)
    return selected


def shape_pack(rows: List[List[Any]]) -> List[List[Any]]:
    """
    Pack mode assumes rows/columns already selected.
    Simply returns rows (already tightly shaped).
    """
    return rows


def shape_keep(
    original_rows: List[List[Any]],
    selected_row_indices: List[int],
    selected_col_indices: List[int],
) -> List[List[Any]]:
    """
    Keep original spacing relative to selected rows/columns.
    Output is a rectangular bounding box from:
      min_row -> max_row
      min_col -> max_col
    """
    if not original_rows:
        return []

    if not selected_row_indices:
        selected_row_indices = list(range(len(original_rows)))

    if not selected_col_indices:
        # keep full width
        selected_col_indices = list(range(len(original_rows[0])))

    min_r = min(selected_row_indices)
    max_r = max(selected_row_indices)
    min_c = min(selected_col_indices)
    max_c = max(selected_col_indices)

    shaped = []

    for r in range(min_r, max_r + 1):
        row = []
        for c in range(min_c, max_c + 1):
            if r in selected_row_indices and c in selected_col_indices:
                if r < len(original_rows) and c < len(original_rows[r]):
                    row.append(original_rows[r][c])
                else:
                    row.append(None)
            else:
                row.append(None)
        shaped.append(row)

    return shaped
