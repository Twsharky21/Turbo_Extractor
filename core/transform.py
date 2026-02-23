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
    Pack mode: rows already have columns compacted by apply_column_selection.
    Simply return rows — already a dense rectangle with no gaps.
    """
    return rows


def shape_keep(
    original_rows: List[List[Any]],
    selected_row_indices: List[int],
    selected_col_indices: List[int],
) -> List[List[Any]]:
    """
    Keep Format: compress rows (no empty rows), but preserve column spacing.

    Output rows = only the selected rows (no gaps between rows).
    Output columns = full bounding box from min_col to max_col, with None
    in positions that were not selected (column gaps preserved).

    Example: source columns A,B,C,D,E with C,E selected (indices 2,4):
      - min_col=2, max_col=4 → output width = 3 (C, D, E)
      - output[r][0] = source col C, output[r][1] = None (D gap),
        output[r][2] = source col E
      - One output row per selected row; no empty rows.
    """
    if not original_rows:
        return []

    if not selected_row_indices:
        selected_row_indices = list(range(len(original_rows)))

    if not selected_col_indices:
        selected_col_indices = list(range(len(original_rows[0])))

    min_c = min(selected_col_indices)
    max_c = max(selected_col_indices)
    col_set = set(selected_col_indices)

    shaped = []

    for r in selected_row_indices:
        if r >= len(original_rows):
            continue
        row = []
        for c in range(min_c, max_c + 1):
            if c in col_set:
                src_row = original_rows[r]
                row.append(src_row[c] if c < len(src_row) else None)
            else:
                row.append(None)
        shaped.append(row)

    return shaped
