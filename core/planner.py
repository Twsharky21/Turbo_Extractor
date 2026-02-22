"""
core/planner.py — Write-plan builder.

Delegates ALL occupancy logic to core.landing.

Key behaviour change: scan and collision probe operate on TARGET COLUMNS ONLY
(columns in the shaped grid that carry actual data). Gap columns produced by
Keep Format bounding boxes are ignored for placement. This enables automatic
merge: two extractions whose bounding boxes overlap land side-by-side as long
as their data columns don't conflict.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, List, Optional, Tuple

from openpyxl.worksheet.worksheet import Worksheet

from .errors import AppError, DEST_BLOCKED, BAD_SPEC
from .landing import (
    is_dest_cell_occupied,
    find_target_col_offsets,
    read_zone,
    scan_target_cols,
    probe_target_cols,
)
from .parsing import col_letters_to_index, col_index_to_letters

# Backward-compat alias for existing tests that import is_cell_occupied from planner
is_cell_occupied = is_dest_cell_occupied


@dataclass(frozen=True)
class WritePlan:
    start_row: int                  # 1-based
    start_col: int                  # 1-based
    width: int
    height: int
    landing_cols: Tuple[int, int]   # full bounding box (col_start, col_end)
    landing_rows: Tuple[int, int]   # (row_start, row_end)
    target_cols: Tuple[int, ...]    # absolute 1-based cols that receive data


def _shape_dims(shaped: List[List[Any]]) -> Tuple[int, int]:
    if not shaped:
        return 0, 0
    h = len(shaped)
    w = max((len(r) for r in shaped), default=0)
    return h, w


def build_plan(
    ws: Worksheet,
    shaped: List[List[Any]],
    start_col_letters: str,
    start_row_str: str,
) -> Optional[WritePlan]:
    """
    Build a validated write plan.

    Scan and probe use TARGET COLUMNS ONLY — the subset of bounding-box columns
    that actually carry non-None data in the shaped grid. Gap columns (all-None,
    produced by Keep Format) are never scanned or probed.

    start_row_str:
      ""       → append / merge mode: scan target cols, place after max used row.
      numeric  → explicit mode: place at exactly that row, probe target cols.

    Raises AppError(DEST_BLOCKED) if any target column cell in the landing
    rectangle is occupied. Returns None if shaped is empty.
    """
    height, width = _shape_dims(shaped)
    if height == 0 or width == 0:
        return None

    start_col = col_letters_to_index(start_col_letters)
    col_end = start_col + width - 1
    append_mode = (start_row_str or "").strip() == ""

    # Determine which columns within the shaped grid carry actual data.
    offsets = find_target_col_offsets(shaped)
    if not offsets:
        # Shaped grid is all-None — nothing to write.
        return None
    target_abs_cols = [start_col + o for o in offsets]

    t_col_min = min(target_abs_cols)
    t_col_max = max(target_abs_cols)

    if append_mode:
        # Scan only target columns to find the highest occupied row.
        scan_map = read_zone(ws, t_col_min, t_col_max, extra_rows=0)
        max_used = scan_target_cols(scan_map, target_abs_cols)
        start_row = max_used + 1 if max_used > 0 else 1
    else:
        try:
            start_row = int(start_row_str)
        except ValueError:
            raise AppError(BAD_SPEC, f"Bad start row: {start_row_str!r}")
        if start_row <= 0:
            raise AppError(BAD_SPEC, f"Start row must be >= 1: {start_row_str!r}")

    row_end = start_row + height - 1

    # Probe target columns in the landing rectangle.
    extra = max(0, row_end - (ws.max_row or 0))
    probe_map = read_zone(ws, t_col_min, t_col_max, extra_rows=extra)

    blocker = probe_target_cols(probe_map, start_row, row_end, target_abs_cols)
    if blocker is not None:
        b_row, b_col, b_val = blocker
        raise AppError(
            DEST_BLOCKED,
            "Destination landing zone is blocked.",
            details={
                "append_mode": append_mode,
                "target_start": f"{col_index_to_letters(start_col)}{start_row}",
                "landing_cols": f"{col_index_to_letters(start_col)}:{col_index_to_letters(col_end)}",
                "landing_rows": f"{start_row}:{row_end}",
                "target_data_cols": [col_index_to_letters(c) for c in target_abs_cols],
                "first_blocker": {
                    "row": b_row,
                    "col": b_col,
                    "col_letter": col_index_to_letters(b_col),
                    "value": b_val,
                },
            },
        )

    return WritePlan(
        start_row=start_row,
        start_col=start_col,
        width=width,
        height=height,
        landing_cols=(start_col, col_end),
        landing_rows=(start_row, row_end),
        target_cols=tuple(target_abs_cols),
    )
