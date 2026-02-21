from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.worksheet.worksheet import Worksheet

from .errors import AppError, DEST_BLOCKED, BAD_SPEC
from .io import is_occupied
from .parsing import col_letters_to_index, col_index_to_letters


def is_cell_occupied(value: Any) -> bool:
    """
    Destination occupancy definition for planner collision + append scan.

    Per spec (Paste_Modes_and_Destination_AI_Spec.txt):
      OCCUPIED  — text, numbers (incl. 0), dates/booleans,
                  formula WITH a visible/cached result
      UNOCCUPIED — None, empty string "",
                   formula with NO visible result (no cached value)

    Formula handling: a bare formula string starting with '=' that has
    reached this function means openpyxl was opened WITHOUT data_only=True
    and the cell has no cached result. Treat as unoccupied.
    When opened with data_only=True (as load_xlsx does), the formula is
    replaced by its cached value, so this branch is rarely hit in practice.
    """
    if value is None:
        return False
    if isinstance(value, str):
        if value == "":
            return False
        if value.startswith("="):
            # Bare formula string with no cached result → unoccupied
            return False
    return True


@dataclass(frozen=True)
class WritePlan:
    start_row: int                 # 1-based
    start_col: int                 # 1-based
    width: int
    height: int
    landing_cols: Tuple[int, int]  # (col_start_1based, col_end_1based)
    landing_rows: Tuple[int, int]  # (row_start_1based, row_end_1based)


def _shape_dims(shaped: List[List[Any]]) -> Tuple[int, int]:
    if not shaped:
        return 0, 0
    h = len(shaped)
    w = max((len(r) for r in shaped), default=0)
    return h, w


def _max_used_row_in_cols(ws: Worksheet, col_start: int, col_end: int) -> int:
    """
    Scan ONLY the landing-zone columns to find the maximum occupied row.
    Uses is_cell_occupied (planner's definition, with formula handling).
    """
    max_row = 0
    upper = ws.max_row or 0
    if upper < 1:
        return 0

    for r in range(1, upper + 1):
        for c in range(col_start, col_end + 1):
            if is_cell_occupied(ws.cell(row=r, column=c).value):
                if r > max_row:
                    max_row = r
    return max_row


def build_plan(
    ws: Worksheet,
    shaped: List[List[Any]],
    start_col_letters: str,
    start_row_str: str,
) -> Optional[WritePlan]:
    """
    Build a write plan for a single shaped output table.

    start_row_str:
      - "" => append mode: place after max used row across landing columns.
      - numeric => explicit row; no append scan.

    Collision probe: full bounding rectangle. Any occupied cell → DEST_BLOCKED.
    If shaped height == 0 or width == 0: returns None (no write).
    """
    height, width = _shape_dims(shaped)
    if height == 0 or width == 0:
        return None

    start_col = col_letters_to_index(start_col_letters)
    col_end = start_col + width - 1

    append_mode = (start_row_str or "").strip() == ""

    if append_mode:
        max_used = _max_used_row_in_cols(ws, start_col, col_end)
        start_row = max_used + 1 if max_used > 0 else 1
    else:
        try:
            start_row = int(start_row_str)
        except ValueError:
            raise AppError(BAD_SPEC, f"Bad start row: {start_row_str!r}")
        if start_row <= 0:
            raise AppError(BAD_SPEC, f"Start row must be >= 1: {start_row_str!r}")

    # Collision probe: full bounding rectangle
    row_end = start_row + height - 1

    for r in range(start_row, row_end + 1):
        for c in range(start_col, col_end + 1):
            cell_val = ws.cell(row=r, column=c).value
            if is_cell_occupied(cell_val):
                raise AppError(
                    DEST_BLOCKED,
                    "Destination landing zone is blocked.",
                    details={
                        "append_mode": append_mode,
                        "target_start": f"{col_index_to_letters(start_col)}{start_row}",
                        "landing_cols": f"{col_index_to_letters(start_col)}:{col_index_to_letters(col_end)}",
                        "landing_rows": f"{start_row}:{row_end}",
                        "first_blocker": {
                            "row": r,
                            "col": c,
                            "col_letter": col_index_to_letters(c),
                            "value": cell_val,
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
    )
