\
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

    Current rule:
      - None / "" -> unoccupied
      - Everything else -> occupied
      - Formula handling: if the stored value is a formula string (starts with '='),
        we conservatively treat it as unoccupied unless it has a visible cached value.
        (We don't attempt to evaluate formulas here; tests avoid formula cases for now.)
    """
    if value is None:
        return False
    if isinstance(value, str):
        if value == "":
            return False
        if value.startswith("="):
            # Treat formula-only as unoccupied (matches spec for "no visible value").
            # If later we need cached result detection, we can enhance this using
            # data_only snapshots.
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
    Scan ONLY the landing-zone columns to find the maximum used row.
    """
    max_row = 0
    # Iterate through worksheet's current max_row as an upper bound.
    # If sheet is sparse, this is still fine for initial implementation.
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
      - "" => append mode (landing-zone aware): append after max used row across landing columns
      - numeric => explicit row (no append scan)
    Collision probe:
      - probes full bounding rectangle of width/height (even in keep mode)
      - if any occupied cell found => raises AppError(DEST_BLOCKED, ...)
    If shaped height == 0 or width == 0:
      - returns None (no write)
    """
    height, width = _shape_dims(shaped)
    if height == 0 or width == 0:
        return None

    start_col = col_letters_to_index(start_col_letters)

    # Determine target start row
    if (start_row_str or "").strip() == "":
        # Full landing zone awareness: scan all columns the shaped output will occupy
        col_end = start_col + width - 1
        max_used = _max_used_row_in_cols(ws, start_col, col_end)
        start_row = max_used + 1 if max_used > 0 else 1
    else:
        try:
            start_row = int(start_row_str)
        except ValueError:
            raise AppError(BAD_SPEC, f"Bad start row: {start_row_str!r}")
        if start_row <= 0:
            raise AppError(BAD_SPEC, f"Start row must be >= 1: {start_row_str!r}")
        col_end = start_col + width - 1

    # Collision probe full bounding rectangle
    blockers: List[Dict[str, Any]] = []
    row_end = start_row + height - 1
    col_end = start_col + width - 1

    for r in range(start_row, row_end + 1):
        for c in range(start_col, col_end + 1):
            if is_cell_occupied(ws.cell(row=r, column=c).value):
                blockers.append({
                    "row": r,
                    "col": c,
                    "col_letter": col_index_to_letters(c),
                    "value": ws.cell(row=r, column=c).value,
                })
                # Fail-fast: first blocker is enough for plan
                raise AppError(
                    DEST_BLOCKED,
                    "Destination landing zone is blocked.",
                    details={
                        "append_mode": (start_row_str or "").strip() == "",
                        "target_start": f"{col_index_to_letters(start_col)}{start_row}",
                        "landing_cols": f"{col_index_to_letters(start_col)}:{col_index_to_letters(col_end)}",
                        "landing_rows": f"{start_row}:{row_end}",
                        "first_blocker": blockers[0],
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
