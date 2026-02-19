\
from __future__ import annotations

from typing import List, Any
from openpyxl.worksheet.worksheet import Worksheet

from .planner import WritePlan


def apply_write_plan(
    ws: Worksheet,
    shaped: List[List[Any]],
    plan: WritePlan,
) -> int:
    """
    Applies a previously validated WritePlan to the worksheet.

    Assumes collision has already been checked in planner.
    Returns number of rows written.
    """
    if not shaped:
        return 0

    for r_offset, row in enumerate(shaped):
        for c_offset, value in enumerate(row):
            ws.cell(
                row=plan.start_row + r_offset,
                column=plan.start_col + c_offset,
                value=value,
            )

    return len(shaped)
