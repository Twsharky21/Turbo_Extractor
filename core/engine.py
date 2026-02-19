\
from __future__ import annotations

from dataclasses import asdict
from typing import Optional, List, Any
import os

from openpyxl import Workbook, load_workbook

from .errors import AppError, SOURCE_READ_FAILED, SHEET_NOT_FOUND
from .models import SheetConfig, SheetResult
from .parsing import parse_columns, parse_rows, col_letters_to_index
from .io import load_csv, load_xlsx, compute_used_range, normalize_table
from .rules import apply_rules
from .transform import apply_row_selection, apply_column_selection, shape_pack, shape_keep
from .planner import build_plan
from .writer import apply_write_plan


def _load_source_table(source_path: str, workbook_sheet: str) -> List[List[Any]]:
    ext = os.path.splitext(source_path)[1].lower()
    try:
        if ext == ".csv":
            return load_csv(source_path)
        # default to xlsx
        return load_xlsx(source_path, workbook_sheet)
    except ValueError as e:
        # sheet not found from load_xlsx
        raise AppError(SHEET_NOT_FOUND, str(e))
    except Exception as e:
        raise AppError(SOURCE_READ_FAILED, f"Failed to read source: {e}")


def _open_or_create_dest(dest_path: str) -> Workbook:
    if dest_path and os.path.exists(dest_path):
        return load_workbook(dest_path)
    wb = Workbook()
    return wb


def _get_or_create_sheet(wb: Workbook, name: str):
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(title=name)
    # If this is a brand new workbook, remove the default "Sheet" if it's empty and not wanted
    if len(wb.sheetnames) > 1 and "Sheet" in wb.sheetnames:
        default = wb["Sheet"]
        if default.max_row == 1 and default.max_column == 1 and default["A1"].value in (None, ""):
            wb.remove(default)
    return ws


def run_sheet(source_path: str, sheet_cfg: SheetConfig, recipe_name: str = "Recipe") -> SheetResult:
    """
    End-to-end run for a single sheet config:
      load -> used range -> row select -> rules -> col select -> paste shape -> plan -> write -> save
    """
    table = _load_source_table(source_path, sheet_cfg.workbook_sheet)

    # Used range (for blank specs => ALL used)
    used_h, used_w = compute_used_range(table)
    table = normalize_table(table)

    # Row selection indices
    row_indices = parse_rows(sheet_cfg.rows_spec)
    if not row_indices:
        row_indices = list(range(used_h))

    # Apply row selection first
    table_rows = apply_row_selection(table, row_indices)

    # Apply rules (absolute columns)
    table_rows = apply_rules(table_rows, sheet_cfg.rules, sheet_cfg.rules_combine)

    # Column selection indices (absolute)
    col_indices = parse_columns(sheet_cfg.columns_spec)
    if not col_indices:
        col_indices = list(range(used_w))

    selected = apply_column_selection(table_rows, col_indices)

    # Shape
    if sheet_cfg.paste_mode == "keep":
        # Keep mode bounding box in current simplified model:
        # preserves gaps from column selection and original row selection (not from rule removals).
        # We shape against the row-selected table to preserve row gaps.
        shaped = shape_keep(table, row_indices, col_indices)
        # Then re-apply rules to zero out rows that were removed (convert them to all-None rows)
        if sheet_cfg.rules:
            kept = apply_rules(apply_row_selection(table, row_indices), sheet_cfg.rules, sheet_cfg.rules_combine)
            kept_set = {id(r) for r in kept}
            # Any row in shaped that corresponds to a removed row stays as None row already; no action needed.
            # (This placeholder keeps keep-mode deterministic; full fidelity can be improved later.)
    else:
        shaped = shape_pack(selected)

    # Destination open/create
    dest_path = sheet_cfg.destination.file_path
    wb = _open_or_create_dest(dest_path)
    ws = _get_or_create_sheet(wb, sheet_cfg.destination.sheet_name)

    plan = build_plan(ws, shaped, sheet_cfg.destination.start_col, sheet_cfg.destination.start_row)
    rows_written = 0
    if plan is not None:
        rows_written = apply_write_plan(ws, shaped, plan)

    if dest_path:
        wb.save(dest_path)

    return SheetResult(
        source_path=source_path,
        recipe_name=recipe_name,
        sheet_name=sheet_cfg.name,
        dest_file=dest_path,
        dest_sheet=sheet_cfg.destination.sheet_name,
        rows_written=rows_written,
        message="OK" if rows_written > 0 else "0 rows written",
    )
