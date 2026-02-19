\
from __future__ import annotations

from typing import List, Any, Iterable, Tuple
import os

from openpyxl import Workbook, load_workbook

from .errors import AppError, SOURCE_READ_FAILED, SHEET_NOT_FOUND
from .models import SheetConfig, SheetResult, RunReport
from .parsing import parse_columns, parse_rows
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
        return load_xlsx(source_path, workbook_sheet)
    except ValueError as e:
        raise AppError(SHEET_NOT_FOUND, str(e))
    except AppError:
        raise
    except Exception as e:
        raise AppError(SOURCE_READ_FAILED, f"Failed to read source: {e}")


def _open_or_create_dest(dest_path: str) -> Workbook:
    if dest_path and os.path.exists(dest_path):
        return load_workbook(dest_path)
    return Workbook()


def _get_or_create_sheet(wb: Workbook, name: str):
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(title=name)
    if len(wb.sheetnames) > 1 and "Sheet" in wb.sheetnames:
        default = wb["Sheet"]
        if default.max_row == 1 and default.max_column == 1 and default["A1"].value in (None, ""):
            wb.remove(default)
    return ws


def run_sheet(source_path: str, sheet_cfg: SheetConfig, recipe_name: str = "Recipe") -> SheetResult:
    table = _load_source_table(source_path, sheet_cfg.workbook_sheet)

    used_h, used_w = compute_used_range(table)
    table = normalize_table(table)

    row_indices = parse_rows(sheet_cfg.rows_spec)
    if not row_indices:
        row_indices = list(range(used_h))

    table_rows = apply_row_selection(table, row_indices)

    table_rows = apply_rules(table_rows, sheet_cfg.rules, sheet_cfg.rules_combine)

    col_indices = parse_columns(sheet_cfg.columns_spec)
    if not col_indices:
        col_indices = list(range(used_w))

    selected = apply_column_selection(table_rows, col_indices)

    if sheet_cfg.paste_mode == "keep":
        shaped = shape_keep(table, row_indices, col_indices)
    else:
        shaped = shape_pack(selected)

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


RunItem = Tuple[str, str, SheetConfig]
"""
(source_path, recipe_name, sheet_cfg)
"""


def run_all(items: Iterable[RunItem]) -> RunReport:
    results: List[SheetResult] = []
    ok = True

    for (source_path, recipe_name, sheet_cfg) in items:
        try:
            res = run_sheet(source_path, sheet_cfg, recipe_name=recipe_name)
            results.append(res)
        except AppError as e:
            ok = False
            results.append(
                SheetResult(
                    source_path=source_path,
                    recipe_name=recipe_name,
                    sheet_name=sheet_cfg.name,
                    dest_file=sheet_cfg.destination.file_path,
                    dest_sheet=sheet_cfg.destination.sheet_name,
                    rows_written=0,
                    message="ERROR",
                    error_code=e.code,
                    error_message=e.message,
                    error_details=e.details,
                )
            )
            break
        except Exception as e:
            ok = False
            results.append(
                SheetResult(
                    source_path=source_path,
                    recipe_name=recipe_name,
                    sheet_name=sheet_cfg.name,
                    dest_file=sheet_cfg.destination.file_path,
                    dest_sheet=sheet_cfg.destination.sheet_name,
                    rows_written=0,
                    message="ERROR",
                    error_code="UNEXPECTED",
                    error_message=str(e),
                    error_details=None,
                )
            )
            break

    return RunReport(ok=ok, results=results)
