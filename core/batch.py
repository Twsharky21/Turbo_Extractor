"""
core/batch.py — Batch (run_all) execution coordinator.

Responsible for:
  - Iterating RunItems in tree order
  - Maintaining a shared in-memory workbook cache per destination file
  - Saving each destination after every successful write (crash safety)
  - Fail-fast on first error
  - Emitting optional progress callbacks

This module has NO knowledge of extraction logic — it delegates entirely
to core.runner.run_sheet.
"""
from __future__ import annotations

from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple

from openpyxl import Workbook

from .errors import AppError
from .models import SheetConfig, SheetResult, RunReport
from .runner import run_sheet


RunItem = Tuple[str, str, SheetConfig]
"""(source_path, recipe_name, sheet_cfg)"""


def run_all(
    items: Iterable[RunItem],
    on_progress: Optional[Callable[[str, Any], None]] = None,
) -> RunReport:
    """
    Execute all items in tree order using a shared in-memory workbook cache
    per destination file. This ensures successive runs to the same destination
    see each other's writes immediately without disk round-trips causing stale
    append-row calculations.

    Each workbook is saved to disk after every successful item for crash safety.
    Fail-fast on first error: remaining items are not executed.
    """
    results: List[SheetResult] = []
    ok = True
    wb_cache: Dict[str, Workbook] = {}

    def _emit(event: str, payload: Any) -> None:
        if on_progress is not None:
            try:
                on_progress(event, payload)
            except Exception:
                pass  # Progress callbacks must never break execution.

    for (source_path, recipe_name, sheet_cfg) in items:
        _emit("start", {
            "source_path": source_path,
            "recipe_name": recipe_name,
            "sheet_name": sheet_cfg.name,
        })

        try:
            result = run_sheet(source_path, sheet_cfg, recipe_name, _wb_cache=wb_cache)
        except AppError as e:
            result = SheetResult(
                source_path=source_path,
                recipe_name=recipe_name,
                sheet_name=sheet_cfg.name,
                dest_file=sheet_cfg.destination.file_path,
                dest_sheet=sheet_cfg.destination.sheet_name,
                rows_written=0,
                message=str(e),
                error_code=e.code,
                error_message=e.message,
                error_details=e.details,
            )
            results.append(result)
            _emit("error", result)
            ok = False
            # Fail-fast: save any workbooks written so far, then stop.
            for path, wb in wb_cache.items():
                if path:
                    try:
                        wb.save(path)
                    except Exception:
                        pass
            break

        # Successful write: persist to disk so file reflects current state.
        dest_path = sheet_cfg.destination.file_path
        if dest_path and dest_path in wb_cache:
            try:
                wb_cache[dest_path].save(dest_path)
            except Exception as save_err:
                result = SheetResult(
                    source_path=source_path,
                    recipe_name=recipe_name,
                    sheet_name=sheet_cfg.name,
                    dest_file=dest_path,
                    dest_sheet=sheet_cfg.destination.sheet_name,
                    rows_written=0,
                    message=f"Save failed: {save_err}",
                    error_code="SAVE_FAILED",
                    error_message=str(save_err),
                )
                results.append(result)
                _emit("error", result)
                ok = False
                break

        results.append(result)
        _emit("result", result)

    _emit("done", RunReport(ok=ok, results=results))
    return RunReport(ok=ok, results=results)
