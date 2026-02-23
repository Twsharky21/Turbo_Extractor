"""
test_engine.py — Smoke tests for the core.engine re-export shim.

Verifies that `from core.engine import run_sheet, run_all` continues to work
after the logic was split into core.runner and core.batch.

Full coverage lives in:
  tests/test_runner.py  — single-sheet extraction
  tests/test_batch.py   — batch execution
"""
from __future__ import annotations

import os
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from core.engine import run_sheet, run_all, RunItem
from core.models import Destination, SheetConfig


def _make_xlsx(path, data=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    if data:
        for r, row in enumerate(data, 1):
            for c, v in enumerate(row, 1):
                ws.cell(row=r, column=c, value=v)
    wb.save(path)
    return path


def _cfg(dest):
    return SheetConfig(
        name="Sheet1", workbook_sheet="Sheet1",
        destination=Destination(file_path=dest, sheet_name="Out",
                                start_col="A", start_row=""),
    )


def test_engine_run_sheet_shim():
    """run_sheet imported from core.engine delegates to core.runner correctly."""
    with TemporaryDirectory() as td:
        src  = _make_xlsx(os.path.join(td, "src.xlsx"), data=[["hello"]])
        dest = os.path.join(td, "dest.xlsx")
        result = run_sheet(src, _cfg(dest))
        assert result.rows_written == 1


def test_engine_run_all_shim():
    """run_all imported from core.engine delegates to core.batch correctly."""
    with TemporaryDirectory() as td:
        src  = _make_xlsx(os.path.join(td, "src.xlsx"), data=[["world"]])
        dest = os.path.join(td, "dest.xlsx")
        report = run_all([(src, "R1", _cfg(dest))])
        assert report.ok
        assert report.results[0].rows_written == 1


def test_engine_run_item_type_alias_importable():
    """RunItem type alias is importable from core.engine."""
    assert RunItem is not None
