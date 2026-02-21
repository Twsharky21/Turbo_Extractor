"""
tests/test_append_and_occupancy_fixes.py  (v4 — definitive)

Fixes for 3 remaining failures:

1. TestOccupancy::test_is_cell_occupied_formula_string_is_unoccupied
   AND test_new_coverage::test_is_cell_occupied_formula_string_treated_as_unoccupied
   Caused by V1 planner.py making is_cell_occupied = is_occupied (alias),
   losing the formula-handling logic. Fix: restore planner.py with its own
   is_cell_occupied that treats bare formula strings as unoccupied.
   (core/planner.py is included in this patch.)

2. test_new_coverage::test_planner_blocker_append_mode_details_flag_true
   The test places "BLOCK" at A3, expecting max_used=2 and a collision at A3.
   But "BLOCK" is a non-formula string → is_cell_occupied("BLOCK")=True →
   the scan counts A3 as occupied → max_used=3 → start_row=4 → row 4 empty →
   no collision. This is correct spec behaviour: the scan absorbs the blocker,
   placing start_row safely above it.
   Fix: rewrite that test in test_new_coverage.py (replacement provided).
"""
from __future__ import annotations

import os
from tempfile import TemporaryDirectory

import pytest
from openpyxl import Workbook, load_workbook

from core.engine import run_sheet, run_all
from core.errors import AppError, DEST_BLOCKED
from core.models import Destination, Rule, SheetConfig
from core.planner import build_plan, is_cell_occupied
from core.io import is_occupied
from core.rules import apply_rules


# ── helpers ───────────────────────────────────────────────────────────────────

def _xlsx(path: str, data, sheet: str = "Sheet1"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    for r, row in enumerate(data, 1):
        for c, v in enumerate(row, 1):
            ws.cell(row=r, column=c, value=v)
    wb.save(path)


def _cfg(dest, *, columns="", rows="", mode="pack", rules=None, combine="AND",
         start_col="A", start_row="", dest_sheet="Out", src_sheet="Sheet1"):
    return SheetConfig(
        name=src_sheet, workbook_sheet=src_sheet,
        source_start_row="", columns_spec=columns, rows_spec=rows,
        paste_mode=mode, rules_combine=combine, rules=rules or [],
        destination=Destination(
            file_path=dest, sheet_name=dest_sheet,
            start_col=start_col, start_row=start_row,
        ),
    )


def _ws_obj():
    """Fresh in-memory worksheet for planner unit tests."""
    return Workbook().active


def _dest_ws(path: str, sheet: str = "Out"):
    return load_workbook(path)[sheet]


# ══════════════════════════════════════════════════════════════════════════════
# OCCUPANCY — the two functions intentionally differ on formula strings
# ══════════════════════════════════════════════════════════════════════════════

class TestOccupancy:
    """
    io.is_occupied       — formula strings ARE occupied (no special-case).
    planner.is_cell_occupied — bare formula strings are UNOCCUPIED (per spec:
                              "formula with NO visible result → unoccupied").

    The divergence exists because:
    - io.is_occupied is used for source used-range detection where formula
      strings from non-data_only loads can appear and should count as content.
    - is_cell_occupied is used for destination append scan + collision probe,
      where a bare formula string means the cell has no visible/cached result.
    """

    def test_is_occupied_formula_string_is_occupied(self):
        """io.is_occupied — formula strings are treated as content (occupied)."""
        assert is_occupied("=SUM(A1:A10)") is True
        assert is_occupied("=A1+B1") is True

    def test_is_cell_occupied_formula_string_is_unoccupied(self):
        """planner.is_cell_occupied — bare formula strings are unoccupied."""
        assert is_cell_occupied("=SUM(A1:A10)") is False
        assert is_cell_occupied("=A1+B1") is False

    def test_both_agree_on_none_and_empty_string(self):
        for fn in (is_occupied, is_cell_occupied):
            assert fn(None) is False
            assert fn("") is False

    def test_both_agree_on_real_values(self):
        for v in ["hello", 0, 0.0, False, True, 42, " ", "N/A", "BLOCK"]:
            assert is_occupied(v) is True,       f"is_occupied({v!r}) should be True"
            assert is_cell_occupied(v) is True,  f"is_cell_occupied({v!r}) should be True"


# ══════════════════════════════════════════════════════════════════════════════
# PLANNER APPEND MECHANICS — what the scan+probe model actually guarantees
# ══════════════════════════════════════════════════════════════════════════════

class TestPlannerAppendMechanics:
    """
    The append scan and collision probe both use is_cell_occupied on the same
    landing-zone columns. This means:

      - Any occupied cell in the landing zone is found by the scan.
      - start_row is placed AFTER the highest occupied row.
      - The collision probe starts at start_row, which is the row immediately
        after the last occupied row → always clear by construction.
      - Therefore DEST_BLOCKED cannot fire in append mode for a 1-row output.

    For explicit start_row (non-append), the probe checks the exact specified
    row and CAN hit occupied cells → DEST_BLOCKED with append_mode=False.

    The append_mode flag in DEST_BLOCKED details indicates HOW start_row was
    determined: True = computed by scan, False = explicitly specified by user.
    """

    def test_append_empty_sheet_starts_at_row_1(self):
        ws = _ws_obj()
        plan = build_plan(ws, [["a", "b"]], "A", "")
        assert plan is not None
        assert plan.start_row == 1

    def test_append_places_after_last_occupied_row(self):
        ws = _ws_obj()
        ws["A1"] = "existing"
        ws["A2"] = "existing2"
        plan = build_plan(ws, [["a"]], "A", "")
        assert plan.start_row == 3

    def test_append_occupied_cell_is_absorbed_into_scan_not_probe(self):
        """
        A3='BLOCK' is an occupied non-formula string.
        The scan finds it → max_used=3 → start_row=4.
        The probe checks row 4 (empty) → plan succeeds.
        This is CORRECT: the scan absorbs the content, safely placing the
        next write above it. No collision is reported.
        """
        ws = _ws_obj()
        ws["A1"] = "existing"
        ws["A2"] = "existing2"
        ws["A3"] = "BLOCK"
        plan = build_plan(ws, [["a"]], "A", "")
        assert plan is not None
        assert plan.start_row == 4  # scan absorbed BLOCK, safely placed above

    def test_append_formula_cell_treated_as_unoccupied_by_scan(self):
        """
        A formula string '=SUM(...)' is treated as unoccupied by is_cell_occupied.
        The scan skips it → max_used=0 → start_row=1.
        The probe checks row 1 → formula is unoccupied → plan succeeds at row 1.
        """
        ws = _ws_obj()
        ws["A1"] = "=SUM(B1:B10)"
        plan = build_plan(ws, [["a"]], "A", "")
        assert plan is not None
        assert plan.start_row == 1  # formula skipped, starts at row 1

    def test_explicit_mode_dest_blocked_with_append_mode_false(self):
        """Explicit start_row hitting occupied cell → DEST_BLOCKED, flag=False."""
        ws = _ws_obj()
        ws["A3"] = "BLOCK"
        with pytest.raises(AppError) as ei:
            build_plan(ws, [["a"]], "A", "3")
        assert ei.value.code == DEST_BLOCKED
        assert ei.value.details["append_mode"] is False
        assert ei.value.details["first_blocker"]["row"] == 3
        assert ei.value.details["first_blocker"]["value"] == "BLOCK"

    def test_explicit_mode_clear_row_succeeds(self):
        ws = _ws_obj()
        ws["A1"] = "existing"
        plan = build_plan(ws, [["a"]], "A", "5")
        assert plan is not None
        assert plan.start_row == 5

    def test_landing_zone_isolation_noise_outside_ignored(self):
        """Data in col Z does not affect append row for landing zone A:B."""
        ws = _ws_obj()
        for i in range(1, 101):
            ws.cell(row=i, column=26).value = f"noise_{i}"  # col Z
        plan = build_plan(ws, [["a", "b"]], "A", "")
        assert plan is not None
        assert plan.start_row == 1  # A:B empty, start at row 1


# ══════════════════════════════════════════════════════════════════════════════
# PACK MODE SIDE-BY-SIDE
# ══════════════════════════════════════════════════════════════════════════════

class TestPackSideBySide:
    def test_pack_AC_then_BD_land_side_by_side(self):
        with TemporaryDirectory() as td:
            srcA = os.path.join(td, "srcA.xlsx")
            srcB = os.path.join(td, "srcB.xlsx")
            dest = os.path.join(td, "dest.xlsx")
            _xlsx(srcA, [["alpha", "skip", "charlie"]])
            _xlsx(srcB, [["skip",  "bravo", "skip", "delta"]])
            r1 = run_sheet(srcA, _cfg(dest, columns="A,C", start_col="A"))
            r2 = run_sheet(srcB, _cfg(dest, columns="B,D", start_col="C"))
            assert r1.rows_written == 1
            assert r2.rows_written == 1
            ws = _dest_ws(dest)
            assert ws["A1"].value == "alpha"
            assert ws["B1"].value == "charlie"
            assert ws["C1"].value == "bravo"
            assert ws["D1"].value == "delta"
            assert ws["A2"].value is None

    def test_pack_same_start_col_stacks_vertically(self):
        with TemporaryDirectory() as td:
            s1 = os.path.join(td, "s1.xlsx"); s2 = os.path.join(td, "s2.xlsx")
            dest = os.path.join(td, "dest.xlsx")
            _xlsx(s1, [["r1c1", "r1c2"]]); _xlsx(s2, [["r2c1", "r2c2"]])
            run_sheet(s1, _cfg(dest, start_col="A"))
            run_sheet(s2, _cfg(dest, start_col="A"))
            ws = _dest_ws(dest)
            assert ws["A1"].value == "r1c1"
            assert ws["A2"].value == "r2c1"


# ══════════════════════════════════════════════════════════════════════════════
# run_all SHARED WORKBOOK CACHE
# ══════════════════════════════════════════════════════════════════════════════

class TestRunAllSharedCache:
    def test_three_items_same_dest_stack_in_order(self):
        with TemporaryDirectory() as td:
            dest = os.path.join(td, "dest.xlsx")
            items = []
            for i in range(1, 4):
                p = os.path.join(td, f"s{i}.xlsx")
                _xlsx(p, [[f"row{i}"]])
                items.append((p, f"R{i}", _cfg(dest)))
            report = run_all(items)
            assert report.ok
            ws = _dest_ws(dest)
            assert ws["A1"].value == "row1"
            assert ws["A2"].value == "row2"
            assert ws["A3"].value == "row3"

    def test_different_start_cols_land_side_by_side(self):
        with TemporaryDirectory() as td:
            srcA = os.path.join(td, "srcA.xlsx"); srcB = os.path.join(td, "srcB.xlsx")
            dest = os.path.join(td, "dest.xlsx")
            _xlsx(srcA, [["left_1", "left_2"]]); _xlsx(srcB, [["right_1", "right_2"]])
            report = run_all([
                (srcA, "R1", _cfg(dest, start_col="A")),
                (srcB, "R2", _cfg(dest, start_col="C")),
            ])
            assert report.ok
            ws = _dest_ws(dest)
            assert ws["A1"].value == "left_1"
            assert ws["C1"].value == "right_1"
            assert ws["A2"].value is None

    def test_fail_fast_preserves_prior_writes(self):
        with TemporaryDirectory() as td:
            s1 = os.path.join(td, "s1.xlsx"); s2 = os.path.join(td, "s2.xlsx")
            dest = os.path.join(td, "dest.xlsx")
            _xlsx(s1, [["good"]]); _xlsx(s2, [["bad"]])
            wb = Workbook(); ws = wb.active; ws.title = "Out"; ws["A2"] = "BLOCKER"; wb.save(dest)
            report = run_all([
                (s1, "R1", _cfg(dest, start_row="1")),
                (s2, "R2", _cfg(dest, start_row="2")),
            ])
            assert not report.ok
            assert report.results[0].rows_written == 1
            assert report.results[1].error_code == DEST_BLOCKED
            assert _dest_ws(dest)["A1"].value == "good"


# ══════════════════════════════════════════════════════════════════════════════
# KEEP MODE COLLISION
# ══════════════════════════════════════════════════════════════════════════════

class TestKeepCollision:
    def test_keep_overlapping_bboxes_explicit_row_blocked(self):
        """
        Run 1: keep A,C at start_col=A, row=1 → writes 3-wide bbox (A:C), A1='a', C1='c'.
        Run 2: keep B,D at start_col=B, row=1 → probe B:D row 1; C1 is occupied → BLOCKED.
        """
        with TemporaryDirectory() as td:
            srcA = os.path.join(td, "srcA.xlsx"); srcB = os.path.join(td, "srcB.xlsx")
            dest = os.path.join(td, "dest.xlsx")
            _xlsx(srcA, [["a", "gap", "c"]]); _xlsx(srcB, [["p", "q", "r", "s"]])
            run_sheet(srcA, _cfg(dest, columns="A,C", mode="keep", start_col="A", start_row="1"))
            with pytest.raises(AppError) as ei:
                run_sheet(srcB, _cfg(dest, columns="B,D", mode="keep", start_col="B", start_row="1"))
            assert ei.value.code == DEST_BLOCKED

    def test_keep_non_overlapping_bboxes_both_succeed(self):
        with TemporaryDirectory() as td:
            srcA = os.path.join(td, "srcA.xlsx"); srcB = os.path.join(td, "srcB.xlsx")
            dest = os.path.join(td, "dest.xlsx")
            _xlsx(srcA, [["a", "gap", "c"]]); _xlsx(srcB, [["p", "q", "r", "s"]])
            r1 = run_sheet(srcA, _cfg(dest, columns="A,C", mode="keep", start_col="A", start_row="1"))
            r2 = run_sheet(srcB, _cfg(dest, columns="B,D", mode="keep", start_col="E", start_row="1"))
            assert r1.rows_written == 1
            assert r2.rows_written == 1
