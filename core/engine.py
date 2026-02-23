"""
core/engine.py — Re-export shim for backward compatibility.

All logic has been split into:
  core.runner  — single-sheet extraction (run_sheet)
  core.batch   — batch execution loop (run_all, RunItem)

Existing imports from core.engine continue to work unchanged:
  from core.engine import run_sheet, run_all
"""
from __future__ import annotations

from .runner import run_sheet
from .batch import run_all, RunItem

__all__ = ["run_sheet", "run_all", "RunItem"]
