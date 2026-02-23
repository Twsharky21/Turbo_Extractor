"""
core/rules.py — Row filtering rules engine.

Rules run AFTER row selection and BEFORE column selection/paste shaping.
Rules always reference absolute source columns (A, B, C ...).

Operator behaviour:
  contains  — str(cell).lower() contains target.lower(). None cells → no match.
               Empty target matches everything (every string contains "").
  equals    — case-insensitive string equality after stripping whitespace.
               If both sides are numeric, compares numerically.
               None cell only matches empty-string target.
  <  /  >   — numeric comparison. Non-numeric cell or target → no match (no crash).

Include / Exclude:
  include   — row kept when condition is True.
  exclude   — row kept when condition is False (inverted match).

AND / OR (combine_mode):
  AND       — all rule results must be True for the row to be kept.
  OR        — any rule result being True is enough.
"""
from __future__ import annotations

from typing import Any, List

from .errors import AppError, INVALID_RULE
from .models import Rule
from .parsing import col_letters_to_index


def _safe_numeric(value: Any):
    """Return float(value) or None if conversion fails."""
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _evaluate(cell_value: Any, rule: Rule) -> bool:
    op     = rule.operator
    target = rule.value

    # ── contains ──────────────────────────────────────────────────────────────
    if op == "contains":
        if cell_value is None:
            return False
        return target.lower() in str(cell_value).lower()

    # ── equals ────────────────────────────────────────────────────────────────
    if op == "equals":
        # None cell: only matches if target is also empty
        if cell_value is None:
            return target.strip() == ""

        # Try numeric comparison first (avoids "2" != "2.0" mismatches)
        left_n  = _safe_numeric(cell_value)
        right_n = _safe_numeric(target)
        if left_n is not None and right_n is not None:
            return left_n == right_n

        # Fall back to case-insensitive string comparison
        return str(cell_value).strip().lower() == target.strip().lower()

    # ── < / > ─────────────────────────────────────────────────────────────────
    if op in ("<", ">"):
        left  = _safe_numeric(cell_value)
        right = _safe_numeric(target)
        if left is None or right is None:
            return False
        return left < right if op == "<" else left > right

    raise AppError(INVALID_RULE, f"Unknown operator: {op!r}")


def apply_rules(
    rows: List[List[Any]],
    rules: List[Rule],
    combine_mode: str,
) -> List[List[Any]]:
    """
    Filter rows by the given rules.

    Rules reference absolute source columns (col_letters_to_index converts
    them to 0-based indices). Rows shorter than the referenced column are
    treated as if that cell is None (no match, no crash).
    """
    if not rules:
        return rows

    combine_mode = combine_mode.strip().upper()
    if combine_mode not in ("AND", "OR"):
        raise AppError(INVALID_RULE, f"Bad combine mode: {combine_mode!r}")

    filtered = []

    for row in rows:
        results = []
        for rule in rules:
            col_idx = col_letters_to_index(rule.column) - 1
            cell    = row[col_idx] if col_idx < len(row) else None

            match = _evaluate(cell, rule)

            if rule.mode == "include":
                results.append(match)
            elif rule.mode == "exclude":
                results.append(not match)
            else:
                raise AppError(INVALID_RULE, f"Bad rule mode: {rule.mode!r}")

        keep = all(results) if combine_mode == "AND" else any(results)
        if keep:
            filtered.append(row)

    return filtered
