\
from __future__ import annotations

from typing import List, Any

from .parsing import col_letters_to_index
from .models import Rule
from .errors import AppError, INVALID_RULE


def _safe_numeric(value: Any):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _evaluate(cell_value: Any, rule: Rule) -> bool:
    op = rule.operator
    target = rule.value

    if op == "contains":
        if cell_value is None:
            return False
        return target.lower() in str(cell_value).lower()

    if op == "equals":
        return str(cell_value) == target

    if op in ("<", ">"):
        left = _safe_numeric(cell_value)
        right = _safe_numeric(target)
        if left is None or right is None:
            return False
        if op == "<":
            return left < right
        return left > right

    raise AppError(INVALID_RULE, f"Unknown operator: {op}")


def apply_rules(
    rows: List[List[Any]],
    rules: List[Rule],
    combine_mode: str,
) -> List[List[Any]]:
    """
    Apply rules after row selection and before column selection.
    Rules reference absolute source columns.
    """
    if not rules:
        return rows

    combine_mode = combine_mode.upper()
    if combine_mode not in ("AND", "OR"):
        raise AppError(INVALID_RULE, f"Bad combine mode: {combine_mode}")

    filtered = []

    for row in rows:
        results = []
        for rule in rules:
            col_idx = col_letters_to_index(rule.column) - 1
            if col_idx >= len(row):
                results.append(False)
                continue

            match = _evaluate(row[col_idx], rule)
            if rule.mode == "include":
                results.append(match)
            elif rule.mode == "exclude":
                results.append(not match)
            else:
                raise AppError(INVALID_RULE, f"Bad rule mode: {rule.mode}")

        if combine_mode == "AND":
            if all(results):
                filtered.append(row)
        else:  # OR
            if any(results):
                filtered.append(row)

    return filtered
