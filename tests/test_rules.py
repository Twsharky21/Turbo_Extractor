"""
test_rules.py — Comprehensive tests for core.rules.apply_rules.

Covers:
  - contains: basic, case-insensitive, empty target, None cell
  - equals: string, case-insensitive, numeric, None cell, mixed types
  - < / >: numeric, non-numeric cell skipped, non-numeric target skipped
  - include vs exclude inversion
  - AND / OR combine modes
  - Multiple rules interactions
  - Edge cases: empty row list, column beyond row width, bad mode, bad operator
"""
from __future__ import annotations

import pytest

from core.rules import apply_rules
from core.models import Rule
from core.errors import AppError, INVALID_RULE


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _rule(op, val, col="A", mode="include"):
    return Rule(mode=mode, column=col, operator=op, value=val)


def _apply(rows, op, val, col="A", mode="include", combine="AND"):
    return apply_rules(rows, [_rule(op, val, col=col, mode=mode)], combine)


# ══════════════════════════════════════════════════════════════════════════════
# NO RULES
# ══════════════════════════════════════════════════════════════════════════════

def test_no_rules_and_returns_all():
    rows = [["a"], ["b"], ["c"]]
    assert apply_rules(rows, [], "AND") == rows


def test_no_rules_or_returns_all():
    rows = [["a"], ["b"]]
    assert apply_rules(rows, [], "OR") == rows


def test_empty_input_rows_returns_empty():
    assert apply_rules([], [_rule("equals", "x")], "AND") == []


# ══════════════════════════════════════════════════════════════════════════════
# CONTAINS
# ══════════════════════════════════════════════════════════════════════════════

def test_contains_basic_match():
    rows = [["apple"], ["banana"], ["cherry"]]
    result = _apply(rows, "contains", "an")
    assert len(result) == 1
    assert result[0][0] == "banana"


def test_contains_case_insensitive():
    rows = [["Green Apple"], ["banana"]]
    result = _apply(rows, "contains", "apple")
    assert len(result) == 1
    assert result[0][0] == "Green Apple"


def test_contains_empty_target_matches_all():
    rows = [["alpha"], ["beta"], ["gamma"]]
    result = _apply(rows, "contains", "")
    assert len(result) == 3


def test_contains_none_cell_no_match():
    rows = [[None], ["hello"]]
    result = _apply(rows, "contains", "hello")
    assert len(result) == 1
    assert result[0][0] == "hello"


def test_contains_numeric_cell_matches_substring():
    rows = [[12345], [99]]
    result = _apply(rows, "contains", "23")
    assert len(result) == 1
    assert result[0][0] == 12345


def test_contains_exclude_inverts():
    rows = [["apple"], ["banana"], ["cherry"]]
    result = _apply(rows, "contains", "an", mode="exclude")
    names = [r[0] for r in result]
    assert "banana" not in names
    assert "apple" in names and "cherry" in names


# ══════════════════════════════════════════════════════════════════════════════
# EQUALS
# ══════════════════════════════════════════════════════════════════════════════

def test_equals_exact_string_match():
    rows = [["alpha"], ["beta"], ["gamma"]]
    result = _apply(rows, "equals", "alpha")
    assert len(result) == 1
    assert result[0][0] == "alpha"


def test_equals_case_insensitive():
    rows = [["Alpha"], ["BETA"], ["gamma"]]
    result = _apply(rows, "equals", "alpha")
    assert len(result) == 1
    assert result[0][0] == "Alpha"


def test_equals_strips_whitespace():
    rows = [["  alpha  "], ["beta"]]
    result = _apply(rows, "equals", "alpha")
    assert len(result) == 1


def test_equals_numeric_cell_matches_numeric_string():
    """Integer 2 should equal rule value "2"."""
    rows = [[1], [2], [3]]
    result = _apply(rows, "equals", "2")
    assert len(result) == 1
    assert result[0][0] == 2


def test_equals_float_cell_matches_int_string():
    """Float 2.0 should equal rule value "2"."""
    rows = [[2.0], [3.5]]
    result = _apply(rows, "equals", "2")
    assert len(result) == 1
    assert result[0][0] == 2.0


def test_equals_none_cell_matches_empty_string_target():
    rows = [[None], ["hello"]]
    result = _apply(rows, "equals", "")
    assert len(result) == 1
    assert result[0][0] is None


def test_equals_none_cell_does_not_match_nonempty_target():
    rows = [[None], ["hello"]]
    result = _apply(rows, "equals", "None")
    assert len(result) == 0


def test_equals_zero_matches_zero_string():
    rows = [[0], [1], [2]]
    result = _apply(rows, "equals", "0")
    assert len(result) == 1
    assert result[0][0] == 0


def test_equals_exclude_inverts():
    rows = [["alpha"], ["beta"], ["gamma"]]
    result = _apply(rows, "equals", "beta", mode="exclude")
    names = [r[0] for r in result]
    assert "beta" not in names
    assert len(result) == 2


# ══════════════════════════════════════════════════════════════════════════════
# GREATER THAN / LESS THAN
# ══════════════════════════════════════════════════════════════════════════════

def test_greater_than_numeric():
    rows = [[10], [20], [30]]
    result = _apply(rows, ">", "15")
    assert len(result) == 2
    assert all(r[0] > 15 for r in result)


def test_less_than_numeric():
    rows = [[10], [20], [30]]
    result = _apply(rows, "<", "25")
    assert len(result) == 2
    assert all(r[0] < 25 for r in result)


def test_greater_than_with_floats():
    rows = [[1.5], [2.5], [3.5]]
    result = _apply(rows, ">", "2.0")
    assert len(result) == 2


def test_less_than_negative_numbers():
    rows = [[-10], [-5], [0], [5]]
    result = _apply(rows, "<", "-3")
    assert len(result) == 2
    assert all(r[0] < -3 for r in result)


def test_greater_than_non_numeric_cell_skipped():
    """Text cells that can't be coerced to float should not match."""
    rows = [["text"], [100], [200]]
    result = _apply(rows, ">", "50")
    assert len(result) == 2
    assert all(isinstance(r[0], (int, float)) for r in result)


def test_less_than_none_cell_skipped():
    rows = [[None], [5], [50]]
    result = _apply(rows, "<", "10")
    assert len(result) == 1
    assert result[0][0] == 5


def test_greater_than_non_numeric_target_matches_nothing():
    rows = [[100], [200]]
    result = _apply(rows, ">", "not_a_number")
    assert len(result) == 0


def test_greater_than_exclude():
    rows = [[10], [20], [30]]
    result = _apply(rows, ">", "15", mode="exclude")
    assert len(result) == 1
    assert result[0][0] == 10


# ══════════════════════════════════════════════════════════════════════════════
# COLUMN SELECTION
# ══════════════════════════════════════════════════════════════════════════════

def test_rule_on_non_first_column():
    rows = [["keep", "yes", 1],
            ["drop", "no",  2],
            ["keep", "yes", 3]]
    result = _apply(rows, "equals", "yes", col="B")
    assert len(result) == 2
    assert all(r[1] == "yes" for r in result)


def test_rule_column_beyond_row_width_no_match_no_crash():
    rows = [["a"], ["b"]]
    result = _apply(rows, "equals", "a", col="Z")
    assert result == []


def test_rule_on_column_c():
    rows = [["x", "y", "target"],
            ["x", "y", "other"]]
    result = _apply(rows, "equals", "target", col="C")
    assert len(result) == 1
    assert result[0][2] == "target"


# ══════════════════════════════════════════════════════════════════════════════
# AND / OR COMBINE MODES
# ══════════════════════════════════════════════════════════════════════════════

def test_and_both_true_keeps_row():
    rows = [["alpha", 20], ["beta", 5]]
    rules = [
        _rule("equals", "alpha", col="A"),
        _rule(">", "10", col="B"),
    ]
    result = apply_rules(rows, rules, "AND")
    assert len(result) == 1
    assert result[0][0] == "alpha"


def test_and_one_false_drops_row():
    rows = [["alpha", 5], ["beta", 20]]
    rules = [
        _rule("equals", "alpha", col="A"),
        _rule(">", "10", col="B"),
    ]
    result = apply_rules(rows, rules, "AND")
    assert len(result) == 0


def test_or_either_true_keeps_row():
    rows = [["alpha", 5], ["beta", 20], ["gamma", 1]]
    rules = [
        _rule("equals", "alpha", col="A"),
        _rule(">", "10", col="B"),
    ]
    result = apply_rules(rows, rules, "OR")
    assert len(result) == 2
    values = [r[0] for r in result]
    assert "alpha" in values and "beta" in values


def test_or_neither_true_drops_row():
    rows = [["gamma", 1]]
    rules = [
        _rule("equals", "alpha", col="A"),
        _rule(">", "10", col="B"),
    ]
    result = apply_rules(rows, rules, "OR")
    assert len(result) == 0


def test_and_three_rules_all_must_pass():
    rows = [
        ["alpha", 20, "yes"],
        ["alpha", 20, "no"],
        ["beta",  20, "yes"],
    ]
    rules = [
        _rule("equals",   "alpha", col="A"),
        _rule(">",        "10",    col="B"),
        _rule("equals",   "yes",   col="C"),
    ]
    result = apply_rules(rows, rules, "AND")
    assert len(result) == 1
    assert result[0] == ["alpha", 20, "yes"]


# ══════════════════════════════════════════════════════════════════════════════
# MIXED INCLUDE / EXCLUDE
# ══════════════════════════════════════════════════════════════════════════════

def test_mixed_include_and_exclude_and_mode():
    """Include col A equals alpha AND exclude col B equals 20."""
    rows = [
        ["alpha", 10],
        ["alpha", 20],
        ["beta",  10],
    ]
    rules = [
        _rule("equals", "alpha", col="A", mode="include"),
        _rule("equals", "20",    col="B", mode="exclude"),
    ]
    result = apply_rules(rows, rules, "AND")
    assert len(result) == 1
    assert result[0] == ["alpha", 10]


def test_all_exclude_or_keeps_rows_not_matching_any():
    rows = [["alpha"], ["beta"], ["gamma"]]
    rules = [
        _rule("equals", "alpha", mode="exclude"),
        _rule("equals", "beta",  mode="exclude"),
    ]
    result = apply_rules(rows, rules, "OR")
    # OR with exclude: kept if ANY exclusion does NOT match
    # gamma: exclude alpha → True (not alpha), exclude beta → True (not beta) → OR True → kept
    # alpha: exclude alpha → False, exclude beta → True → OR True → kept
    # beta:  exclude alpha → True, exclude beta → False → OR True → kept
    assert len(result) == 3


def test_all_exclude_and_drops_any_matching_row():
    rows = [["alpha"], ["beta"], ["gamma"]]
    rules = [
        _rule("equals", "alpha", mode="exclude"),
        _rule("equals", "beta",  mode="exclude"),
    ]
    result = apply_rules(rows, rules, "AND")
    # AND with exclude: kept only if ALL exclusions True (no rule matched)
    assert len(result) == 1
    assert result[0][0] == "gamma"


# ══════════════════════════════════════════════════════════════════════════════
# ERROR PATHS
# ══════════════════════════════════════════════════════════════════════════════

def test_unknown_operator_raises_invalid_rule():
    with pytest.raises(AppError) as ei:
        apply_rules([["a"]], [_rule("LIKE", "x")], "AND")
    assert ei.value.code == INVALID_RULE


def test_bad_rule_mode_raises_invalid_rule():
    rule = Rule(mode="maybe", column="A", operator="equals", value="x")
    with pytest.raises(AppError) as ei:
        apply_rules([["a"]], [rule], "AND")
    assert ei.value.code == INVALID_RULE


def test_bad_combine_mode_raises_invalid_rule():
    with pytest.raises(AppError) as ei:
        apply_rules([["a"]], [_rule("equals", "a")], "XOR")
    assert ei.value.code == INVALID_RULE
