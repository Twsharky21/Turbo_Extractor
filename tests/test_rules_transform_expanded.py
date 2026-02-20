"""
Expanded tests for core.rules and core.transform — error paths,
edge cases, and OR/AND combine logic not covered by existing tests.
"""
import pytest

from core.rules import apply_rules
from core.models import Rule
from core.errors import AppError, INVALID_RULE, BAD_SPEC
from core.transform import apply_row_selection, apply_column_selection, shape_keep


# ---- Fixtures ----

def _rows():
    return [
        ["alpha", "10", "tag_a"],
        ["beta",  "20", "tag_b"],
        ["gamma", "30", "tag_a"],
        ["beta",  "5",  "tag_b"],
        ["",      None, None],
    ]


# ============================================================
# RULES — ERROR PATHS
# ============================================================

def test_unknown_operator_raises_invalid_rule():
    rules = [Rule(mode="include", column="A", operator="LIKE", value="x")]
    with pytest.raises(AppError) as ei:
        apply_rules(_rows(), rules, "AND")
    assert ei.value.code == INVALID_RULE


def test_bad_rule_mode_raises_invalid_rule():
    rules = [Rule(mode="maybe", column="A", operator="equals", value="alpha")]
    with pytest.raises(AppError) as ei:
        apply_rules(_rows(), rules, "AND")
    assert ei.value.code == INVALID_RULE


def test_bad_combine_mode_raises_invalid_rule():
    rules = [Rule(mode="include", column="A", operator="equals", value="alpha")]
    with pytest.raises(AppError) as ei:
        apply_rules(_rows(), rules, "XOR")
    assert ei.value.code == INVALID_RULE


def test_rule_column_beyond_row_width_counts_as_false_not_crash():
    """Column Z is way beyond our 3-column rows; should not crash, just not match."""
    rules = [Rule(mode="include", column="Z", operator="equals", value="anything")]
    result = apply_rules(_rows(), rules, "AND")
    assert result == []   # nothing matches


def test_or_combine_all_rules_false_returns_empty():
    rules = [
        Rule(mode="include", column="A", operator="equals", value="NO_MATCH_1"),
        Rule(mode="include", column="A", operator="equals", value="NO_MATCH_2"),
    ]
    result = apply_rules(_rows(), rules, "OR")
    assert result == []


# ============================================================
# RULES — COMBINE LOGIC
# ============================================================

def test_and_combine_all_must_pass():
    rules = [
        Rule(mode="include", column="A", operator="equals", value="beta"),
        Rule(mode="include", column="B", operator=">", value="10"),
    ]
    result = apply_rules(_rows(), rules, "AND")
    assert len(result) == 1
    assert result[0][0] == "beta"
    assert result[0][1] == "20"


def test_or_combine_either_passes():
    rules = [
        Rule(mode="include", column="A", operator="equals", value="alpha"),
        Rule(mode="include", column="B", operator=">", value="25"),
    ]
    result = apply_rules(_rows(), rules, "OR")
    assert len(result) == 2


def test_exclude_and_include_combined():
    """Include tag_a AND exclude beta → only alpha and gamma."""
    rules = [
        Rule(mode="include", column="C", operator="equals", value="tag_a"),
        Rule(mode="exclude", column="A", operator="equals", value="beta"),
    ]
    result = apply_rules(_rows(), rules, "AND")
    names = [r[0] for r in result]
    assert "alpha" in names
    assert "gamma" in names
    assert "beta" not in names


def test_no_rules_returns_all_rows():
    rows = _rows()
    assert apply_rules(rows, [], "AND") is rows


def test_contains_operator():
    rules = [Rule(mode="include", column="A", operator="contains", value="lph")]
    result = apply_rules(_rows(), rules, "AND")
    assert len(result) == 1
    assert result[0][0] == "alpha"


def test_less_than_operator():
    rules = [Rule(mode="include", column="B", operator="<", value="15")]
    result = apply_rules(_rows(), rules, "AND")
    # "10" < 15 and "5" < 15 → 2 rows
    assert len(result) == 2


def test_non_numeric_value_in_numeric_comparison_excluded_safely():
    """Rows with non-numeric cell values in < / > comparisons are excluded (not crash)."""
    rows = [["x", "not_a_number"], ["y", "42"]]
    rules = [Rule(mode="include", column="B", operator=">", value="10")]
    result = apply_rules(rows, rules, "AND")
    assert len(result) == 1
    assert result[0][0] == "y"


def test_none_cell_value_in_contains_returns_false():
    rows = [["a", None], ["b", "hello"]]
    rules = [Rule(mode="include", column="B", operator="contains", value="he")]
    result = apply_rules(rows, rules, "AND")
    assert len(result) == 1
    assert result[0][0] == "b"


# ============================================================
# TRANSFORM — EDGE CASES
# ============================================================

def test_row_selection_out_of_bounds_indices_ignored():
    rows = [["r1"], ["r2"]]
    result = apply_row_selection(rows, [0, 99])
    assert len(result) == 1
    assert result[0] == ["r1"]


def test_column_selection_out_of_bounds_index_gives_none():
    rows = [["a", "b"]]
    result = apply_column_selection(rows, [0, 5])
    assert result[0] == ["a", None]


def test_column_selection_empty_rows_returns_empty():
    assert apply_column_selection([], [0, 1]) == []


def test_row_selection_empty_indices_returns_all():
    rows = [["a"], ["b"]]
    assert apply_row_selection(rows, []) is rows


def test_shape_keep_single_row_single_col():
    rows = [["only"]]
    shaped = shape_keep(rows, [0], [0])
    assert shaped == [["only"]]


def test_shape_keep_empty_original_returns_empty():
    assert shape_keep([], [0], [0]) == []


def test_shape_keep_all_indices_same_as_full_table():
    rows = [["a", "b"], ["c", "d"]]
    shaped = shape_keep(rows, [0, 1], [0, 1])
    assert shaped == rows


def test_shape_keep_gap_row_is_all_none():
    rows = [["a", "b"], ["c", "d"], ["e", "f"]]
    # Select rows 0 and 2, col 0 only → bounding box 3 rows x 1 col
    shaped = shape_keep(rows, [0, 2], [0])
    assert len(shaped) == 3
    assert shaped[0] == ["a"]
    assert shaped[1] == [None]   # gap
    assert shaped[2] == ["e"]
