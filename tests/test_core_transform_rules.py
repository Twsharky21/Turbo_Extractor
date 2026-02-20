"""Tests for core.transform (row/col selection, pack/keep shaping) and core.rules (filtering)."""
from core.transform import (
    apply_row_selection,
    apply_column_selection,
    shape_pack,
    shape_keep,
)
from core.rules import apply_rules
from core.models import Rule


# ---- Fixtures ----

def _sample_rows():
    return [
        ["A1", "B1", "C1", "D1"],
        ["A2", "B2", "C2", "D2"],
        ["A3", "B3", "C3", "D3"],
        ["A4", "B4", "C4", "D4"],
    ]


def _rules_rows():
    return [
        ["alpha", "10"],
        ["beta",  "20"],
        ["gamma", "30"],
        ["beta",  "5"],
        ["",      None],
    ]


# ---- Transform ----

def test_row_selection_basic():
    rows = _sample_rows()
    result = apply_row_selection(rows, [0, 2])
    assert len(result) == 2
    assert result[1][0] == "A3"


def test_column_selection_basic():
    rows = _sample_rows()
    result = apply_column_selection(rows, [0, 2])
    assert result[0] == ["A1", "C1"]


def test_pack_mode_identity():
    rows = _sample_rows()
    assert shape_pack(rows) == rows


def test_keep_mode_preserves_spacing():
    rows = _sample_rows()
    shaped = shape_keep(rows, [0, 2], [0, 2])

    # Bounding box rows 0–2, cols 0–2
    assert len(shaped) == 3
    assert len(shaped[0]) == 3

    # Row 1 (gap) is blank
    assert shaped[1] == [None, None, None]

    # Selected positions preserved
    assert shaped[0][0] == "A1"
    assert shaped[0][2] == "C1"
    assert shaped[2][0] == "A3"
    assert shaped[2][2] == "C3"


# ---- Rules ----

def test_include_equals():
    rules = [Rule(mode="include", column="A", operator="equals", value="beta")]
    assert len(apply_rules(_rules_rows(), rules, "AND")) == 2


def test_exclude_equals():
    rules = [Rule(mode="exclude", column="A", operator="equals", value="beta")]
    assert len(apply_rules(_rules_rows(), rules, "AND")) == 3


def test_contains_case_insensitive():
    rules = [Rule(mode="include", column="A", operator="contains", value="ALP")]
    assert len(apply_rules(_rules_rows(), rules, "AND")) == 1


def test_numeric_greater_than():
    rules = [Rule(mode="include", column="B", operator=">", value="15")]
    assert len(apply_rules(_rules_rows(), rules, "AND")) == 2


def test_numeric_less_than_safe_on_non_numeric():
    rules = [Rule(mode="include", column="B", operator="<", value="15")]
    assert len(apply_rules(_rules_rows(), rules, "AND")) == 2


def test_and_combine():
    rules = [
        Rule(mode="include", column="A", operator="equals", value="beta"),
        Rule(mode="include", column="B", operator=">", value="10"),
    ]
    assert len(apply_rules(_rules_rows(), rules, "AND")) == 1


def test_or_combine():
    rules = [
        Rule(mode="include", column="A", operator="equals", value="alpha"),
        Rule(mode="include", column="B", operator=">", value="25"),
    ]
    assert len(apply_rules(_rules_rows(), rules, "OR")) == 2
