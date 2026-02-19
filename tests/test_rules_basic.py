\
from core.rules import apply_rules
from core.models import Rule


def sample_rows():
    return [
        ["alpha", "10"],
        ["beta", "20"],
        ["gamma", "30"],
        ["beta", "5"],
        ["", None],
    ]


def test_include_equals():
    rows = sample_rows()
    rules = [Rule(mode="include", column="A", operator="equals", value="beta")]
    result = apply_rules(rows, rules, "AND")
    assert len(result) == 2


def test_exclude_equals():
    rows = sample_rows()
    rules = [Rule(mode="exclude", column="A", operator="equals", value="beta")]
    result = apply_rules(rows, rules, "AND")
    assert len(result) == 3


def test_contains_case_insensitive():
    rows = sample_rows()
    rules = [Rule(mode="include", column="A", operator="contains", value="ALP")]
    result = apply_rules(rows, rules, "AND")
    assert len(result) == 1


def test_numeric_greater_than():
    rows = sample_rows()
    rules = [Rule(mode="include", column="B", operator=">", value="15")]
    result = apply_rules(rows, rules, "AND")
    assert len(result) == 2


def test_numeric_less_than_safe_on_non_numeric():
    rows = sample_rows()
    rules = [Rule(mode="include", column="B", operator="<", value="15")]
    result = apply_rules(rows, rules, "AND")
    assert len(result) == 2


def test_and_combine():
    rows = sample_rows()
    rules = [
        Rule(mode="include", column="A", operator="equals", value="beta"),
        Rule(mode="include", column="B", operator=">", value="10"),
    ]
    result = apply_rules(rows, rules, "AND")
    assert len(result) == 1


def test_or_combine():
    rows = sample_rows()
    rules = [
        Rule(mode="include", column="A", operator="equals", value="alpha"),
        Rule(mode="include", column="B", operator=">", value="25"),
    ]
    result = apply_rules(rows, rules, "OR")
    assert len(result) == 2
