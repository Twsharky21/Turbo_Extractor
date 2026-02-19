\
from core.errors import AppError


def test_app_error_str_includes_code_message():
    e = AppError("X", "Nope")
    assert str(e).startswith("X: Nope")


def test_app_error_str_includes_details_when_present():
    e = AppError("X", "Nope", {"a": 1})
    s = str(e)
    assert "X: Nope" in s
    assert "a" in s
