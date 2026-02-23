"""
test_friendly_errors.py — Tests for friendly_message() in core.errors.

Verifies that all error codes produce readable plain-English messages
with no raw tracebacks or code gibberish.
"""
from __future__ import annotations

import pytest

from core.errors import (
    AppError,
    friendly_message,
    DEST_BLOCKED,
    BAD_SPEC,
    SOURCE_READ_FAILED,
    SHEET_NOT_FOUND,
    INVALID_RULE,
    FILE_LOCKED,
    SAVE_FAILED,
    MISSING_DEST_PATH,
)


def test_file_locked_message_mentions_program():
    e = AppError(FILE_LOCKED, "File is locked", {"path": "C:/data/output.xlsx"})
    msg = friendly_message(e)
    assert "open in another program" in msg
    assert "output.xlsx" in msg


def test_file_locked_message_no_path():
    e = AppError(FILE_LOCKED, "File is locked")
    msg = friendly_message(e)
    assert "open in another program" in msg


def test_save_failed_permission_mentions_close():
    e = AppError(SAVE_FAILED, "PermissionError: access denied", {"path": "out.xlsx"})
    msg = friendly_message(e)
    assert "open in another program" in msg.lower() or "close" in msg.lower()


def test_save_failed_generic_message():
    e = AppError(SAVE_FAILED, "disk full", {"path": "out.xlsx"})
    msg = friendly_message(e)
    assert "save" in msg.lower() or "file" in msg.lower()


def test_dest_blocked_message_includes_cell():
    e = AppError(DEST_BLOCKED, "blocked", {
        "target_start": "B3",
        "target_data_cols": ["B", "D"],
        "first_blocker": {"row": 3, "col": 2, "col_letter": "B"},
    })
    msg = friendly_message(e)
    assert "B3" in msg
    assert "B3" in msg or "landing" in msg.lower()
    assert "B" in msg


def test_dest_blocked_message_no_details():
    e = AppError(DEST_BLOCKED, "blocked")
    msg = friendly_message(e)
    assert "data" in msg.lower() or "blocked" in msg.lower()


def test_sheet_not_found_message():
    e = AppError(SHEET_NOT_FOUND, "Sheet 'Data' not found")
    msg = friendly_message(e)
    assert "sheet" in msg.lower()


def test_source_read_failed_permission():
    e = AppError(SOURCE_READ_FAILED, "PermissionError reading file")
    msg = friendly_message(e)
    assert "open in another program" in msg or "permission" in msg.lower()


def test_source_read_failed_not_found():
    e = AppError(SOURCE_READ_FAILED, "no such file or directory")
    msg = friendly_message(e)
    assert "not found" in msg.lower() or "path" in msg.lower()


def test_bad_spec_column():
    e = AppError(BAD_SPEC, "Bad column token: '??'")
    msg = friendly_message(e)
    assert "column" in msg.lower()


def test_bad_spec_row():
    e = AppError(BAD_SPEC, "Bad row token: 'abc'")
    msg = friendly_message(e)
    assert "row" in msg.lower()


def test_missing_dest_path_message():
    e = AppError(MISSING_DEST_PATH, "Destination file path is blank.")
    msg = friendly_message(e)
    assert "destination" in msg.lower() or "path" in msg.lower()


def test_invalid_rule_message():
    e = AppError(INVALID_RULE, "Unknown operator: 'LIKE'")
    msg = friendly_message(e)
    assert "rule" in msg.lower()


def test_friendly_message_never_returns_empty():
    """Every error code must return a non-empty string."""
    codes = [DEST_BLOCKED, BAD_SPEC, SOURCE_READ_FAILED, SHEET_NOT_FOUND,
             INVALID_RULE, FILE_LOCKED, SAVE_FAILED, MISSING_DEST_PATH]
    for code in codes:
        e = AppError(code, "some message")
        msg = friendly_message(e)
        assert isinstance(msg, str)
        assert len(msg.strip()) > 0, f"Empty message for code {code}"


def test_unknown_error_code_returns_first_line_of_message():
    e = AppError("UNKNOWN_CODE", "Something went wrong\nwith a traceback here")
    msg = friendly_message(e)
    assert "Something went wrong" in msg
    assert "traceback" not in msg
