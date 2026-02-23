from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any, Dict, Optional


@dataclass
class AppError(Exception):
    """
    User-facing error with a short code and structured details.
    Raise AppError from core modules; GUI should display .message and .details cleanly.
    """
    code: str
    message: str
    details: Optional[Dict[str, Any]] = None

    def __str__(self) -> str:
        if self.details:
            return f"{self.code}: {self.message} ({self.details})"
        return f"{self.code}: {self.message}"


# ── Error codes (keep stable for tests and GUI) ───────────────────────────────

DEST_BLOCKED      = "DEST_BLOCKED"
BAD_SPEC          = "BAD_SPEC"
SOURCE_READ_FAILED = "SOURCE_READ_FAILED"
SHEET_NOT_FOUND   = "SHEET_NOT_FOUND"
INVALID_RULE      = "INVALID_RULE"
FILE_LOCKED       = "FILE_LOCKED"
SAVE_FAILED       = "SAVE_FAILED"
MISSING_DEST_PATH = "MISSING_DEST_PATH"
MISSING_SOURCE_PATH = "MISSING_SOURCE_PATH"


# ── Friendly message lookup ───────────────────────────────────────────────────

def friendly_message(e: AppError) -> str:
    """
    Return a plain-English one-liner suitable for display in a GUI dialog.
    Never exposes raw tracebacks or internal code paths.
    """
    code = e.code
    msg  = e.message or ""

    if code == FILE_LOCKED:
        # Extract filename if available
        fname = ""
        if e.details and "path" in e.details:
            fname = f" ({os.path.basename(e.details['path'])})"
        return f"File is open in another program{fname}. Close it and try again."

    if code == SAVE_FAILED:
        fname = ""
        if e.details and "path" in e.details:
            fname = f" ({os.path.basename(e.details['path'])})"
        # Check if underlying cause was a permission/lock error
        if "permission" in msg.lower() or "locked" in msg.lower() or "access" in msg.lower():
            return f"Could not save — file is open in another program{fname}. Close it and try again."
        return f"Could not save the destination file{fname}. Check that the path is valid and the folder exists."

    if code == DEST_BLOCKED:
        details = e.details or {}
        target  = details.get("target_start", "")
        cols    = details.get("target_data_cols", [])
        col_str = ", ".join(cols) if cols else ""
        blocker = details.get("first_blocker", {})
        b_cell  = f"{blocker.get('col_letter','')}{blocker.get('row','')}" if blocker else ""
        parts = ["Destination already has data in the landing zone."]
        if target:
            parts.append(f"Tried to write at {target}.")
        if col_str:
            parts.append(f"Columns: {col_str}.")
        if b_cell:
            parts.append(f"First blocked cell: {b_cell}.")
        parts.append("Use a different start column/row, or clear the destination first.")
        return " ".join(parts)

    if code == SHEET_NOT_FOUND:
        # Try to extract sheet name from message
        return f"Sheet not found in source file. Check that the sheet name is correct.\n({msg})"

    if code == SOURCE_READ_FAILED:
        if "permission" in msg.lower() or "locked" in msg.lower() or "access" in msg.lower():
            return "Source file is open in another program. Close it and try again."
        if "no such file" in msg.lower() or "not found" in msg.lower():
            return "Source file not found. Check that the file path is correct."
        return f"Could not read the source file. Check that it is a valid XLSX or CSV.\n({msg})"

    if code == BAD_SPEC:
        if "column" in msg.lower():
            return f"Invalid column specification. Use letters like A, B, A-C, or A,C,E.\n({msg})"
        if "row" in msg.lower():
            return f"Invalid row specification. Use numbers like 1, 1-10, or 1,5,10.\n({msg})"
        return f"Invalid setting — please check your configuration.\n({msg})"

    if code == "BAD_SOURCE_START_ROW":
        return f"Source Start Row must be a number (1 or higher).\n({msg})"

    if code == MISSING_DEST_PATH:
        return "No destination file path set. Enter a file path in the Destination field."

    if code == MISSING_SOURCE_PATH:
        return "No source file path set. Select a source file first."

    if code == INVALID_RULE:
        return f"A filter rule has an invalid setting. Check the Rules section.\n({msg})"

    # Fallback — clean up the raw message, never show raw tracebacks
    clean = msg.splitlines()[0] if msg else "An unexpected error occurred."
    return clean
