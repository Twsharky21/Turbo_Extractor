\
from __future__ import annotations

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


# Common error codes (keep stable for tests and GUI)
DEST_BLOCKED = "DEST_BLOCKED"
BAD_SPEC = "BAD_SPEC"
SOURCE_READ_FAILED = "SOURCE_READ_FAILED"
SHEET_NOT_FOUND = "SHEET_NOT_FOUND"
INVALID_RULE = "INVALID_RULE"
