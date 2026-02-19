\
from __future__ import annotations

import re
from typing import List

from .errors import AppError, BAD_SPEC


_COL_RE = re.compile(r"^[A-Z]+$")
_COL_TOKEN_RE = re.compile(r"^\s*([A-Z]+)\s*(?:-\s*([A-Z]+)\s*)?$")
_ROW_TOKEN_RE = re.compile(r"^\s*(\d+)\s*(?:-\s*(\d+)\s*)?$")


def col_letters_to_index(col: str) -> int:
    """
    Convert Excel column letters to 1-based index (A->1, Z->26, AA->27).
    """
    s = (col or "").strip().upper()
    if not s or not _COL_RE.match(s):
        raise AppError(BAD_SPEC, f"Bad column: {col!r}")
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def col_index_to_letters(n: int) -> str:
    """
    Convert 1-based index to Excel column letters (1->A).
    """
    if n <= 0:
        raise AppError(BAD_SPEC, f"Bad column index: {n}")
    out = []
    x = n
    while x:
        x, rem = divmod(x - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out))


def parse_columns(spec: str) -> List[int]:
    """
    Parse a column spec like: 'A,C,AC-ZZ' into 0-based indices.
    Blank spec => [] (caller interprets as ALL).
    """
    s = (spec or "").strip().upper()
    if s == "":
        return []

    items = []
    for part in [p for p in s.split(",") if p.strip() != ""]:
        m = _COL_TOKEN_RE.match(part)
        if not m:
            raise AppError(BAD_SPEC, f"Bad column token: {part!r}")
        a, b = m.group(1), m.group(2)
        ia = col_letters_to_index(a) - 1
        if b is None:
            items.append(ia)
        else:
            ib = col_letters_to_index(b) - 1
            lo, hi = (ia, ib) if ia <= ib else (ib, ia)
            items.extend(range(lo, hi + 1))

    # unique + sorted
    return sorted(set(items))


def parse_rows(spec: str) -> List[int]:
    """
    Parse a row spec like: '1,1-3,9-80' into 0-based indices.
    Blank spec => [] (caller interprets as ALL).
    """
    s = (spec or "").strip()
    if s == "":
        return []

    items = []
    for part in [p for p in s.split(",") if p.strip() != ""]:
        m = _ROW_TOKEN_RE.match(part)
        if not m:
            raise AppError(BAD_SPEC, f"Bad row token: {part!r}")
        a = int(m.group(1))
        b = int(m.group(2)) if m.group(2) is not None else None
        if a <= 0 or (b is not None and b <= 0):
            raise AppError(BAD_SPEC, f"Row numbers must be >= 1: {part!r}")

        ia = a - 1
        if b is None:
            items.append(ia)
        else:
            ib = b - 1
            lo, hi = (ia, ib) if ia <= ib else (ib, ia)
            items.extend(range(lo, hi + 1))

    return sorted(set(items))
