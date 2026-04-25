"""
utils.py — shared pure-Python helpers used across excel_tools, main, and tests.

Keep this module free of win32com, streamlit, and openai dependencies so it
can be imported in any context including unit tests.
"""
from __future__ import annotations


def col_letter(idx: int) -> str:
    """
    Convert a 1-based column index to an Excel column letter string.

    Examples:
        col_letter(1)  -> 'A'
        col_letter(26) -> 'Z'
        col_letter(27) -> 'AA'
        col_letter(702) -> 'ZZ'
    """
    if idx < 1:
        raise ValueError(f"Column index must be >= 1, got {idx}")
    result = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        result = chr(65 + r) + result
    return result


def col_index(letter: str) -> int:
    """
    Convert an Excel column letter string to a 1-based column index.

    Examples:
        col_index('A')  -> 1
        col_index('Z')  -> 26
        col_index('AA') -> 27
    """
    letter = letter.upper().strip()
    if not letter.isalpha():
        raise ValueError(f"Invalid column letter: {letter!r}")
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - 64)
    return result
