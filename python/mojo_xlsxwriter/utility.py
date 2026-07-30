"""Common A1-reference helpers compatible with XlsxWriter."""

from __future__ import annotations

import re

_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_RANGE = re.compile(
    r"^\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)$"
)


def xl_col_to_name(col: int, col_abs: bool = False) -> str:
    if col < 0:
        return ""
    name = ""
    value = col
    while True:
        name = chr(value % 26 + 65) + name
        value = value // 26 - 1
        if value < 0:
            break
    return ("$" if col_abs else "") + name


def xl_rowcol_to_cell(
    row: int, col: int, row_abs: bool = False, col_abs: bool = False
) -> str:
    if row < 0 or col < 0:
        return ""
    return (
        xl_col_to_name(col, col_abs)
        + ("$" if row_abs else "")
        + str(row + 1)
    )


def xl_range(
    first_row: int, first_col: int, last_row: int, last_col: int
) -> str:
    first = xl_rowcol_to_cell(first_row, first_col)
    if first_row == last_row and first_col == last_col:
        return first
    return f"{first}:{xl_rowcol_to_cell(last_row, last_col)}"


def xl_range_abs(
    first_row: int, first_col: int, last_row: int, last_col: int
) -> str:
    first = xl_rowcol_to_cell(first_row, first_col, True, True)
    if first_row == last_row and first_col == last_col:
        return first
    return f"{first}:{xl_rowcol_to_cell(last_row, last_col, True, True)}"


def xl_cell_to_rowcol(cell: str) -> tuple[int, int]:
    match = _CELL.match(cell)
    if not match:
        raise ValueError(f"invalid cell reference {cell!r}")
    letters, row = match.groups()
    col = 0
    for char in letters.upper():
        col = col * 26 + ord(char) - 64
    return int(row) - 1, col - 1


def cell_args(row, col=None):
    if isinstance(row, str):
        if col is None:
            return xl_cell_to_rowcol(row)
        actual_row, actual_col = xl_cell_to_rowcol(row)
        return actual_row, actual_col
    return int(row), int(col)


def range_args(first_row, first_col=None, last_row=None, last_col=None):
    if isinstance(first_row, str):
        match = _RANGE.match(first_row)
        if match:
            c1, r1, c2, r2 = match.groups()
            row1, col1 = xl_cell_to_rowcol(f"{c1}{r1}")
            row2, col2 = xl_cell_to_rowcol(f"{c2}{r2}")
            return row1, col1, row2, col2
        row, col = xl_cell_to_rowcol(first_row)
        return row, col, row, col
    return int(first_row), int(first_col), int(last_row), int(last_col)
