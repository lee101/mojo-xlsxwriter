"""Worksheet data model for the shared-string XLSX writer."""

from __future__ import annotations

import datetime as dt
import math
import re
from dataclasses import dataclass
from fractions import Fraction
from numbers import Real
from urllib.parse import quote

from .exceptions import OverlappingRange
from .format import Format
from .utility import cell_args, range_args, xl_range

MAX_ROWS = 1_048_576
MAX_COLS = 16_384
MAX_STRING = 32_767
_URL = re.compile(r"^(?:https?://|ftp://|mailto:|internal:|external:)", re.I)


@dataclass(slots=True)
class Cell:
    kind: str
    value: object
    cell_format: Format | None = None


def _datetime_to_excel(value, date_1904=False, remove_timezone=False):
    original = value
    epoch = dt.datetime(1904, 1, 1) if date_1904 else dt.datetime(1899, 12, 31)
    is_timedelta = isinstance(value, dt.timedelta)
    if isinstance(value, dt.datetime):
        if value.tzinfo is not None:
            if not remove_timezone:
                raise TypeError(
                    "Excel doesn't support timezones in datetimes. "
                    "Set remove_timezone=True."
                )
            value = value.replace(tzinfo=None)
        delta = value - epoch
    elif isinstance(value, dt.date):
        delta = dt.datetime.fromordinal(value.toordinal()) - epoch
    elif isinstance(value, dt.time):
        if value.tzinfo is not None:
            if not remove_timezone:
                raise TypeError(
                    "Excel doesn't support timezones in datetimes. "
                    "Set remove_timezone=True."
                )
            value = value.replace(tzinfo=None)
        delta = dt.datetime.combine(epoch, value) - epoch
    elif is_timedelta:
        delta = value
    else:
        raise TypeError("Unknown or unsupported datetime type")
    serial = delta.days + (delta.seconds + delta.microseconds / 1e6) / 86400
    if isinstance(original, dt.datetime) and value.date() == dt.date(1900, 1, 1):
        serial -= 1
    if not date_1904 and not is_timedelta and serial > 59:
        serial += 1
    return serial


class Worksheet:
    def __init__(self, workbook, name: str, index: int):
        self.workbook = workbook
        self.name = name
        self.index = index
        self.cells: dict[int, dict[int, Cell]] = {}
        self.row_options: dict[int, tuple] = {}
        self.col_options: dict[int, tuple] = {}
        self.merges: list[tuple[int, int, int, int]] = []
        self._merged_cells: set[tuple[int, int]] = set()
        self.panes: tuple[int, int, int, int] | None = None
        self.autofilter_ref: str | None = None
        self.hyperlinks: list[dict] = []
        self.hidden = False
        self.selected = index == 0
        self.active = index == 0
        self.dim_rowmin: int | None = None
        self.dim_rowmax: int | None = None
        self.dim_colmin: int | None = None
        self.dim_colmax: int | None = None

    def _coordinates(self, row, col):
        if isinstance(row, str):
            return cell_args(row)
        return int(row), int(col)

    def _check(self, row: int, col: int, update=True) -> bool:
        if row < 0 or col < 0 or row >= MAX_ROWS or col >= MAX_COLS:
            return True
        if update:
            self.dim_rowmin = row if self.dim_rowmin is None else min(self.dim_rowmin, row)
            self.dim_rowmax = row if self.dim_rowmax is None else max(self.dim_rowmax, row)
            self.dim_colmin = col if self.dim_colmin is None else min(self.dim_colmin, col)
            self.dim_colmax = col if self.dim_colmax is None else max(self.dim_colmax, col)
        return False

    def _store(self, row, col, cell):
        if self._check(row, col):
            return -1
        self.cells.setdefault(row, {})[col] = cell
        return 0

    def write(self, row: int, col: int, *args):
        if isinstance(row, str):
            if not args:
                args = (col,)
            elif len(args) == 1:
                args = (col, args[0])
            row, col = cell_args(row)
        if not args:
            raise TypeError("write() takes a value argument")
        value = args[0]
        cell_format = args[1] if len(args) > 1 else None
        if value is None:
            return self.write_blank(row, col, value, cell_format)
        if isinstance(value, str):
            if self.workbook.strings_to_formulas and value.startswith("="):
                return self.write_formula(row, col, value, cell_format)
            if (
                self.workbook.strings_to_urls
                and value.find(":", 0, 9) >= 0
                and _URL.match(value)
            ):
                return self.write_url(row, col, value, cell_format)
            if self.workbook.strings_to_numbers:
                try:
                    return self.write_number(row, col, float(value), cell_format)
                except ValueError:
                    pass
            return self.write_string(row, col, value, cell_format)
        if isinstance(value, bool):
            return self.write_boolean(row, col, value, cell_format)
        if isinstance(value, (dt.datetime, dt.date, dt.time, dt.timedelta)):
            return self.write_datetime(row, col, value, cell_format)
        if isinstance(value, (Real, Fraction)):
            return self.write_number(row, col, value, cell_format)
        raise TypeError(f"Unsupported type {type(value)!r} in write()")

    def write_string(self, row, col=None, string=None, cell_format=None):
        if isinstance(row, str):
            cell_format = string if isinstance(string, Format) else cell_format
            string = col
            row, col = cell_args(row)
        if not isinstance(string, str):
            raise TypeError("write_string() requires a string")
        if self._check(row, col):
            return -1
        result = 0
        if len(string) > MAX_STRING:
            string = string[:MAX_STRING]
            result = -2
        occurrence = self.workbook._add_string(string)
        self.cells.setdefault(row, {})[col] = Cell("string", occurrence, cell_format)
        return result

    def write_number(self, row, col=None, number=None, cell_format=None):
        if isinstance(row, str):
            cell_format = number if isinstance(number, Format) else cell_format
            number = col
            row, col = cell_args(row)
        number = float(number) if isinstance(number, Fraction) else number
        if isinstance(number, float) and (math.isnan(number) or math.isinf(number)):
            if not self.workbook.nan_inf_to_errors:
                raise TypeError(
                    "NAN/INF not supported in write_number() without "
                    "'nan_inf_to_errors' Workbook() option"
                )
            if math.isnan(number):
                return self.write_formula(row, col, "#NUM!", cell_format, "#NUM!")
            formula = "1/0" if number > 0 else "-1/0"
            return self.write_formula(row, col, formula, cell_format, "#DIV/0!")
        return self._store(row, col, Cell("number", number, cell_format))

    def write_boolean(self, row, col=None, boolean=None, cell_format=None):
        if isinstance(row, str):
            cell_format = boolean if isinstance(boolean, Format) else cell_format
            boolean = col
            row, col = cell_args(row)
        return self._store(row, col, Cell("boolean", bool(boolean), cell_format))

    def write_blank(self, row, col=None, blank=None, cell_format=None):
        if isinstance(row, str):
            cell_format = blank if isinstance(blank, Format) else cell_format
            blank = col
            row, col = cell_args(row)
        if cell_format is None:
            return 0
        return self._store(row, col, Cell("blank", None, cell_format))

    def write_formula(self, row, col=None, formula=None, cell_format=None, value=0):
        if isinstance(row, str):
            if cell_format is not None and not isinstance(cell_format, Format):
                value = cell_format
            cell_format = formula if isinstance(formula, Format) else None
            formula = col
            row, col = cell_args(row)
        if not formula:
            return -1
        if formula.startswith("="):
            formula = formula[1:]
        return self._store(row, col, Cell("formula", (formula, value), cell_format))

    def write_datetime(self, row, col=None, date=None, cell_format=None):
        if isinstance(row, str):
            cell_format = date if isinstance(date, Format) else cell_format
            date = col
            row, col = cell_args(row)
        serial = _datetime_to_excel(
            date, self.workbook.date_1904, self.workbook.remove_timezone
        )
        if cell_format is None:
            cell_format = self.workbook.default_date_format
        return self._store(row, col, Cell("number", serial, cell_format))

    def write_url(self, row, col=None, url=None, cell_format=None, string=None, tip=None):
        if isinstance(row, str):
            actual_url = col
            actual_format = url if isinstance(url, Format) else None
            actual_string = cell_format if isinstance(cell_format, str) else None
            actual_tip = string if isinstance(string, str) else tip
            url = col
            cell_format, string, tip = actual_format, actual_string, actual_tip
            row, col = cell_args(row)
        if self._check(row, col, update=False):
            return -1
        if len(url.split("#", 1)[0]) > self.workbook.max_url_length:
            return -3
        display = string or url
        if url.lower().startswith("mailto:") and string is None:
            display = url[7:]
        if cell_format is None:
            cell_format = self.workbook.default_url_format
        result = self.write_string(row, col, display, cell_format)
        self.hyperlinks.append({"row": row, "col": col, "url": url, "tip": tip})
        return result

    def write_row(self, row, col=None, data=None, cell_format=None):
        if isinstance(row, str):
            cell_format = data if isinstance(data, Format) else cell_format
            data = col
            row, col = cell_args(row)
        for value in data:
            if (
                isinstance(value, str)
                and not self.workbook.strings_to_numbers
                and not (
                    self.workbook.strings_to_formulas and value.startswith("=")
                )
                and not (
                    self.workbook.strings_to_urls and value.find(":", 0, 9) >= 0
                )
            ):
                error = self.write_string(row, col, value, cell_format)
            else:
                error = self.write(row, col, value, cell_format)
            if error:
                return error
            col += 1
        return 0

    def write_column(self, row, col=None, data=None, cell_format=None):
        if isinstance(row, str):
            cell_format = data if isinstance(data, Format) else cell_format
            data = col
            row, col = cell_args(row)
        for value in data:
            error = self.write(row, col, value, cell_format)
            if error:
                return error
            row += 1
        return 0

    def set_column(self, first_col, last_col=None, width=None, cell_format=None, options=None):
        if isinstance(first_col, str):
            actual_width = last_col
            actual_format = width if isinstance(width, Format) else None
            actual_options = cell_format if isinstance(cell_format, dict) else options
            column_range = first_col.replace("$", "")
            parts = column_range.split(":")
            from .utility import xl_cell_to_rowcol

            first_col = xl_cell_to_rowcol(parts[0] + "1")[1]
            last_col = xl_cell_to_rowcol(parts[-1] + "1")[1]
            width, cell_format, options = actual_width, actual_format, actual_options
        first_col, last_col = sorted((int(first_col), int(last_col)))
        if first_col < 0 or last_col >= MAX_COLS:
            return -1
        options = options or {}
        record = (
            width,
            cell_format,
            bool(options.get("hidden")),
            min(7, max(0, int(options.get("level", 0)))),
            bool(options.get("collapsed")),
        )
        for column in range(first_col, last_col + 1):
            self.col_options[column] = record
        return 0

    def set_row(self, row, height=None, cell_format=None, options=None):
        row = int(row)
        if row < 0 or row >= MAX_ROWS:
            return -1
        options = options or {}
        hidden = bool(options.get("hidden"))
        if height == 0:
            hidden, height = True, 15
        self.row_options[row] = (
            15 if height is None else height,
            cell_format,
            hidden,
            min(7, max(0, int(options.get("level", 0)))),
            bool(options.get("collapsed")),
        )
        self._check(row, self.dim_colmin or 0)
        return 0

    def freeze_panes(self, row, col=None, top_row=None, left_col=None, pane_type=0):
        if isinstance(row, str):
            row, col = cell_args(row)
        top_row = row if top_row is None else top_row
        left_col = col if left_col is None else left_col
        self.panes = int(row), int(col), int(top_row), int(left_col)

    def autofilter(self, first_row, first_col=None, last_row=None, last_col=None):
        first_row, first_col, last_row, last_col = range_args(
            first_row, first_col, last_row, last_col
        )
        first_row, last_row = sorted((first_row, last_row))
        first_col, last_col = sorted((first_col, last_col))
        self.autofilter_ref = xl_range(first_row, first_col, last_row, last_col)

    def merge_range(
        self, first_row, first_col=None, last_row=None, last_col=None, data=None, cell_format=None
    ):
        if isinstance(first_row, str):
            cell_format = last_row if isinstance(last_row, Format) else cell_format
            data = first_col
            first_row, first_col, last_row, last_col = range_args(first_row)
        else:
            first_row, first_col, last_row, last_col = range_args(
                first_row, first_col, last_row, last_col
            )
        first_row, last_row = sorted((first_row, last_row))
        first_col, last_col = sorted((first_col, last_col))
        if first_row == last_row and first_col == last_col:
            return -1
        coordinates = {
            (row, col)
            for row in range(first_row, last_row + 1)
            for col in range(first_col, last_col + 1)
        }
        if coordinates & self._merged_cells:
            raise OverlappingRange("Merge range overlaps previous merge range")
        self._merged_cells |= coordinates
        self.merges.append((first_row, first_col, last_row, last_col))
        result = self.write(first_row, first_col, data, cell_format)
        for row, col in coordinates - {(first_row, first_col)}:
            self.write_blank(row, col, "", cell_format)
        return result

    def activate(self):
        for worksheet in self.workbook.worksheets_objs:
            worksheet.active = False
        self.active = True
        self.selected = True

    def select(self):
        self.selected = True

    def hide(self):
        self.hidden = True
        self.selected = False

    def get_name(self):
        return self.name

    def _external_url(self, url: str):
        if url.lower().startswith("external:"):
            path = url[9:].replace("\\", "/")
            return quote(path, safe="/:#[]!$&'()*+,;=@"), None
        if url.lower().startswith("internal:"):
            return None, url[9:]
        return url, None
