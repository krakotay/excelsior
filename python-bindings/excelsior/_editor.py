"""Python dataframe adaptation and explicit value/mutation semantics."""

from datetime import date, datetime, time, timedelta
from decimal import Decimal
from math import isfinite
from os import PathLike
import re
from typing import Any, TypeVar

from .excelsior import AlignSpec, Editor as _Editor, Scanner as _Scanner, create_excel

_EditorT = TypeVar("_EditorT", bound="Editor")

_MAX_ROWS = 1_048_576
_MAX_COLUMNS = 16_384
_FLOAT_FORMAT = "0.###############"


def _dtype_format(dtype: Any) -> str | None:
    name = str(dtype).split("(", 1)[0]
    if name.startswith(("Int", "UInt")):
        return "0"
    if name.startswith("Float"):
        return _FLOAT_FORMAT
    if name == "Decimal":
        scale = getattr(dtype, "scale", 0) or 0
        return "0" + ("." + "0" * scale if scale else "")
    if name in {"String", "Utf8", "Categorical", "Enum"}:
        return "@"
    if name == "Date":
        return "yyyy-mm-dd"
    if name == "Datetime":
        return "yyyy-mm-dd hh:mm:ss"
    if name == "Time":
        return "hh:mm:ss.000"
    if name == "Duration":
        return "[h]:mm:ss.000"
    return None


def _excel_value(value: Any, date_1904: bool) -> tuple[tuple[int, str], str | None]:
    # The small tagged-cell protocol is independent of Polars and Arrow.
    if value is None:
        return (0, ""), None
    if isinstance(value, bool):
        return (3, "1" if value else "0"), None
    if isinstance(value, (datetime, date)):
        # XLSX has no timezone field: retain the displayed local date and time.
        day = value.date() if isinstance(value, datetime) else value
        serial = (day - date(1904, 1, 1) if date_1904 else day - date(1899, 12, 30)).days
        if not date_1904 and 0 < serial < 61:
            serial -= 1  # Excel's fictitious 1900-02-29.
        fmt = "yyyy-mm-dd"
        if isinstance(value, datetime):
            serial += (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000) / 86_400
            fmt = "yyyy-mm-dd hh:mm:ss"
        return (2, str(serial)), fmt
    if isinstance(value, time):
        seconds = value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1_000_000
        return (2, str(seconds / 86_400)), "hh:mm:ss.000"
    if isinstance(value, timedelta):
        return (2, str(value.total_seconds() / 86_400)), "[h]:mm:ss.000"
    if isinstance(value, int):
        return (2, str(value)), "0"
    if isinstance(value, (float, Decimal)):
        if not isfinite(value):
            # Excel has no numeric NaN/infinity; retain their spelling as text.
            return (1, str(value)), None
        return (2, str(value)), _FLOAT_FORMAT
    return (1, str(value)), "@"


def _dtype_converter(dtype: Any, date_1904: bool):
    """Choose once per column; retain scalar conversion for unusual objects."""
    name = str(dtype).split("(", 1)[0]
    def generic(value):
        return _excel_value(value, date_1904)[0]
    if name in {"String", "Utf8", "Categorical", "Enum"}:
        def text(value):
            if value is None:
                return (0, "")
            return (1, value) if isinstance(value, str) else generic(value)
        return text
    if name.startswith("Float"):
        def number(value):
            if value is None:
                return (0, "")
            # Raw finite floats are formatted by Rust outside the GIL. Other
            # values keep the scalar adapter's spelling and conversion rules.
            return (5, value) if isinstance(value, float) and isfinite(value) else generic(value)
        return number
    if name.startswith(("Int", "UInt")):
        def integer(value):
            if value is None:
                return (0, "")
            return (2, str(value)) if type(value) is int else generic(value)
        return integer
    return generic


class Editor:
    """XLSX editor with two explicit method contracts.

    Methods returning Editor operate on an independent copy; assign the result.
    Methods returning None mutate this instance (save also writes to disk).
    """

    def __init__(self, path: str | PathLike[str], sheet_name: str | None = None) -> None:
        self._native = _Editor(path, sheet_name)

    @classmethod
    def create(cls: type[_EditorT], path: str | PathLike[str], sheet_name: str = "Sheet1") -> _EditorT:
        """Create an empty XLSX at path and open it. Replaces an existing file."""
        create_excel(path, sheet_name)
        return cls(path, sheet_name)

    @classmethod
    def open(cls: type[_EditorT], path: str | PathLike[str], sheet_name: str | None = None) -> _EditorT:
        return cls(path, sheet_name)

    def copy(self: _EditorT) -> _EditorT:
        """Return an independent snapshot, including all unsaved edits."""
        result = type(self).__new__(type(self))
        result.__dict__ = self.__dict__.copy()
        result._native = self._native.clone()
        return result

    def __copy__(self: _EditorT) -> _EditorT:
        return self.copy()

    def __deepcopy__(self: _EditorT, memo: dict) -> _EditorT:
        from copy import deepcopy
        result = type(self).__new__(type(self))
        memo[id(self)] = result
        result.__dict__ = {k: self._native.clone() if k == "_native" else deepcopy(v, memo) for k, v in self.__dict__.items()}
        return result

    def _updated(self: _EditorT, method: str, *args: Any) -> _EditorT:
        result = self.copy()
        getattr(result._native, method)(*args)
        return result

    def add_worksheet(self: _EditorT, sheet_name: str) -> _EditorT:
        return self._updated("add_worksheet", sheet_name)

    def add_worksheet_at(self: _EditorT, sheet_name: str, index: int) -> _EditorT:
        return self._updated("add_worksheet_at", sheet_name, index)

    def with_worksheet(self: _EditorT, sheet_name: str) -> _EditorT:
        return self._updated("with_worksheet", sheet_name)

    def rename_worksheet(self: _EditorT, old_name: str, new_name: str) -> _EditorT:
        return self._updated("rename_worksheet", old_name, new_name)

    def delete_worksheet(self: _EditorT, sheet_name: str) -> _EditorT:
        return self._updated("delete_worksheet", sheet_name)

    def set_number_format(self: _EditorT, range: str, fmt: str) -> _EditorT:
        return self._updated("set_number_format", range, fmt)

    def set_fill(self: _EditorT, range: str, rgb: str) -> _EditorT:
        return self._updated("set_fill", range, rgb)

    def merge_cells(self: _EditorT, range: str) -> _EditorT:
        return self._updated("merge_cells", range)

    def set_border(self: _EditorT, range: str, style: str) -> _EditorT:
        return self._updated("set_border", range, style)

    def set_font(self: _EditorT, range: str, name: str, size: float, bold: bool = False, italic: bool = False, align: AlignSpec | None = None) -> _EditorT:
        return self._updated("set_font", range, name, size, bold, italic, align)

    def set_alignment(self: _EditorT, range: str, spec: AlignSpec) -> _EditorT:
        return self._updated("set_alignment", range, spec)

    def set_column_width(self: _EditorT, col_letter: str, width: float) -> _EditorT:
        return self._updated("set_column_width", col_letter, width)

    def set_column_width_range(self: _EditorT, col_range: str, width: float) -> _EditorT:
        return self._updated("set_column_width_range", col_range, width)

    def set_columns_width(self: _EditorT, col_letters: list[str], width: float) -> _EditorT:
        return self._updated("set_columns_width", col_letters, width)

    def set_column_widths(self: _EditorT, widths: dict[str, float]) -> _EditorT:
        return self._updated("set_column_widths", widths)

    def remove_style(self: _EditorT, range: str) -> _EditorT:
        return self._updated("remove_style", range)

    def set_cell(self, coords: str, cell: str) -> None:
        self._native.set_cell(coords, cell)

    def append_row(self, cells: list[str]) -> None:
        self._native.append_row(cells)

    def append_table_at(self, cells: list[list[str]], start_cell: str) -> None:
        self._native.append_table_at(cells, start_cell)

    def last_row_index(self, col_name: str) -> int:
        return self._native.last_row_index(col_name)

    def last_rows_index(self, col_name: str) -> list[int]:
        return self._native.last_rows_index(col_name)

    def save(self, path: str | PathLike[str] | None = None, *, compression_level: int = 3, compression_threads: int = 1) -> None:
        """Save atomically; compression_threads compresses ZIP members in parallel."""
        self._native.save(path, compression_level=compression_level, compression_threads=compression_threads)

    def with_polars(self, py_df: Any, start_cell: str | None = None, *, force_column_formats: bool = True) -> None:
        """Mutate this worksheet with typed headers/rows from a dataframe.

        By default dtype number formats override the entire destination columns,
        including existing cells outside the written area. False preserves all
        template formats. String values remain text in either mode. NaN/inf are
        written as text. No Polars import or Arrow conversion is performed.
        """
        frame = _prepare_frame(py_df, start_cell, self._native.uses_1904_dates())
        if frame is None:
            return
        start, formats, rows = frame
        column = _start_column(start)
        if force_column_formats and getattr(py_df, "dtypes", None) is not None and self._native.is_empty_sheet():
            # Empty sheets need no cell rewrite: establish column defaults first,
            # then the native row writer assigns those styles as it writes cells.
            draft = self._native.clone()
            draft.force_column_formats([(column + i, fmt) for i, fmt in enumerate(formats) if fmt is not None])
            draft.write_typed_rows(start, rows)
            self._native = draft
        else:
            self._native.write_typed_rows(start, rows)
            if force_column_formats:
                self._native.force_column_formats([(column + i, fmt) for i, fmt in enumerate(formats) if fmt is not None])



def _start_column(start: str) -> int:
    column = 0
    for letter in re.match(r"[A-Z]+", start).group():
        column = column * 26 + ord(letter) - 64
    return column - 1


def _prepare_frame(py_df: Any, start_cell: str | None, date_1904: bool):
    start = "A1" if start_cell is None else start_cell
    match = re.fullmatch(r"([A-Za-z]+)([1-9][0-9]*)", start)
    if match is None:
        raise ValueError(f"Invalid Excel start cell: {start!r}")
    letters, row_text = match.groups()
    start_column = 0
    for letter in letters.upper():
        start_column = start_column * 26 + ord(letter) - 64
    start_column -= 1
    start_row = int(row_text)
    if start_column >= _MAX_COLUMNS or start_row > _MAX_ROWS:
        raise ValueError("Start cell exceeds XLSX worksheet limits")
    columns = list(py_df.columns)
    if not columns:
        return None
    if start_column + len(columns) > _MAX_COLUMNS:
        raise ValueError("DataFrame exceeds XLSX column limit")
    height = getattr(py_df, "height", None)
    if height is not None and start_row + height > _MAX_ROWS:
        raise ValueError("DataFrame exceeds XLSX row limit (including header)")
    dtypes = getattr(py_df, "dtypes", None)
    formats = [_dtype_format(dtype) for dtype in dtypes] if dtypes is not None else [None] * len(columns)
    if len(formats) != len(columns):
        raise ValueError("DataFrame dtype count does not match its columns")

    converters = [_dtype_converter(dtype, date_1904) for dtype in dtypes] if dtypes is not None else None

    def rows():
        yield [(1, str(column)) for column in columns]
        for offset, row in enumerate(py_df.iter_rows(), start=1):
            if start_row + offset > _MAX_ROWS:
                raise ValueError("DataFrame exceeds XLSX row limit (including header)")
            if len(row) != len(columns):
                raise ValueError("DataFrame row width does not match its columns")
            if converters is not None:
                yield [convert(value) for convert, value in zip(converters, row)]
            else:
                values = []
                for index, value in enumerate(row):
                    cell, fmt = _excel_value(value, date_1904)
                    values.append(cell)
                    if fmt is not None:
                        formats[index] = fmt
                yield values

    return f"{letters.upper()}{start_row}", formats, rows()


class Scanner:
    def __init__(self, path: str | PathLike[str]) -> None:
        self._path = path
        self._scanner = _Scanner(path)

    def get_sheets(self) -> list[str]:
        return self._scanner.get_sheets()

    def open_editor(self, sheet_name: str | None = None) -> Editor:
        return Editor(self._path, sheet_name)
