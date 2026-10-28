"""Python adapters over the Rust editor; no Polars or Arrow imports required."""

from datetime import date, datetime
from os import PathLike
import re
from typing import Any

from .excelsior import Editor as _Editor, Scanner as _Scanner, create_excel

_MAX_ROWS = 1_048_576
_MAX_COLUMNS = 16_384


def _column_name(index: int) -> str:
    result = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _excel_value(value: Any) -> tuple[str, str | None]:
    if value is None:
        return "", None
    if isinstance(value, bool):
        return ("TRUE" if value else "FALSE"), None
    if isinstance(value, (datetime, date)):
        # XLSX has no timezone field: retain the displayed local date and time.
        day = value.date() if isinstance(value, datetime) else value
        serial = (day - date(1899, 12, 30)).days
        if 0 < serial < 61:  # Excel's fictitious 1900-02-29.
            serial -= 1
        fmt = "yyyy-mm-dd"
        if isinstance(value, datetime):
            seconds = (
                value.hour * 3600 + value.minute * 60 + value.second
                + value.microsecond / 1_000_000
            )
            serial += seconds / 86_400
            fmt = "yyyy-mm-dd hh:mm:ss"
        return f"{serial:.15f}".rstrip("0").rstrip("."), fmt
    return str(value), None


class Editor(_Editor):
    @classmethod
    def create(cls, path: str | PathLike[str], sheet_name: str = "Sheet1") -> "Editor":
        """Create an empty XLSX at path and open it. Replaces an existing file."""
        create_excel(path, sheet_name)
        return cls(path, sheet_name)

    @classmethod
    def open(cls, path: str | PathLike[str], sheet_name: str | None = None) -> "Editor":
        """Open a worksheet, defaulting to the first sheet."""
        return cls(path, sheet_name)

    def with_polars(self, py_df: Any, start_cell: str | None = None) -> None:
        """Write headers and rows using only DataFrame.columns and iter_rows().

        Polars is neither imported nor converted through Arrow. Any installed
        version exposing these public APIs can supply the DataFrame.
        """
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
            return
        if start_column + len(columns) > _MAX_COLUMNS:
            raise ValueError("DataFrame exceeds XLSX column limit")
        table = [columns]
        formats: list[str | None] = [None] * len(columns)
        for row in py_df.iter_rows():
            if start_row + len(table) > _MAX_ROWS:
                raise ValueError("DataFrame exceeds XLSX row limit (including header)")
            if len(row) != len(columns):
                raise ValueError("DataFrame row width does not match its columns")
            values = []
            for index, value in enumerate(row):
                text, fmt = _excel_value(value)
                values.append(text)
                if fmt is not None and formats[index] != "yyyy-mm-dd hh:mm:ss":
                    formats[index] = fmt
            table.append(values)

        self.append_table_at(table, f"{letters.upper()}{start_row}")
        if len(table) > 1:
            for offset, fmt in enumerate(formats):
                if fmt is not None:
                    column = _column_name(start_column + offset)
                    end_row = start_row + len(table) - 1
                    self.set_number_format(f"{column}{start_row + 1}:{column}{end_row}", fmt)


class Scanner:
    def __init__(self, path: str | PathLike[str]) -> None:
        self._path = path
        self._scanner = _Scanner(path)

    def get_sheets(self) -> list[str]:
        return self._scanner.get_sheets()

    def open_editor(self, sheet_name: str | None = None) -> Editor:
        return Editor(self._path, sheet_name)
