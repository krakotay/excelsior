"""Bounded-memory XLSX creation; all write methods mutate and return None."""
from os import PathLike
from typing import Any, Iterable, TypeVar
from .excelsior import WorkbookWriter as _Writer
from ._editor import _excel_value, _prepare_frame, _start_column

_WriterT = TypeVar("_WriterT", bound="WorkbookWriter")

class WorkbookWriter:
    """Create a new workbook without retaining worksheet XML in memory.

    Rows use a temporary disk spool; completed sheets are immediately written to
    ZIP. save/close finalizes the writer. The existing destination is preserved
    until save succeeds. Set compression at construction, before closing sheets.
    compression_threads closes independent sheets in parallel using disk files.
    """
    def __init__(self, path: str | PathLike[str], sheet_name: str = "Sheet1", *, compression_level: int = 3, compression_threads: int = 1) -> None:
        self._native = _Writer(path, sheet_name, compression_level=compression_level, compression_threads=compression_threads)

    @classmethod
    def create(cls: type[_WriterT], path: str | PathLike[str], sheet_name: str = "Sheet1", *, compression_level: int = 3, compression_threads: int = 1) -> _WriterT:
        return cls(path, sheet_name, compression_level=compression_level, compression_threads=compression_threads)

    def add_worksheet(self, sheet_name: str) -> None:
        """Close the current sheet and select a new one; do not assign the result."""
        self._native.add_worksheet(sheet_name)

    def with_polars(self, py_df: Any, start_cell: str | None = None, *, force_column_formats: bool = True) -> None:
        """Append headers and typed rows; default start is after existing rows.

        Explicit starts must follow previous rows. Existing sheets cannot be
        edited; use Editor for templates or overlapping writes.
        """
        start = f"A{self._native.next_row()}" if start_cell is None else start_cell
        frame = _prepare_frame(py_df, start, False)
        if frame is None:
            return
        start, formats, rows = frame
        self._native.write_typed_rows(start, rows)
        if force_column_formats:
            column = _start_column(start)
            self._native.force_column_formats([(column + i, fmt) for i, fmt in enumerate(formats) if fmt is not None])

    def write_rows(self, rows: Iterable[Iterable[Any]], start_cell: str | None = None) -> None:
        """Append rows from an iterator, preserving Python cell types."""
        start = f"A{self._native.next_row()}" if start_cell is None else start_cell
        formats = {}
        def typed_rows():
            for row in rows:
                values = []
                for column, value in enumerate(row):
                    cell, fmt = _excel_value(value, False)
                    values.append(cell)
                    if fmt is not None:
                        formats[column] = fmt
                yield values
        self._native.write_typed_rows(start, typed_rows())
        column = _start_column(start.upper())
        self._native.force_column_formats([(column + i, fmt) for i, fmt in formats.items()])

    def append_row(self, cells: Iterable[Any]) -> None:
        self.write_rows((cells,))

    def set_number_format(self, col_range: str, fmt: str) -> None:
        """Set entire columns of the current sheet, even after appending rows."""
        self._native.set_number_format(col_range, fmt)

    def set_column_width(self, col_range: str, width: float) -> None:
        self._native.set_column_width(col_range, width)

    def save(self, path: str | PathLike[str] | None = None) -> None:
        """Finalize; save-as preserves the remembered original destination."""
        self._native.save(path)

    def close(self) -> None:
        self.save()

    def abort(self) -> None:
        """Discard temporary files and preserve any existing destination."""
        self._native.abort()

    def __enter__(self: _WriterT) -> _WriterT:
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> bool:
        if exc_type is None:
            self.save()
        else:
            self.abort()
        return False

    def __copy__(self):
        raise TypeError("WorkbookWriter cannot be copied; Editor supports independent snapshots")

    def __deepcopy__(self, memo):
        return self.__copy__()
