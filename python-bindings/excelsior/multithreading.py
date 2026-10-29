"""Parallel file exports, scheduled by a bounded pool of native Rust threads.

Record a workbook's writes, then submit them on save/context exit. Shared pools
wait for every file at their outer context exit. Dataframes/iterables are consumed
in workers: keep them unchanged until completion. No Polars import is required.
"""

from __future__ import annotations

import atexit
import operator
from os import PathLike
from os.path import normcase
from pathlib import Path
from typing import Any, Iterable
from weakref import WeakSet

from .excelsior import _WorkbookPool, _ExportTask
from ._writer import WorkbookWriter as _SerialWriter

__all__ = ["WorkbookPool", "WorkbookWriter", "ExportTask"]

_pools: WeakSet[WorkbookPool] = WeakSet()


def _positive(value: int, name: str) -> int:
    value = operator.index(value)
    if value < 1:
        raise ValueError(f"{name} must be positive")
    return value


def _compression(level: int, threads: int) -> tuple[int, int]:
    level = operator.index(level)
    if not 0 <= level <= 9:
        raise ValueError("compression_level must be between 0 and 9")
    return level, _positive(threads, "compression_threads")


class ExportTask:
    """A submitted file. result() waits and re-raises the original export error."""

    def __init__(self, path: Path, native: _ExportTask) -> None:
        self.path = path
        self._native = native

    def done(self) -> bool:
        return self._native.done()

    def cancel(self) -> bool:
        """Cancel only if not started; running exports finish normally."""
        return self._native.cancel()

    def result(self) -> Path:
        self._native.wait()
        return self.path


class WorkbookPool:
    """Own Rust file workers; close waits for all jobs, then raises the first error.

    At most workers books execute and workers more wait in the native queue.
    Submission can block to bound the queue. Separate books can finish in any
    order. Each destination is allowed once per pool. Existing destinations are
    replaced atomically only after that file succeeds. A failure is per file;
    successful exports remain available. Use one outer context for many files.
    """

    def __init__(
        self,
        workers: int = 4,
        *,
        compression_level: int = 3,
        compression_threads: int = 1,
    ) -> None:
        self.compression_level, self.compression_threads = _compression(
            compression_level, compression_threads
        )
        self._native = _WorkbookPool(_positive(workers, "workers"))
        self._closed = False
        _pools.add(self)

    def workbook(
        self,
        path: str | PathLike[str],
        sheet_name: str = "Sheet1",
        *,
        compression_level: int | None = None,
        compression_threads: int | None = None,
    ) -> WorkbookWriter:
        """Record one file; its context submits without waiting for other books."""
        if self._closed:
            raise RuntimeError("WorkbookPool is closed")
        return WorkbookWriter(
            path,
            sheet_name,
            pool=self,
            compression_level=compression_level,
            compression_threads=compression_threads,
        )

    def write_polars(
        self,
        path: str | PathLike[str],
        df: Any,
        *,
        max_rows: int = 1_000_000,
        sheet_name: str = "Sheet1",
        force_column_formats: bool = True,
        compression_level: int | None = None,
        compression_threads: int | None = None,
    ) -> ExportTask:
        """Submit a dataframe split into sheets, each with headers.

        max_rows counts data rows (1..1_048_575); an empty frame still writes
        headers. Later sheet names are Sheet2, Sheet3, ... (skipping collisions).
        """
        max_rows = _positive(max_rows, "max_rows")
        if max_rows > 1_048_575:
            raise ValueError("max_rows must leave room for a header (<= 1048575)")
        height = operator.index(df.height)
        if height < 0:
            raise ValueError("dataframe height must be nonnegative")
        with self.workbook(
            path,
            sheet_name,
            compression_level=compression_level,
            compression_threads=compression_threads,
        ) as wb:
            number = 2
            for offset in range(0, max(1, height), max_rows):
                if offset:
                    while f"Sheet{number}".casefold() == sheet_name.strip().casefold():
                        number += 1
                    wb.add_worksheet(f"Sheet{number}")
                    number += 1
                wb.with_polars(
                    df.slice(offset, max_rows),
                    force_column_formats=force_column_formats,
                )
        assert wb.task is not None
        return wb.task

    def close(self, *, cancel_pending: bool = False) -> None:
        """Join all workers; queued jobs may be cancelled, running jobs finish."""
        try:
            self._native.close(cancel_pending=cancel_pending)
        finally:
            self._closed = self._native.is_closed()

    def __enter__(self) -> WorkbookPool:
        if self._closed:
            raise RuntimeError("WorkbookPool is closed")
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> bool:
        if exc_type is None:
            self.close()
        else:
            # Keep the caller's exception even if a background file also failed.
            try:
                self.close(cancel_pending=True)
            except BaseException:
                pass
        return False

    def __copy__(self):
        raise TypeError("WorkbookPool cannot be copied")

    def __deepcopy__(self, memo):
        return self.__copy__()


class WorkbookWriter:
    """Deferred workbook commands, executed by Rust workers after submission.

    Pass pool= to overlap separate files; without it this context owns one
    worker and waits on exit. Mutators return None. save/close submits once;
    wait waits for this file. To append more rows, record them before save.
    """

    def __init__(
        self,
        path: str | PathLike[str],
        sheet_name: str = "Sheet1",
        *,
        pool: WorkbookPool | None = None,
        compression_level: int | None = None,
        compression_threads: int | None = None,
    ) -> None:
        self._path = Path(path).resolve()
        self._sheet_name = sheet_name
        level = (
            (pool.compression_level if pool else 3)
            if compression_level is None
            else compression_level
        )
        threads = (
            (pool.compression_threads if pool else 1)
            if compression_threads is None
            else compression_threads
        )
        self._level, self._threads = _compression(level, threads)
        self._owns_pool = pool is None
        self._pool = pool if pool is not None else WorkbookPool(workers=1)
        if self._pool._closed:
            raise RuntimeError("WorkbookPool is closed")
        self._commands: list[tuple[str, tuple, dict]] = []
        self._sealed = False
        self.task: ExportTask | None = None

    def _record(self, name: str, *args, **kwargs) -> None:
        if self._sealed:
            raise RuntimeError("WorkbookWriter is already submitted or aborted")
        if self._pool._closed:
            raise RuntimeError("WorkbookPool is closed")
        self._commands.append((name, args, kwargs))

    def add_worksheet(self, sheet_name: str) -> None:
        self._record("add_worksheet", sheet_name)

    def with_polars(
        self,
        py_df: Any,
        start_cell: str | None = None,
        *,
        force_column_formats: bool = True,
    ) -> None:
        self._record(
            "with_polars", py_df, start_cell, force_column_formats=force_column_formats
        )

    def write_rows(
        self, rows: Iterable[Iterable[Any]], start_cell: str | None = None
    ) -> None:
        self._record("write_rows", rows, start_cell)

    def append_row(self, cells: Iterable[Any]) -> None:
        self.write_rows((cells,))

    def set_number_format(self, col_range: str, fmt: str) -> None:
        self._record("set_number_format", col_range, fmt)

    def set_column_width(self, col_range: str, width: float) -> None:
        self._record("set_column_width", col_range, width)

    def save(self, path: str | PathLike[str] | None = None) -> None:
        """Submit to the pool; an explicit path exports under a different name.

        A submitted writer cannot be replayed to another path: its iterators
        may have been consumed. Repeated save() is harmless.
        """
        if self._sealed:
            if self.task is None or path is not None:
                raise RuntimeError("WorkbookWriter is already submitted or aborted")
            return
        destination = self._path if path is None else Path(path).resolve()
        commands = self._commands
        sheet, level, threads = self._sheet_name, self._level, self._threads

        def run():
            writer = None
            try:
                with _SerialWriter(
                    destination,
                    sheet,
                    compression_level=level,
                    compression_threads=threads,
                ) as writer:
                    for name, args, kwargs in commands:
                        getattr(writer, name)(*args, **kwargs)
            except BaseException as error:
                if writer is not None:
                    writer.abort()
                if hasattr(error, "add_note"):
                    try:
                        error.add_note(f"excelsior export: {destination}")
                    except Exception:
                        pass
                raise
            finally:
                commands.clear()

        native = self._pool._native.submit(Path(normcase(destination)), run)
        self.task = ExportTask(destination, native)
        self._commands = []
        self._sealed = True
        if self._owns_pool:
            self._pool.close()

    def close(self) -> None:
        self.save()

    def wait(self) -> Path:
        if not self._sealed:
            self.save()
        if self.task is None:
            raise RuntimeError("WorkbookWriter was aborted")
        return self.task.result()

    def abort(self) -> None:
        """Discard recorded commands, or cancel a submitted job if still queued."""
        self._sealed = True
        self._commands.clear()
        if self.task is not None:
            self.task.cancel()
        if self._owns_pool:
            try:
                self._pool.close(cancel_pending=True)
            except BaseException:
                pass

    def __enter__(self) -> WorkbookWriter:
        if self._sealed:
            raise RuntimeError("WorkbookWriter is already submitted or aborted")
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> bool:
        if exc_type is None:
            self.save()
        else:
            self.abort()
        return False

    def __copy__(self):
        raise TypeError("WorkbookWriter cannot be copied")

    def __deepcopy__(self, memo):
        return self.__copy__()


def _shutdown_pools() -> None:
    # Native threads are invisible to threading's interpreter-shutdown join.
    # Drain explicit pools before Python finalization; normal use closes via with.
    for pool in list(_pools):
        try:
            pool.close()
        except BaseException:
            pass


atexit.register(_shutdown_pools)
