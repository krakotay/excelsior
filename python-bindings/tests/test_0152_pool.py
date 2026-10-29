"""File scheduling, error handling and output checks for the native Rust pool."""

import copy
import gc
import subprocess
import sys
from threading import Barrier, Event
from zipfile import ZipFile, ZIP_STORED

import openpyxl
import polars as pl
import pytest

import excelsior
from excelsior import multithreading as mt


def load(path):
    with ZipFile(path) as archive:
        assert archive.testzip() is None
    return openpyxl.load_workbook(path)


def test_native_workers_overlap_and_do_not_construct_python_threads(
    tmp_path, monkeypatch
):
    import threading
    import concurrent.futures

    def forbidden(*args, **kwargs):
        raise AssertionError("Python executor/thread must not be used")

    monkeypatch.setattr(threading, "Thread", forbidden)
    monkeypatch.setattr(concurrent.futures, "ThreadPoolExecutor", forbidden)
    barrier = Barrier(2, timeout=15)

    def rows(i):
        barrier.wait()  # Deadlocks/times out if file jobs execute sequentially.
        yield [f"{i:020d}", i + 0.25]

    with mt.WorkbookPool(workers=2, compression_threads=2) as pool:
        for i in range(2):
            with mt.WorkbookWriter(tmp_path / f"{i}.xlsx", pool=pool) as wb:
                assert wb.write_rows(rows(i)) is None
            assert wb.task is not None
        assert wb.task.result() == (tmp_path / "1.xlsx").resolve()
    for i in range(2):
        ws = load(tmp_path / f"{i}.xlsx").active
        assert ws["A1"].value == f"{i:020d}"
        assert ws["A1"].data_type == "s"
        assert ws["B1"].value == i + 0.25
    assert not list(tmp_path.glob(".excelsior-*"))


@pytest.mark.parametrize("rows", [0, 1, 7])
@pytest.mark.parametrize("sheet_name", ["Sheet2", " Sheet2 "])
def test_dataframe_convenience_splits_and_formats(tmp_path, rows, sheet_name):
    df = pl.DataFrame(
        {
            "id": [f"{i:020d}" for i in range(rows)],
            "amount": [i / 8 for i in range(rows)],
        },
        schema={"id": pl.String, "amount": pl.Float64},
    )
    path = tmp_path / "frame.xlsx"
    with mt.WorkbookPool(2, compression_level=0, compression_threads=3) as pool:
        task = pool.write_polars(path, df, max_rows=3, sheet_name=sheet_name)
    assert task.done()
    assert task.result() == path
    book = load(path)
    assert book.sheetnames == ["Sheet2", "Sheet3", "Sheet4"][: max(1, (rows + 2) // 3)]
    for index, sheet in enumerate(book):
        expected = list(df.slice(index * 3, 3).iter_rows())
        assert list(sheet.values) == [("id", "amount"), *expected]
        assert sheet.column_dimensions["A"].number_format == "@"
        assert sheet.column_dimensions["B"].number_format != "General"
    with ZipFile(path) as archive:
        assert {member.compress_type for member in archive.infolist()} == {ZIP_STORED}


def test_standalone_context_waits_and_mutators_are_none(tmp_path):
    path = tmp_path / "single.xlsx"
    with excelsior.multithreading.WorkbookWriter(path, compression_threads=8) as wb:
        assert wb.append_row(["header"]) is None
        assert wb.write_rows([[1], [2]], "A3") is None
        assert wb.with_polars(pl.DataFrame({"a": [3]})) is None
        assert wb.set_number_format("A", "0.000") is None
        assert wb.set_column_width("A", 23) is None
        assert wb.add_worksheet("last") is None
        wb.append_row(["end"])
    assert wb.task.done()
    assert wb.wait() == path
    book = load(path)
    assert [book["Sheet1"][f"A{i}"].value for i in range(1, 7)] == [
        "header",
        None,
        1,
        2,
        "a",
        3,
    ]
    assert book["Sheet1"].column_dimensions["A"].number_format == "0.000"
    assert book["Sheet1"].column_dimensions["A"].width == 23
    assert book["last"]["A1"].value == "end"
    assert wb.save() is None
    with pytest.raises(RuntimeError, match="submitted"):
        wb.append_row([4])


def test_force_formats_can_be_disabled(tmp_path):
    with mt.WorkbookWriter(tmp_path / "format.xlsx") as wb:
        wb.with_polars(
            pl.DataFrame({"id": ["111111119999999"]}), force_column_formats=False
        )
    ws = load(tmp_path / "format.xlsx").active
    assert ws["A2"].value == "111111119999999"
    assert ws["A2"].data_type == "s"
    assert ws.column_dimensions["A"].number_format == "General"


def test_original_iterator_error_waits_for_other_jobs_and_preserves_target(tmp_path):
    path = tmp_path / "bad.xlsx"
    path.write_bytes(b"original")

    def broken():
        for i in range(600):
            yield [i]
        raise LookupError("original iterator exception")

    pool = mt.WorkbookPool(2)
    with pytest.raises(LookupError, match="original iterator exception") as error:
        with pool:
            with pool.workbook(path) as wb:
                wb.write_rows(broken())
            with pool.workbook(tmp_path / "good.xlsx") as other:
                other.append_row(["okay"])
    assert wb.task.done() and other.task.done()
    assert path.read_bytes() == b"original"
    assert load(tmp_path / "good.xlsx").active["A1"].value == "okay"
    assert any(
        "broken" == frame.name
        for frame in __import__("traceback").extract_tb(error.value.__traceback__)
    )
    if sys.version_info >= (3, 11):
        assert str(path) in "\n".join(error.value.__notes__)
    with pytest.raises(LookupError, match="original iterator"):
        wb.wait()
    assert not list(tmp_path.glob(".excelsior-*"))


def test_cancel_queued_job_and_join_running_job(tmp_path):
    started, release = Event(), Event()

    def blocked():
        started.set()
        assert release.wait(15)
        yield ["completed"]

    pool = mt.WorkbookPool(1)
    with pool.workbook(tmp_path / "running.xlsx") as running:
        running.write_rows(blocked())
    assert started.wait(15)
    with pool.workbook(tmp_path / "cancelled.xlsx") as queued:
        queued.append_row(["must not write"])
    assert not running.task.cancel()
    assert queued.task.cancel()
    release.set()
    with pytest.raises(RuntimeError, match="cancelled"):
        pool.close()
    assert running.task.result().exists()
    assert queued.task.done()
    assert not (tmp_path / "cancelled.xlsx").exists()


def test_outer_and_inner_body_errors_preserved(tmp_path):
    with pytest.raises(ZeroDivisionError, match="body"):
        with mt.WorkbookPool(1) as pool:
            with pool.workbook(tmp_path / "discard.xlsx") as wb:
                wb.append_row([1])
                raise ZeroDivisionError("body")
    assert not (tmp_path / "discard.xlsx").exists()
    with pytest.raises(RuntimeError, match="aborted"):
        wb.wait()
    with pytest.raises(ZeroDivisionError, match="outer"):
        with mt.WorkbookPool(2) as pool:
            with pool.workbook(tmp_path / "bad.xlsx") as wb:
                wb.set_column_width("A", -1)
            raise ZeroDivisionError("outer")
    assert wb.task.done()
    assert not list(tmp_path.glob(".excelsior-*"))


def test_reentrant_close_and_wait_raise_without_deadlock_or_closing_pool(tmp_path):
    ready = Event()

    def rows():
        assert ready.wait(15)
        with pytest.raises(RuntimeError, match="own worker"):
            pool.close()
        assert not pool._closed
        with pytest.raises(RuntimeError, match="own worker"):
            wb.task.result()
        with pytest.raises(RuntimeError, match="own worker"):
            with pool.workbook(tmp_path / "nested.xlsx"):
                pass
        yield ["okay"]

    with mt.WorkbookPool(1) as pool:
        with pool.workbook(tmp_path / "outer.xlsx") as wb:
            wb.write_rows(rows())
        ready.set()
        wb.task.result()
    assert load(tmp_path / "outer.xlsx").active["A1"].value == "okay"


def test_duplicate_aliases_and_closed_pool(tmp_path):
    with mt.WorkbookPool(2) as pool:
        with pool.workbook(tmp_path / "same.xlsx") as wb:
            wb.append_row([1])
        with pytest.raises(ValueError, match="destination"):
            with pool.workbook(tmp_path / "child/../same.xlsx"):
                pass
    with pytest.raises(RuntimeError, match="closed"):
        pool.workbook(tmp_path / "new.xlsx")
    pool.close()
    assert load(tmp_path / "same.xlsx").active["A1"].value == 1


def test_save_as_and_abort(tmp_path):
    original = tmp_path / "original.xlsx"
    original.write_bytes(b"preserve")
    with mt.WorkbookWriter(original) as wb:
        wb.append_row(["copy"])
        wb.save(tmp_path / "copy.xlsx")
    assert original.read_bytes() == b"preserve"
    assert load(tmp_path / "copy.xlsx").active["A1"].value == "copy"
    with pytest.raises(RuntimeError, match="submitted"):
        wb.save(tmp_path / "another.xlsx")
    wb = mt.WorkbookWriter(original)
    wb.append_row(["discard"])
    assert wb.abort() is None
    assert original.read_bytes() == b"preserve"


@pytest.mark.parametrize(
    "kwargs",
    [
        {"workers": 0},
        {"workers": -1},
        {"compression_level": -1},
        {"compression_level": 10},
        {"compression_threads": 0},
    ],
)
def test_pool_invalid_options(kwargs):
    with pytest.raises(ValueError):
        mt.WorkbookPool(**kwargs)


@pytest.mark.parametrize("rows", [0, -1, 1048576])
def test_invalid_chunk_size(tmp_path, rows):
    with mt.WorkbookPool(1) as pool:
        with pytest.raises(ValueError):
            pool.write_polars(
                tmp_path / "bad.xlsx", pl.DataFrame({"a": [1]}), max_rows=rows
            )
    assert not (tmp_path / "bad.xlsx").exists()


def test_no_copy_and_abandoned_pools_join_without_gil_deadlock(tmp_path):
    pool = mt.WorkbookPool(1)
    with pool.workbook(tmp_path / "file.xlsx") as wb:
        wb.append_row([1])
    for obj in (pool, wb):
        with pytest.raises(TypeError):
            copy.copy(obj)
        with pytest.raises(TypeError):
            copy.deepcopy(obj)
    task = wb.task
    del wb, obj, pool
    gc.collect()
    # Abandoning a pool cancels queued jobs, but a job that already started may finish.
    try:
        task.result()
    except RuntimeError as error:
        assert "cancelled" in str(error)


def test_no_polars_dependency_and_interpreter_shutdown(tmp_path):
    script = """
import importlib.abc, sys
class Block(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname, path=None, target=None):
        if fullname.split('.')[0] in {'polars', 'pyarrow', 'pandas'}:
            raise AssertionError(fullname)
sys.meta_path.insert(0, Block())
from excelsior import multithreading as mt
pool = mt.WorkbookPool(2)
for i in range(4):
    with pool.workbook(sys.argv[1] + '/' + str(i) + '.xlsx') as wb:
        wb.write_rows(([str(j), j / 8] for j in range(500)))
# Deliberately leave pool open: shutdown must drain native threads before finalization.
"""
    subprocess.run(
        [sys.executable, "-c", script, str(tmp_path)], check=True, timeout=30
    )
    for i in range(4):
        assert load(tmp_path / f"{i}.xlsx").active["B500"].value == 499 / 8


def test_failed_export_cycles_release_dataframe_and_pool(tmp_path):
    import weakref

    class Payload:
        pass

    def failed_export():
        payload = Payload()
        pool = mt.WorkbookPool(1)
        wb = pool.workbook(tmp_path / "cycle.xlsx")

        def broken():
            # A realistic iterator closure retains its caller and large inputs.
            keep_alive = (payload, pool, wb)
            yield [1]
            raise LookupError(str(len(keep_alive)))

        wb.write_rows(broken())
        wb.save()
        try:
            pool.close()
        except LookupError:
            pass
        return weakref.ref(payload), weakref.ref(pool), weakref.ref(wb)

    refs = failed_export()
    gc.collect()
    assert all(ref() is None for ref in refs)


def test_more_jobs_than_queue_capacity_make_progress(tmp_path):
    barrier = Barrier(2, timeout=15)
    native_ids = set()

    def rows(index):
        import threading

        native_ids.add(threading.get_native_id())
        barrier.wait()
        yield [index]

    with mt.WorkbookPool(2) as pool:
        for i in range(20):
            with pool.workbook(tmp_path / f"queued-{i}.xlsx") as wb:
                wb.write_rows(rows(i))
    assert len(native_ids) == 2
    for i in range(20):
        assert load(tmp_path / f"queued-{i}.xlsx").active["A1"].value == i


def test_save_io_failure_removes_worker_files(tmp_path):
    destination = tmp_path / "directory.xlsx"
    destination.mkdir()
    with pytest.raises(RuntimeError):
        with mt.WorkbookWriter(destination, compression_threads=3) as wb:
            for i in range(3):
                wb.append_row([i])
                wb.add_worksheet(f"Other{i}")
    assert destination.is_dir()
    assert not list(tmp_path.glob(".excelsior-*"))
