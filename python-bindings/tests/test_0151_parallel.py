from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from zipfile import ZipFile, ZIP_DEFLATED, ZIP_STORED

import openpyxl
import polars as pl
import pytest

from excelsior import Editor, WorkbookWriter


def dataframe(rows=700):
    return pl.DataFrame(
        {
            "id": [f"{i:020d}" for i in range(rows)],
            "amount": [i / 8 for i in range(rows)],
            "flag": [i % 2 == 0 for i in range(rows)],
            "date": [date(2025, 1, 2)] * rows,
        }
    )


@pytest.mark.parametrize("level", [0, 1, 3, 9])
def test_parallel_editor_compresses_all_parts_and_preserves_source(tmp_path, level):
    source, output = tmp_path / "source.xlsx", tmp_path / "output.xlsx"
    editor = Editor.create(source)
    for index in range(3):
        if index:
            editor = editor.add_worksheet(f"Sheet{index + 1}")
        editor.with_polars(dataframe(20))
    editor.save()
    payload = b"unchanged binary\0" * 10_000
    with ZipFile(source, "a") as archive:
        archive.writestr("xl/media/payload.bin", payload)
    original = source.read_bytes()
    editor = Editor.open(source)
    editor.set_cell("E3", "modified")
    editor.save(output, compression_level=level, compression_threads=3)
    assert source.read_bytes() == original
    with ZipFile(output) as archive:
        assert archive.testzip() is None
        assert archive.read("xl/media/payload.bin") == payload
        assert {info.compress_type for info in archive.infolist()} == {
            ZIP_STORED if level == 0 else ZIP_DEFLATED
        }
    book = openpyxl.load_workbook(output)
    assert book.sheetnames == ["Sheet1", "Sheet2", "Sheet3"]
    assert book["Sheet1"]["E3"].value == "modified"
    assert book["Sheet3"]["A2"].value == "00000000000000000000"
    assert book["Sheet3"]["B3"].value == 0.125
    assert not list(tmp_path.glob(".excelsior-*"))


@pytest.mark.parametrize("level", [0, 1, 3, 9])
def test_parallel_writer_orders_sheets_and_preserves_column_styles(tmp_path, level):
    path = tmp_path / "parallel-stream.xlsx"
    writer = WorkbookWriter(path, compression_level=level, compression_threads=3)
    for index in range(5):
        if index:
            writer.add_worksheet(f"Sheet{index + 1}")
        writer.with_polars(dataframe())
        writer.set_number_format("B", "0.00" if index % 2 else "0.0000")
    writer.save()
    book = openpyxl.load_workbook(path)
    assert book.sheetnames == [f"Sheet{i}" for i in range(1, 6)]
    for index, ws in enumerate(book):
        assert ws.max_row == 701
        assert ws["A701"].value == "00000000000000000699"
        assert ws["B701"].value == 699 / 8
        assert ws["B701"].number_format == ("0.00" if index % 2 else "0.0000")
        assert ws["C701"].value is False
        assert ws["D701"].value == datetime(2025, 1, 2)
    with ZipFile(path) as archive:
        assert archive.testzip() is None
        assert {info.compress_type for info in archive.infolist()} == {
            ZIP_STORED if level == 0 else ZIP_DEFLATED
        }
    writer.save(tmp_path / "copy.xlsx")
    assert path.read_bytes() == (tmp_path / "copy.xlsx").read_bytes()
    assert not list(tmp_path.glob(".excelsior-*"))


def test_separate_workbooks_accept_shared_dataframe_from_threads(tmp_path):
    frame = dataframe(1200)

    def export(index):
        path = tmp_path / f"{index}.xlsx"
        if index % 2:
            writer = WorkbookWriter(path, compression_level=1, compression_threads=2)
            writer.with_polars(frame)
            writer.add_worksheet("Other")
            writer.with_polars(frame.slice(300))
            writer.save()
        else:
            editor = Editor.create(path)
            editor.with_polars(frame)
            editor = editor.add_worksheet("Other")
            editor.with_polars(frame.slice(300))
            editor.save(compression_threads=2)
        return path

    with ThreadPoolExecutor(max_workers=4) as pool:
        outputs = list(pool.map(export, range(4)))
    for path in outputs:
        with ZipFile(path) as archive:
            assert archive.testzip() is None
        book = openpyxl.load_workbook(path)
        assert book["Other"]["A2"].value == "00000000000000000300"
        assert book["Other"]["B901"].value == 1199 / 8
    assert not list(tmp_path.glob(".excelsior-*"))


@pytest.mark.parametrize("threads", [0, -1])
def test_invalid_thread_count_preserves_destination(tmp_path, threads):
    target = tmp_path / "precious.xlsx"
    target.write_bytes(b"preserve")
    editor = Editor.create(tmp_path / "source.xlsx")
    with pytest.raises(ValueError, match="compression_threads"):
        editor.save(target, compression_threads=threads)
    with pytest.raises(ValueError, match="compression_threads"):
        WorkbookWriter(target, compression_threads=threads)
    assert target.read_bytes() == b"preserve"
    assert not list(tmp_path.glob(".excelsior-*"))


def test_aborting_parallel_writer_joins_and_removes_worker_files(tmp_path):
    target = tmp_path / "precious.xlsx"
    target.write_bytes(b"preserve")
    writer = WorkbookWriter(target, compression_threads=3)
    for index in range(4):
        writer.with_polars(dataframe())
        writer.add_worksheet(f"Next{index}")
    writer.abort()
    assert target.read_bytes() == b"preserve"
    assert not list(tmp_path.glob(".excelsior-*"))


@pytest.mark.parametrize("mode", ["editor", "stream"])
def test_error_past_native_batch_boundary_rolls_back_entire_call(tmp_path, mode):
    class BrokenFrame:
        columns = ["id", "amount"]
        dtypes = ["String", "Float64"]

        def iter_rows(self):
            for i in range(600):
                yield (f"id-{i}", i / 8)
            raise LookupError("original iterator failure")

    path = tmp_path / f"{mode}.xlsx"
    writer = (
        Editor.create(path)
        if mode == "editor"
        else WorkbookWriter(path, compression_threads=2)
    )
    writer.append_row(["before"])
    with pytest.raises(LookupError, match="original iterator failure"):
        writer.with_polars(BrokenFrame(), "A2")
    writer.append_row(["after"])
    writer.save()
    ws = openpyxl.load_workbook(path).active
    assert ws.max_row == 2
    assert ws["A1"].value == "before" and ws["A2"].value == "after"


@pytest.mark.parametrize("mode", ["editor", "stream"])
def test_prefetch_does_not_override_an_earlier_invalid_cell_error(tmp_path, mode):
    class BrokenFrame:
        columns = ["id"]
        dtypes = ["String"]

        def iter_rows(self):
            yield ("invalid\0XML",)
            raise LookupError("later error must not win")

    path = tmp_path / f"{mode}.xlsx"
    writer = Editor.create(path) if mode == "editor" else WorkbookWriter(path)
    with pytest.raises(RuntimeError, match="invalid XML character"):
        writer.with_polars(BrokenFrame())
    writer.append_row(["after"])
    writer.save()
    assert openpyxl.load_workbook(path).active["A1"].value == "after"


def test_fast_float_adapter_keeps_nonfinite_text_and_numeric_extremes(tmp_path):
    values = [
        0.0,
        -0.0,
        1.0,
        1e-300,
        1e300,
        float("nan"),
        float("inf"),
        -float("inf"),
        None,
    ]
    path = tmp_path / "floats.xlsx"
    editor = Editor.create(path)
    editor.with_polars(pl.DataFrame({"amount": values}))
    editor.save(compression_threads=2)
    ws = openpyxl.load_workbook(path).active
    for row, value in enumerate(values, 2):
        cell = ws.cell(row, 1)
        if row in [7, 8, 9]:
            assert cell.value == str(value) and cell.data_type == "s"
        else:
            assert cell.value == value
