"""Saving over an open source must also work on Windows (MoveFileEx)."""

from zipfile import ZipFile

import openpyxl
import pytest

from excelsior import Editor, Scanner, WorkbookWriter


@pytest.mark.parametrize("threads", [1, 3])
def test_repeated_in_place_save_retains_unloaded_sheets(tmp_path, threads):
    path = tmp_path / "source.xlsx"
    with WorkbookWriter(path) as writer:
        writer.append_row(["original"])
        writer.add_worksheet("Unloaded")
        writer.append_row(["00123"])
    original = Scanner(path).open_editor()
    branch = original.add_worksheet("Added")
    branch.append_row(["branch"])
    for _ in range(2):
        branch.save(path, compression_threads=threads)
        with ZipFile(path) as archive:
            assert archive.testzip() is None
        book = openpyxl.load_workbook(path, read_only=True)
        try:
            assert book.sheetnames == ["Sheet1", "Unloaded", "Added"]
            assert book["Unloaded"]["A1"].value == "00123"
            assert book["Added"]["A1"].value == "branch"
        finally:
            book.close()
    original.save(path, compression_threads=threads)
    assert Scanner(path).get_sheets() == ["Sheet1", "Unloaded"]


@pytest.mark.parametrize("threads", [1, 3])
def test_repeated_streaming_save_replaces_published_destination(tmp_path, threads):
    path = tmp_path / "streaming.xlsx"
    writer = WorkbookWriter(path, compression_threads=threads)
    writer.append_row(["00123", 1.25])
    writer.save()
    expected = path.read_bytes()
    for destination in [path, tmp_path / "copy.xlsx", path]:
        writer.save(destination)
        assert destination.read_bytes() == expected
    editor = Editor.open(path)
    editor.set_cell("B1", "2.5")
    editor.save()
    writer.save()
    assert path.read_bytes() == expected
