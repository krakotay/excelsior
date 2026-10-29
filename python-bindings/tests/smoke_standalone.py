"""Run directly in a clean environment, before installing test dependencies."""

import importlib.abc
from pathlib import Path
import sys
import tempfile
import zipfile


class BlockOptionalDependencies(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname, path=None, target=None):
        if fullname.split(".")[0] in {
            "polars",
            "pyarrow",
            "numpy",
            "openpyxl",
            "pytest",
        }:
            raise ImportError(f"Optional dependency must not be imported: {fullname}")


sys.meta_path.insert(0, BlockOptionalDependencies())

from excelsior import Editor, Scanner, WorkbookWriter, __version__, create_excel


class Frame:
    columns = ["name", "count"]

    def iter_rows(self):
        yield "standalone", 42

    def to_arrow(self, *args, **kwargs):
        raise AssertionError("with_polars must not invoke Arrow")


with tempfile.TemporaryDirectory() as directory:
    path = Path(directory) / "empty.xlsx"
    create_excel(path)
    editor = Scanner(path).open_editor()
    editor = editor.add_worksheet("Data")
    editor.with_polars(Frame())
    editor.save(path)
    assert Scanner(path).get_sheets() == ["Sheet1", "Data"]
    assert isinstance(Editor.open(path), Editor)
    with zipfile.ZipFile(path) as archive:
        assert archive.testzip() is None
        assert b"standalone" in archive.read("xl/worksheets/sheet2.xml")
    streaming = Path(directory) / "streaming.xlsx"
    with WorkbookWriter(streaming, compression_level=3) as writer:
        writer.with_polars(Frame())
        writer.add_worksheet("Second")
        writer.write_rows((["00123", n] for n in range(3)))
    assert Scanner(streaming).get_sheets() == ["Sheet1", "Second"]
    with zipfile.ZipFile(streaming) as archive:
        assert archive.testzip() is None
        assert b"00123" in archive.read("xl/worksheets/sheet2.xml")
print(f"Standalone package {__version__}: OK")
