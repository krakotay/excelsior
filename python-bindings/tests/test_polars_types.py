from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import polars as pl
import pytest

from excelsior import Editor, Scanner

from helpers_excel import get_sheet


def test_with_polars_formats_bool_and_datetime_columns(tmp_path: Path) -> None:
    src = tmp_path / "polars_types.xlsx"
    out = tmp_path / "polars_types_out.xlsx"

    df = pl.DataFrame(
        {
            "name": ["alpha", "beta"],
            "flag": [True, False],
            "created_at": [datetime(2025, 1, 2, 3, 4, 5), datetime(2025, 1, 3, 6, 7, 8)],
            "birthday": [date(2025, 1, 2), date(2025, 1, 3)],
        }
    )

    editor = Editor.create(str(src), "Data")
    editor.with_polars(df, "B3")
    editor.save(str(out))

    ws = get_sheet(out, "Data")

    assert ws["B3"].value == "name"
    assert ws["C4"].value is True
    assert ws["C5"].value is False

    assert ws["D4"].is_date is True
    assert ws["D5"].is_date is True
    assert ws["D4"].number_format == "yyyy-mm-dd hh:mm:ss"

    assert ws["E4"].is_date is True
    assert ws["E5"].is_date is True
    assert ws["E4"].number_format == "yyyy-mm-dd"


def test_sliced_frame_and_chained_editor_do_not_use_arrow(tmp_path, monkeypatch):
    def reject_arrow(*args, **kwargs):
        raise TypeError("compat_level has invalid type: int")

    monkeypatch.setattr(pl.DataFrame, "to_arrow", reject_arrow)
    frame = pl.DataFrame({"name": ["zero", "one", "two"], "value": [0, 1, 2]})
    path = tmp_path / "sliced.xlsx"
    editor = Editor.create(path)
    editor = editor.add_worksheet("Data")
    editor.with_polars(frame.slice(1, 2), "b3")
    editor.save(path)
    ws = get_sheet(path, "Data")
    assert ws["B3"].value == "name"
    assert ws["B4"].value == "one"
    assert ws["C5"].value == 2


@pytest.mark.parametrize("factory", [Editor, Editor.open, lambda path: Scanner(path).open_editor()])
def test_open_factories_return_python_adapter(tmp_path, factory):
    path = tmp_path / "factory.xlsx"
    Editor.create(path)
    editor = factory(path)
    assert isinstance(editor, Editor)
    editor.with_polars(py_df=pl.DataFrame({"x": [7]}))
    editor.save(path)
    assert get_sheet(path, "Sheet1")["A2"].value == 7


@pytest.mark.parametrize("frame, expected", [
    (pl.DataFrame(), None),
    (pl.DataFrame(schema={"name": pl.String}), "name"),
    (pl.DataFrame({"name": [None]}, schema={"name": pl.String}), "name"),
])
def test_empty_frames_and_nulls(tmp_path, frame, expected):
    path = tmp_path / "empty.xlsx"
    editor = Editor.create(path)
    editor.with_polars(frame)
    editor.save(path)
    ws = get_sheet(path, "Sheet1")
    assert ws["A1"].value == expected
    assert ws["A2"].value in (None, "")


@pytest.mark.parametrize("start", ["", "A0", "1", "A1x", "A-1", "XFE1", "A1048577"])
def test_invalid_start_does_not_mutate_workbook(tmp_path, start):
    path = tmp_path / "invalid.xlsx"
    editor = Editor.create(path)
    editor.set_cell("A1", "unchanged")
    with pytest.raises(ValueError):
        editor.with_polars(pl.DataFrame({"x": [1]}), start)
    editor.save(path)
    assert get_sheet(path, "Sheet1")["A1"].value == "unchanged"


@pytest.mark.parametrize("start", ["XFD1", "A1048576"])
def test_dataframe_limits_include_headers_and_width(tmp_path, start):
    path = tmp_path / "limit.xlsx"
    editor = Editor.create(path)
    with pytest.raises(ValueError):
        editor.with_polars(pl.DataFrame({"a": [1], "b": [2]}), start)
    editor.save(path)
    assert get_sheet(path, "Sheet1")["A1"].value is None


def test_python_frame_protocol_and_date_values(tmp_path):
    class Frame:
        columns = ["date", "datetime", "bool", "null"]

        def iter_rows(self):
            yield date(1900, 1, 1), datetime(2025, 1, 2, 3, 4, 5), True, None

    path = tmp_path / "protocol.xlsx"
    editor = Editor.create(path)
    editor.with_polars(Frame())
    editor.save(path)
    ws = get_sheet(path, "Sheet1")
    assert ws["A2"].value == datetime(1900, 1, 1)
    assert ws["B2"].value == datetime(2025, 1, 2, 3, 4, 5)
    assert ws["C2"].value is True
    assert ws["D2"].value in (None, "")
