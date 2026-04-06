from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import polars as pl

from excelsior import Editor

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
    assert ws["C4"].value == "TRUE"
    assert ws["C5"].value == "FALSE"

    assert ws["D4"].is_date is True
    assert ws["D5"].is_date is True
    assert ws["D4"].number_format == "yyyy-mm-dd hh:mm:ss"

    assert ws["E4"].is_date is True
    assert ws["E5"].is_date is True
    assert ws["E4"].number_format == "yyyy-mm-dd"
