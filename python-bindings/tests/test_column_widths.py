from pathlib import Path

from excelsior import Editor, create_excel

from helpers_excel import assert_with_libreoffice, column_width_map


def test_column_width_api_supports_single_range_and_mapping(tmp_path: Path) -> None:
    src = tmp_path / "widths.xlsx"
    out = tmp_path / "widths_out.xlsx"

    editor = Editor.create(str(src), "Widths")
    editor.set_column_width("A", 18.5)
    editor.set_column_width_range("B:D", 22.0)
    editor.set_columns_width(["F", "H:I"], 12.25)
    editor.set_column_widths({"K": 14.0, "M:N": 33.0})
    editor.save(str(out))

    widths = column_width_map(out)

    assert abs(widths["A"] - 18.5) < 0.001
    assert abs(widths["B"] - 22.0) < 0.001
    assert abs(widths["C"] - 22.0) < 0.001
    assert abs(widths["D"] - 22.0) < 0.001
    assert abs(widths["F"] - 12.25) < 0.001
    assert abs(widths["H"] - 12.25) < 0.001
    assert abs(widths["I"] - 12.25) < 0.001
    assert abs(widths["K"] - 14.0) < 0.001
    assert abs(widths["M"] - 33.0) < 0.001
    assert abs(widths["N"] - 33.0) < 0.001

    assert_with_libreoffice(out)


def test_column_width_validation_rejects_invalid_input(tmp_path: Path) -> None:
    out = tmp_path / "widths_invalid.xlsx"

    editor = Editor.create(str(out), "W")

    try:
        editor.set_column_width("", 10.0)
        assert False, "Expected invalid column error"
    except RuntimeError:
        pass

    try:
        editor.set_column_width("A", 0.0)
        assert False, "Expected invalid width error"
    except RuntimeError:
        pass

    try:
        editor.set_column_width_range("D:B", 10.0)
        assert False, "Expected invalid range error"
    except RuntimeError:
        pass
