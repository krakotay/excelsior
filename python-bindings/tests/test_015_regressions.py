from copy import copy, deepcopy
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from zipfile import ZipFile

import openpyxl
from openpyxl.utils.datetime import CALENDAR_MAC_1904
import polars as pl
import pytest

from excelsior import Editor, Scanner
from helpers_excel import get_sheet


def xml_parts(path):
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


@pytest.mark.parametrize("method,args", [
    ("add_worksheet", ("New",)), ("add_worksheet_at", ("New", 0)),
    ("with_worksheet", ("Other",)), ("rename_worksheet", ("Other", "Renamed")),
    ("delete_worksheet", ("Other",)), ("set_number_format", ("A1", "0.00")),
    ("set_fill", ("A1", "FFFF00")), ("merge_cells", ("B1:C1",)),
    ("set_border", ("A1", "thin")), ("set_font", ("A1", "Arial", 14.0)),
    ("set_column_width", ("A", 22.0)), ("set_column_width_range", ("A:C", 22.0)),
    ("set_columns_width", (["A", "C:D"], 22.0)), ("set_column_widths", ({"A": 22.0},)),
    ("remove_style", ("A1",)),
])
def test_editor_returning_methods_do_not_mutate_input(tmp_path, method, args):
    path = tmp_path / "source.xlsx"
    editor = Editor.create(path).add_worksheet("Other").with_worksheet("Sheet1")
    editor.append_row(["original", "2", "3"])
    editor = editor.set_fill("A1", "FF0000")
    editor.save(tmp_path / "before.xlsx")
    branch = getattr(editor, method)(*args)
    assert isinstance(branch, Editor) and branch is not editor
    branch.set_cell("Z5", "branch")
    branch.save(tmp_path / "branch.xlsx")
    editor.save(tmp_path / "after.xlsx")
    assert xml_parts(tmp_path / "before.xlsx") == xml_parts(tmp_path / "after.xlsx")
    assert xml_parts(tmp_path / "branch.xlsx") != xml_parts(tmp_path / "after.xlsx")


def test_alignment_copy_and_subclass_survive_factories(tmp_path):
    from excelsior import AlignSpec, HorizAlignment
    class CustomEditor(Editor):
        pass
    editor = CustomEditor.create(tmp_path / "source.xlsx")
    editor.set_cell("A1", "original")
    editor.extra = [1]
    branch = editor.set_alignment("A1", AlignSpec(horiz=HorizAlignment.Right))
    assert type(branch) is CustomEditor
    deep = deepcopy(editor)
    deep.extra.append(2)
    assert editor.extra == [1]
    for candidate in (branch, copy(editor), deep):
        candidate.set_cell("A1", "branch")
    editor.save(tmp_path / "original.xlsx")
    assert get_sheet(tmp_path / "original.xlsx", "Sheet1")["A1"].value == "original"


def test_none_returning_operations_mutate(tmp_path):
    path = tmp_path / "mutating.xlsx"
    editor = Editor.create(path)
    assert editor.append_row(["first"]) is None
    assert editor.append_table_at([["second"]], "A2") is None
    assert editor.set_cell("B1", "third") is None
    assert editor.with_polars(pl.DataFrame({"fourth": [4]}), "C1") is None
    assert editor.save(path) is None
    ws = get_sheet(path, "Sheet1")
    assert [ws[c].value for c in ["A1", "A2", "B1", "C1", "C2"]] == ["first", "second", "third", "fourth", 4]


def test_polars_strings_keep_exact_values_and_never_become_formulas(tmp_path):
    values = ["111111119999999", "00123", "12345678901234567890", "=1+1", "1e50", " 01 ", "< & >", ""]
    path = tmp_path / "strings.xlsx"
    editor = Editor.create(path)
    editor.with_polars(pl.DataFrame({"00123": values}, schema={"00123": pl.String}))
    editor.save(path)
    ws = get_sheet(path, "Sheet1")
    assert ws["A1"].value == "00123" and ws["A1"].data_type == "s"
    for row, expected in enumerate(values[:-1], start=2):
        assert ws.cell(row, 1).value == expected
        assert ws.cell(row, 1).data_type == "s"
        assert ws.cell(row, 1).number_format == "@"
    assert ws.column_dimensions["A"].number_format == "@"


@pytest.mark.parametrize("dtype", [pl.Int8, pl.Int16, pl.Int32, pl.Int64, pl.UInt8, pl.UInt16, pl.UInt32, pl.UInt64, pl.Float32, pl.Float64])
def test_numeric_dtype_forces_entire_column_including_outside_written_area(tmp_path, dtype):
    path = tmp_path / "numbers.xlsx"
    book = openpyxl.Workbook()
    ws = book.active
    ws["B20"] = 123
    ws["B20"].number_format = "yyyy-mm-dd"
    ws["C20"] = 123
    ws["C20"].number_format = "0.00000"
    book.save(path)
    editor = Editor.open(path)
    editor.with_polars(pl.DataFrame({"n": [1, None, 2]}, schema={"n": dtype}), "B3")
    editor.save(path)
    ws = get_sheet(path, "Sheet")
    expected = "0.###############" if dtype in (pl.Float32, pl.Float64) else "0"
    assert ws["B4"].value == 1 and ws["B4"].data_type == "n"
    assert ws["B20"].value == 123 and ws["B20"].number_format == expected
    assert ws.column_dimensions["B"].number_format == expected
    assert ws["C20"].number_format == "0.00000"
    with ZipFile(path) as archive:
        assert archive.read("xl/worksheets/sheet1.xml").count(b"<row ") == 5  # no million-row expansion


def test_disabling_column_formats_preserves_template_but_not_type_inference(tmp_path):
    path = tmp_path / "template.xlsx"
    editor = Editor.create(path)
    editor.append_table_at([["old", "old"], ["123", "1"]], "A1")
    editor = editor.set_number_format("A2:B2", "0.00000")
    editor.with_polars(pl.DataFrame({"id": ["00123"], "n": [12]}), force_column_formats=False)
    editor.save(path)
    ws = get_sheet(path, "Sheet1")
    assert ws["A2"].value == "00123" and ws["A2"].data_type == "s"
    assert ws["A2"].number_format == ws["B2"].number_format == "0.00000"
    assert ws["B2"].value == 12
    assert ws.column_dimensions["A"].number_format == "General"


@pytest.mark.parametrize("dtype,fmt", [(pl.String, "@"), (pl.Int64, "0"), (pl.Float64, "0.###############")])
def test_empty_typed_columns_still_have_a_column_format(tmp_path, dtype, fmt):
    path = tmp_path / "empty.xlsx"
    editor = Editor.create(path)
    editor.with_polars(pl.DataFrame(schema={"column": dtype}))
    editor.save(path)
    assert get_sheet(path, "Sheet1").column_dimensions["A"].number_format == fmt


def test_nonfinite_numbers_are_valid_text_cells(tmp_path):
    path = tmp_path / "finite.xlsx"
    editor = Editor.create(path)
    editor.with_polars(pl.DataFrame({"n": [float("nan"), float("inf"), -float("inf"), 1.25]}))
    editor.append_row(["nan", "inf", "-inf"])
    editor.save(path)
    ws = get_sheet(path, "Sheet1")
    assert [ws.cell(r, 1).value for r in range(2, 6)] == ["nan", "inf", "-inf", 1.25]
    assert [ws.cell(6, c).data_type for c in range(1, 4)] == ["s", "s", "s"]


def test_decimal_temporal_and_1904_epoch(tmp_path):
    path = tmp_path / "date1904.xlsx"
    book = openpyxl.Workbook()
    book.epoch = CALENDAR_MAC_1904
    book.save(path)
    editor = Editor.open(path)
    editor.with_polars(pl.DataFrame({
        "decimal": pl.Series([Decimal("123.450")], dtype=pl.Decimal(12, 3)),
        "date": [date(2025, 1, 2)], "dt": [datetime(2025, 1, 2, 3, 4, 5)],
        "time": [time(12, 30)], "duration": [timedelta(hours=30)],
    }))
    editor.save(path)
    ws = get_sheet(path, "Sheet")
    assert ws["A2"].value == 123.45 and ws["A2"].number_format == "0.000"
    assert ws["B2"].value == datetime(2025, 1, 2)
    assert ws["C2"].value == datetime(2025, 1, 2, 3, 4, 5)
    assert ws["D2"].value == time(12, 30)
    assert ws["E2"].value == timedelta(hours=30)


def test_append_beyond_z_and_ordered_cell_insertion(tmp_path):
    path = tmp_path / "columns.xlsx"
    editor = Editor.create(path)
    editor.append_row([str(i) for i in range(30)])
    editor.set_cell("D2", "4")
    editor.set_cell("b2", "2")
    editor.set_cell("C2", "3")
    editor.append_table_at([["1", "20", "30"]], "A2")
    editor.save(path)
    ws = get_sheet(path, "Sheet1")
    assert ws["AD1"].value == 29
    assert [ws.cell(2, c).value for c in range(1, 5)] == [1, 20, 30, 4]


@pytest.mark.parametrize("coord", ["A0", "XFE1", "A1048577", "1", "A-1", "A1x", "A18446744073709551616"])
def test_invalid_coordinates_fail_without_damaging_state(tmp_path, coord):
    path = tmp_path / "invalid.xlsx"
    editor = Editor.create(path)
    editor.set_cell("A1", "original")
    for operation in [lambda: editor.set_cell(coord, "bad"), lambda: editor.append_table_at([["bad"]], coord), lambda: editor.set_fill(coord, "FF0000")]:
        with pytest.raises(RuntimeError):
            operation()
    editor.set_cell("B1", "alive")
    editor.save(path)
    assert get_sheet(path, "Sheet1")["A1"].value == "original"


def test_failed_sheet_selection_and_iterator_leave_editor_usable(tmp_path):
    path = tmp_path / "failure.xlsx"
    editor = Editor.create(path)
    editor.set_cell("A1", "original")
    with pytest.raises(RuntimeError):
        editor.with_worksheet("Missing")
    class BrokenFrame:
        columns = ["a"]
        def iter_rows(self):
            yield (1,)
            raise ValueError("broken row source")
    with pytest.raises(ValueError, match="broken row source"):
        editor.with_polars(BrokenFrame())
    editor.set_cell("B1", "alive")
    editor.save(path)
    assert get_sheet(path, "Sheet1")["A1"].value == "original"


def test_write_into_self_closing_styled_row(tmp_path):
    path = tmp_path / "selfclosing.xlsx"
    editor = Editor.create(path)
    with ZipFile(path) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    parts["xl/worksheets/sheet1.xml"] = parts["xl/worksheets/sheet1.xml"].replace(b"<sheetData/>", b'<sheetData><row r="2"/></sheetData>')
    with ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    editor = Editor.open(path)
    editor.set_cell("A2", "works")
    editor.save(path)
    assert get_sheet(path, "Sheet1")["A2"].value == "works"


def test_styled_dataframe_dimensions_work_with_read_only_readers(tmp_path):
    path = tmp_path / "dimensions.xlsx"
    editor = Editor.create(path)
    editor.with_polars(pl.DataFrame({"id": ["0001", "0002"], "n": [1, 2]}), "C5")
    editor.save(path)
    book = openpyxl.load_workbook(path, read_only=True)
    ws = book["Sheet1"]
    assert ws.max_row == 7 and ws.max_column == 4
    assert list(ws.values)[6][2:] == ("0002", 2)
    book.close()


@pytest.mark.parametrize("selector", ["", ":", "A:", "1:A", "XFE:XFE"])
def test_invalid_query_selectors_raise_exceptions(tmp_path, selector):
    editor = Editor.create(tmp_path / "query.xlsx")
    with pytest.raises(RuntimeError):
        editor.last_rows_index(selector)


@pytest.mark.parametrize("selector", ["", ":", "A0:B2", "XFE:", "A1048577:"])
def test_invalid_style_selectors_raise_without_panicking(tmp_path, selector):
    editor = Editor.create(tmp_path / "style-selector.xlsx")
    with pytest.raises(RuntimeError):
        editor.set_number_format(selector, "0.00")


def test_escaped_style_attributes_round_trip_after_reopening(tmp_path):
    path = tmp_path / "escaped-styles.xlsx"
    font = 'Font & "Name"'
    fmt = '0.00" & USD"'
    editor = Editor.create(path)
    editor.set_cell("A1", "1.25")
    editor = editor.set_font("A1", font, 12).set_number_format("A1", fmt)
    editor.save(path)
    editor = Editor.open(path)
    editor = editor.set_font("A1", font, 12).set_fill("A1", "00FF00")
    editor.save(path)
    ws = get_sheet(path, "Sheet1")
    assert ws["A1"].font.name == font
    assert ws["A1"].number_format == fmt
    with ZipFile(path) as archive:
        assert archive.read("xl/styles.xml").count(b"<font>") == 2


@pytest.mark.parametrize("selector", ['A1:B2" bad="1', "A0:B2", "B2:A1", "XFE1:XFE2"])
def test_invalid_merge_ranges_do_not_corrupt_xml(tmp_path, selector):
    editor = Editor.create(tmp_path / "merge.xlsx")
    with pytest.raises(RuntimeError):
        editor.merge_cells(selector)


def test_invalid_style_ids_report_normal_errors_and_preserve_input(tmp_path):
    path = tmp_path / "badstyle.xlsx"
    editor = Editor.create(path)
    editor.set_cell("A1", "original")
    editor.save(path)
    with ZipFile(path) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    parts["xl/worksheets/sheet1.xml"] = parts["xl/worksheets/sheet1.xml"].replace(b'<c r="A1"', b'<c r="A1" s="999999"')
    with ZipFile(path, "w") as archive:
        for name, data in parts.items(): archive.writestr(name, data)
    editor = Editor.open(path)
    with pytest.raises(RuntimeError, match="style id"):
        editor.set_fill("A1", "FF0000")
    with pytest.raises(RuntimeError, match="style id"):
        editor.set_fill("A1:B2", "FF0000")
    editor.set_cell("B1", "alive")
    editor.save(tmp_path / "saved.xlsx")
    with ZipFile(tmp_path / "saved.xlsx") as archive:
        assert b"alive" in archive.read("xl/worksheets/sheet1.xml")


def test_failed_save_preserves_existing_output(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "precious.xlsx"
    editor = Editor.create(source)
    output.write_bytes(b"preserve existing output")
    with pytest.raises(RuntimeError):
        editor.save(output / "child.xlsx")  # parent is a file, so output creation fails
    assert output.read_bytes() == b"preserve existing output"
    assert not list(tmp_path.glob(".excelsior-save-*"))


def test_save_without_path_and_save_as_do_not_rebind_editor(tmp_path):
    original, copied = tmp_path / "original.xlsx", tmp_path / "copy.xlsx"
    editor = Editor.create(original)
    before = original.read_bytes()
    editor.set_cell("A1", "first")
    assert editor.save(copied) is None
    assert original.read_bytes() == before
    assert get_sheet(copied, "Sheet1")["A1"].value == "first"
    editor.set_cell("B1", "second")
    assert editor.save() is None
    assert get_sheet(original, "Sheet1")["B1"].value == "second"
    assert get_sheet(copied, "Sheet1")["B1"].value is None


@pytest.mark.parametrize("factory", [Editor.open, lambda path: Scanner(path).open_editor()])
def test_save_default_path_is_retained_by_factories_and_functional_copies(tmp_path, factory):
    path = tmp_path / "opened.xlsx"
    Editor.create(path)
    editor = factory(path).add_worksheet("New").set_column_width("A", 20)
    editor.with_polars(pl.DataFrame({"id": ["0001"]}))
    editor.save()
    assert get_sheet(path, "New")["A2"].value == "0001"


def test_save_default_path_survives_working_directory_change(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    editor = Editor.create("relative.xlsx")
    editor.set_cell("A1", "stays here")
    elsewhere = tmp_path / "elsewhere"
    elsewhere.mkdir()
    monkeypatch.chdir(elsewhere)
    editor.save()
    assert not (elsewhere / "relative.xlsx").exists()
    assert get_sheet(tmp_path / "relative.xlsx", "Sheet1")["A1"].value == "stays here"
