from pathlib import Path
import pytest

from excelsior import Editor, Scanner, create_excel

from helpers_excel import assert_with_libreoffice, cell_text, get_sheet


def test_create_empty_excel_and_edit(tmp_path: Path) -> None:
    src = tmp_path / "created.xlsx"
    out = tmp_path / "created_out.xlsx"

    create_excel(str(src), "Data")
    editor = Editor.open(str(src), "Data")
    editor.append_table_at([["name", "value"], ["alpha", "10"]], "A1")
    editor.save(str(out))

    scanner = Scanner(str(out))
    assert scanner.get_sheets() == ["Data"]

    ws = get_sheet(out, "Data")
    assert cell_text(ws["A1"].value) == "name"
    assert cell_text(ws["B2"].value) == "10"

    assert_with_libreoffice(out)


def test_editor_create_shortcut(tmp_path: Path) -> None:
    src = tmp_path / "editor_create.xlsx"
    out = tmp_path / "editor_create_out.xlsx"

    editor = Editor.create(str(src), "Main")
    editor.set_cell("C3", "hello")
    editor.save(str(out))

    ws = get_sheet(out, "Main")
    assert cell_text(ws["C3"].value) == "hello"


def test_editor_open_first_sheet_when_not_specified(tmp_path: Path) -> None:
    src = tmp_path / "default_open.xlsx"
    out = tmp_path / "default_open_out.xlsx"

    create_excel(str(src), "SheetA")
    editor = Editor(str(src))
    editor.set_cell("A1", "ok")
    editor.save(str(out))

    ws = get_sheet(out, "SheetA")
    assert cell_text(ws["A1"].value) == "ok"


def test_create_relative_filename_in_working_directory(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    editor = Editor.create("new.xlsx")
    editor.append_row(["hello", "42"])
    editor.save("new.xlsx")
    ws = get_sheet(tmp_path / "new.xlsx", "Sheet1")
    assert ws["A1"].value == "hello"
    assert ws["B1"].value == 42


@pytest.mark.parametrize("name", ["Я" * 31, 'Data & <values> "2026"'])
def test_create_nested_path_with_unicode_and_escaped_sheet_names(tmp_path, name):
    path = tmp_path / "nested" / "new.xlsx"
    create_excel(path, name)
    assert Scanner(path).get_sheets() == [name]
    assert get_sheet(path, name)["A1"].value is None


@pytest.mark.parametrize(
    "name", ["", "x" * 32, "bad/name", "'quoted", "quoted'", "bad\x00name"]
)
def test_invalid_sheet_name_preserves_existing_file(tmp_path, name):
    path = tmp_path / "existing.xlsx"
    path.write_bytes(b"keep")
    with pytest.raises(RuntimeError):
        create_excel(path, name)
    assert path.read_bytes() == b"keep"


def test_escaped_names_can_be_opened_switched_and_renamed(tmp_path):
    path = tmp_path / "escaped.xlsx"
    name = 'Data & "values"'
    renamed = "Renamed & values"
    editor = Editor.create(path, name)
    editor.set_cell("A1", "original")
    editor = editor.add_worksheet("Other")
    editor = editor.with_worksheet(name)
    editor = editor.rename_worksheet(name, renamed)
    editor.save(path)
    assert Scanner(path).get_sheets() == [renamed, "Other"]
    editor = Editor.open(path, renamed)
    editor.set_cell("B1", "reopened")
    editor.save(path)
    assert get_sheet(path, renamed)["A1"].value == "original"
    assert get_sheet(path, renamed)["B1"].value == "reopened"
