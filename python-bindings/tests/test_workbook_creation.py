from pathlib import Path

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
