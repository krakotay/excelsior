from pathlib import Path

from excelsior import Editor, Scanner, create_excel

from helpers_excel import column_width_map, get_sheet


def test_scanner_open_editor_without_sheet_name(tmp_path: Path) -> None:
    src = tmp_path / "scanner.xlsx"
    out = tmp_path / "scanner_out.xlsx"
    create_excel(str(src), "First")

    scanner = Scanner(str(src))
    editor = scanner.open_editor()
    editor.set_cell("A1", "scanner-default")
    editor.save(str(out))

    ws = get_sheet(out, "First")
    assert str(ws["A1"].value) == "scanner-default"


def test_editor_open_and_chain_methods(tmp_path: Path) -> None:
    src = tmp_path / "chain.xlsx"
    out = tmp_path / "chain_out.xlsx"

    editor = Editor.create(str(src), "Main")
    editor.set_column_width_range("A:C", 16.0)
    editor.append_table_at([["h1", "h2", "h3"], ["1", "2", "3"]], "A1")
    editor.save(str(out))

    ws = get_sheet(out, "Main")
    assert str(ws["A2"].value) == "1"
    widths = column_width_map(out)
    assert abs(widths["B"] - 16.0) < 0.001
