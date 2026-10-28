from pathlib import Path

from excelsior import Editor, Scanner


def test_scan_and_open_generated_workbook(tmp_path: Path) -> None:
    path = tmp_path / "scanner.xlsx"
    editor = Editor.create(path, "dog")
    editor = editor.add_worksheet("cat")
    editor.save(path)

    scanner = Scanner(path)
    assert scanner.get_sheets() == ["dog", "cat"]
    editor = scanner.open_editor("dog")
    assert isinstance(editor, Editor)
    assert editor.last_rows_index("A:D") == [0, 0, 0, 0]
