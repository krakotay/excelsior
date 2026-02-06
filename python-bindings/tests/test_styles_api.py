from pathlib import Path

from excelsior import AlignSpec, Editor, HorizAlignment, VertAlignment, create_excel

from helpers_excel import assert_with_libreoffice, get_sheet


def test_styles_alignment_font_fill_border_and_remove(tmp_path: Path) -> None:
    src = tmp_path / "styles.xlsx"
    out = tmp_path / "styles_out.xlsx"
    out2 = tmp_path / "styles_out2.xlsx"

    editor = Editor.create(str(src), "Styles")
    editor.append_table_at([["x", "y", "z"], ["1", "2", "3"]], "A1")

    editor.set_font("A1:C1", "Calibri", 12.0, bold=True, italic=False)
    editor.set_fill("A1:C1", "FFFF00")
    editor.set_border("A1:C2", "thin")
    editor.set_alignment(
        "A1:C1",
        AlignSpec(
            horiz=HorizAlignment.Center,
            vert=VertAlignment.Center,
            wrap=True,
        ),
    )
    editor.set_number_format("B2", "0.00")
    editor.save(str(out))

    ws = get_sheet(out, "Styles")
    assert ws["A1"].font.name == "Calibri"
    assert bool(ws["A1"].font.bold)
    assert ws["A1"].fill.fgColor.rgb in {"00FFFF00", "FFFFFF00", "FFFF00"}
    assert ws["A1"].alignment.horizontal == "center"
    assert ws["A1"].alignment.vertical == "center"
    assert ws["A1"].alignment.wrap_text is True
    assert ws["B2"].number_format == "0.00"

    editor = Editor.open(str(out), "Styles")
    editor.remove_style("A1:C1")
    editor.save(str(out2))

    ws_after = get_sheet(out2, "Styles")
    assert ws_after["A1"].alignment.horizontal != "center"

    assert_with_libreoffice(out2)


def test_alignment_open_column_selector_from_cell(tmp_path: Path) -> None:
    src = tmp_path / "styles_col_from.xlsx"
    out = tmp_path / "styles_col_from_out.xlsx"

    editor = Editor.create(str(src), "Styles")
    editor.append_table_at(
        [["h1", "h2"], ["r2c1", "r2c2"], ["r3c1", "r3c2"]],
        "A1",
    )
    editor.set_alignment(
        "A2:",
        AlignSpec(horiz=HorizAlignment.Center, vert=VertAlignment.Center, wrap=True),
    )
    editor.save(str(out))

    ws = get_sheet(out, "Styles")
    assert ws["A1"].alignment.horizontal != "center"
    assert ws["A2"].alignment.horizontal == "center"
    assert ws["A3"].alignment.horizontal == "center"
