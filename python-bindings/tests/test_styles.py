from pathlib import Path

from excelsior import Scanner, AlignSpec, HorizAlignment

from excel_check import (
    XL_CENTER,
    assert_excel_file,
    check_borders_thin,
    check_font_range,
    check_merge,
)


base_dir = Path(__file__).resolve().parent
inp_filename = (base_dir / "../../test/style_test.xlsx").resolve()
out_filename = (base_dir / "../../test/style_test_out_py.xlsx").resolve()


scanner = Scanner(str(inp_filename))
editor = scanner.open_editor(scanner.get_sheets()[0])
# .set_fill("B24:B28", "FFFF00")\
editor.set_font("D4:D8", "Arial", 12.0, True, False).set_border(
    "A1:C3", "thin"
).set_font(
    "A1:C3",
    "Calibri",
    10.0,
    False,
    True,
    AlignSpec(horiz=HorizAlignment.Center),
).merge_cells("B12:D12")
editor.save(str(out_filename))


assert_excel_file(
    out_filename,
    lambda ws: check_font_range(ws, "D4:D8", "Arial", 12.0, True, False),
    lambda ws: check_font_range(
        ws,
        "A1:C3",
        "Calibri",
        10.0,
        False,
        True,
        horiz=XL_CENTER,
    ),
    lambda ws: check_borders_thin(ws, "A1:C3"),
    lambda ws: check_merge(ws, "B12:D12", expected_addr="$B$12:$D$12"),
)

print('Done')