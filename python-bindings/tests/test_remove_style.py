from pathlib import Path

from excelsior import Scanner, AlignSpec, HorizAlignment

from excel_check import (
    XL_CENTER,
    assert_excel_file,
    check_font_range,
    check_font_not_range,
)


base_dir = Path(__file__).resolve().parent
inp_filename = (base_dir / "../../test/style_test.xlsx").resolve()
styled_out = (base_dir / "../../test/style_test_styled_rm_py.xlsx").resolve()
cleared_out = (base_dir / "../../test/style_test_cleared_rm_py.xlsx").resolve()


def _prepare_editor():
    scanner = Scanner(str(inp_filename))
    editor = scanner.open_editor(scanner.get_sheets()[0])
    return editor


editor = _prepare_editor()

# 1) Применяем стили и сохраняем первый файл
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
editor.save(str(styled_out))

# Проверяем, что стили действительно применены
assert_excel_file(
    styled_out,
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
)

# 2) Снимаем стили с этих же диапазонов и сохраняем во второй файл
editor.remove_style("A1:D2")
    # .remove_style('B:').remove_style('C:').remove_style('D:')
editor.save(str(cleared_out))

# Проверяем, что диапазоны больше НЕ имеют тех же стилей
assert_excel_file(
    cleared_out,
    lambda ws: check_font_not_range(ws, "A1:D2", "Arial", 12.0, True, False),
    lambda ws: check_font_not_range(
        ws,
        "A1:D2",
        "Calibri",
        10.0,
        False,
        True,
        horiz=XL_CENTER,
    ),
)
