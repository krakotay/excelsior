from pathlib import Path

from excelsior import Scanner, AlignSpec, HorizAlignment
from utils import measure
from excel_check import (
    XL_CENTER,
    assert_excel_file,
    check_font_range,
    check_font_not_range,
    check_borders_thin
)


base_dir = Path(__file__).resolve().parent
inp_filename = (base_dir / "../../test/100mb.xlsx").resolve()
styled_out = (base_dir / "../../test/style_test_styled_rm_py_100mb.xlsx").resolve()
cleared_out = (base_dir / "../../test/style_test_cleared_rm_py_100mb.xlsx").resolve()

def _prepare_editor():
    scanner = Scanner(str(inp_filename))
    editor = scanner.open_editor(scanner.get_sheets()[0])
    return editor

with measure('open file'):
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
)
editor.set_border('G1:H10', 'thin')
editor.set_fill('G1:H10', '#FFFF00')
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
    lambda ws: check_borders_thin(ws, 'G1:H10')
)

# 2) Снимаем стили с этих же диапазонов и сохраняем во второй файл
print('Убираю стили в A1:D2')
with measure('A1:D2'):
    editor.remove_style("A1:D2")
print('Убрал стили в A1:D2. Убираю в G:')
with measure('G:'):
    editor.remove_style('G:')
print('Убрал в G:')

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
    lambda ws: check_borders_thin(ws, 'G1:G10') # Должно выдать ошибку

)
print(cleared_out)
print('Done')