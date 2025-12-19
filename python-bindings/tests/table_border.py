from excelsior import Scanner
from excel_check import assert_excel_file, check_borders_thin, check_font_range, check_merge
import os

base_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(base_dir, "../../test/style_test.xlsx")
file_path_out = os.path.join(base_dir, "../../test/style_test_borders_py.xlsx")
scanner = Scanner(file_path)
print(scanner.get_sheets())
editor = scanner.open_editor(scanner.get_sheets()[0])
editor.set_border("D4:G8", "thin").merge_cells("D4:G4").set_font(
    "D4:G4", "Arial", 12, False, False
)
editor.save(file_path_out)

assert_excel_file(
    file_path_out,
    lambda ws: check_borders_thin(ws, "D4:G8"),
    lambda ws: check_merge(ws, "D4:G4", expected_addr="$D$4:$G$4"),
    lambda ws: check_font_range(ws, "D4:G4", "Arial", 12.0, False, False),
)
