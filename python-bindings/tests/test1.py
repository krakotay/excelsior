from excelsior import scan_excel, Editor
import os
base_dir = os.path.dirname(os.path.abspath(__file__))
assert scan_excel(os.path.join(base_dir, "../../test/test_sum.xlsx")) == ["Sheet1"]
editor = Editor(os.path.join(base_dir, "../../test/test_last_row_index.xlsx"), "Sheet1")
assert editor.last_row_index("A") == 4
assert editor.last_row_index("B") == 5
assert editor.last_row_index("C") == 8
assert editor.last_row_index("D") == 8
assert editor.last_rows_index("A:D") == [4, 5, 8, 8]