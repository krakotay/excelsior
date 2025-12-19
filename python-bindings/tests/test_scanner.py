from excelsior import Scanner
import os
base_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(base_dir, "../../test/test_scanner.xlsx")
scanner = Scanner(file_path)
print(scanner.get_sheets())
assert scanner.get_sheets() == ["dog", "cat"]

editor = scanner.open_editor("dog")
assert editor.last_row_index("A") == 0
assert editor.last_row_index("B") == 0
assert editor.last_row_index("C") == 0
assert editor.last_row_index("D") == 0
assert editor.last_rows_index("A:D") == [0, 0, 0, 0]

editor = scanner.open_editor("cat")
assert editor.last_row_index("A") == 2
assert editor.last_row_index("B") == 9
assert editor.last_row_index("C") == 22
assert editor.last_row_index("D") == 15
assert editor.last_rows_index("A:D") == [2, 9, 22, 15]

print('passed')