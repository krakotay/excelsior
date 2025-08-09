from excelsior import AlignSpec, HorizAlignment, Scanner
import os
base_dir = os.path.dirname(os.path.abspath(__file__))
scanner = Scanner(os.path.join(base_dir, "../tests/100mb.xlsx"))
editor = scanner.open_editor("Tablo3")
def generate_table(width: int, height: int):
    table = [[str(i) for i in range(width)] for _i in range(height)]
    return table
data = generate_table(5, 2000)
last_col = max(editor.last_rows_index("A:E")) + 1
for row in data:
    editor.append_row(row)
editor.set_font(f'A{last_col}:E{last_col + 2000}', "Arial", size=14, align=AlignSpec(HorizAlignment.Center))  
editor.set_border(f'A{last_col}:E{last_col + 2000}', 'thick')  
editor.save(os.path.join(base_dir, "100mb_excelsior_style.xlsx"))
