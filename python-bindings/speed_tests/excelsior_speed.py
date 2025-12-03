from excelsior import Editor, HorizAlignment, AlignSpec
import os
base_dir = os.path.dirname(os.path.abspath(__file__))

editor = Editor(os.path.join(base_dir, "../../test/100mb.xlsx"), "Tablo3")
def generate_table(width: int, height: int):
    table = [['a' + str(i + j) * 10 for i in range(width)] for j in range(height)]
    return table

data = generate_table(5, 200)

last_col = max(editor.last_rows_index("A:E")) + 1
for row in data:
    editor.append_row(row)
editor.set_alignment(f'A{last_col}:E{last_col + 200}', AlignSpec(HorizAlignment.Center))    

editor.save(os.path.join(base_dir, "100mb_excelsior.xlsx"))
