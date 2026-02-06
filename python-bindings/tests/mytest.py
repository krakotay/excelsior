from excelsior import Scanner, AlignSpec
import os

# from excelsior.excelsior import HorizAlignment, VertAlignment

base_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(base_dir, "../../100mb.xlsx")
out_path = os.path.join(base_dir, "../../test/100mb_out.xlsx")

scanner = Scanner(file_path)
editor = scanner.open_editor(scanner.get_sheets()[0])

editor.append_table_at([[str(k) for k in list(range(50))] for _k in list(range(50))], "B4")
# editor = 
editor.save(out_path)

print("done")
