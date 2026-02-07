from excelsior import Scanner
import os
import polars as pl
# from excelsior.excelsior import HorizAlignment, VertAlignment

base_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(base_dir, "../../test/100mb.xlsx")
out_path = os.path.join(base_dir, "../../test/100mb_out.xlsx")

scanner = Scanner(file_path)
editor = scanner.open_editor(scanner.get_sheets()[0])


editor.append_table_at([[str(k) for k in list(range(50))] for _k in list(range(5))], "B4")
h = 3000
w = 3
editor.with_worksheet(scanner.get_sheets()[1]).with_polars(pl.DataFrame({str(i + 1): list(range(w)) for i in range(h)}), "C9")
# editor = 
editor.save(out_path)

print("done")
