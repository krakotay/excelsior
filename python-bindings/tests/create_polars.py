import polars as pl
from excelsior import create_excel, Scanner
import os

h = 3000
w = 3
base_dir = os.path.dirname(os.path.abspath(__file__))
out_path = os.path.join(base_dir, "../../test/new_file.xlsx")
out_path_2 = os.path.join(base_dir, "../../test/new_file_out.xlsx")

df = pl.DataFrame({"dog": [*(["dog"] * 1000)], "cat": [*(["cat"] * 1000)]})

create_excel(out_path)
scanner = Scanner(out_path)
editor = scanner.open_editor(scanner.get_sheets()[0])
editor.with_polars(df)
editor.save(out_path_2)
