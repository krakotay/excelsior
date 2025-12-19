from excelsior import Editor
import os
import polars as pl

base_dir = os.path.dirname(os.path.abspath(__file__))

editor = Editor(os.path.join(base_dir, "../../test/100mb.xlsx"), "Tablo3")
last_row = editor.last_row_index("A1")
df = pl.DataFrame(
    {
        "int": [1, 2, 3] * 1000,
        "float": [1.1, 2.2, 3.3] * 1000,
        "string": ["a" * 100, "b" * 100, "c" * 100] * 1000,
    }
)
editor.add_worksheet('polars_df').with_polars(df)
editor.save(os.path.join(base_dir, "100mb_excelsior_polars.xlsx"))
