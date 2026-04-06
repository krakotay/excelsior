from excelsior import Editor
import os
import polars as pl

base_dir = os.path.dirname(os.path.abspath(__file__))

editor = Editor(os.path.join(base_dir, "../../test/100mb.xlsx"), "Tablo3")
df = pl.DataFrame(
    {
        "int": [10**5, 20**5, 30**5] * 1000,
        "float": [1.1, 2.2, 3.3] * 1000,
        "None": [None] * 3000,
        "string0": ["a" * 100, "b" * 100, "c" * 100] * 1000,
        "string1": ["a" * 100, "b" * 100, "c" * 100] * 1000,
        "string2": ["a" * 100, "b" * 100, "c" * 100] * 1000,
        "string3": ["a" * 100, "b" * 100, "c" * 100] * 1000,
        "string4": ["a" * 100, "b" * 100, "c" * 100] * 1000,
        "string5": ["a" * 100, "b" * 100, "c" * 100] * 1000,
        "string6" : ["a" * 100, "b" * 100, "c" * 100] * 1000,
        "string7" : ["a" * 100, "b" * 100, "c" * 100] * 1000,
    }
)
editor.with_polars(df, "D10") # THIS WILL NOT CHANGE FORMATTING OF EXISTING CELLS!
editor.save(os.path.join(base_dir, "100mb_excelsior_polars_1.xlsx"))
