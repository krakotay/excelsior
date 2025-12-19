import pandas as pd
import polars as pl
import os
import shutil

base_dir = os.path.dirname(os.path.abspath(__file__))
in_filename = os.path.join(base_dir, "../../test/100mb.xlsx")
out_filename = os.path.join(base_dir, "100mb_openpyxl_pandas.xlsx")
shutil.copyfile(in_filename, out_filename)

df = pl.DataFrame(
    {
        "int": [1, 2, 3] * 1000,
        "float": [1.1, 2.2, 3.3] * 1000,
        "string": ["a" * 100, "b" * 100, "c" * 100] * 1000,
    }
)

# just a demo: polars CAN'T modify excel files with his own engines
with pd.ExcelWriter(
    out_filename, mode="a", engine="openpyxl", if_sheet_exists="overlay"
) as writer:
    df.to_pandas().to_excel(writer, "pandas_df")
