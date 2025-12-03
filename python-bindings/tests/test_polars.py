from excelsior import Scanner
from excel_check import assert_excel_file, check_range_values, on_sheet
import polars as pl
import os

df = pl.DataFrame(
    {
        # "int": [1, 2, 3],
        # "float": [1.1, 2.2, 3.3],
        "string": ["a", "b", "c"],
        # "bool": [True, False, True], # not yet fully supported, only as string
        "long_string": ["123" * 10, "123" * 10, "123" * 10],
    }
)
print(df)
base_dir = os.path.dirname(os.path.abspath(__file__))
inp_filename = os.path.join(base_dir, "../../test/test.xlsx")
out_filename = os.path.join(base_dir, "../../test/test_polars_appended.xlsx")

scanner = Scanner(inp_filename)
editor = scanner.open_editor(scanner.get_sheets()[0])
editor.with_polars(df, "A1")
# editor.add_worksheet("polars_ws").with_polars(df, "B4")
# editor.add_worksheet("polars_ws_2").with_polars(df)
editor.save(out_filename)

expected_values = [df.columns] + (
    df.select(pl.all().cast(pl.Utf8))
    .to_numpy()
    .tolist()
)

assert_excel_file(
    out_filename,
    on_sheet(
        "polars_ws",
        lambda ws: check_range_values(
            ws,
            "B4:F7",
            expected_values,
        ),
    ),
    on_sheet(
        "polars_ws_2",
        lambda ws: check_range_values(
            ws,
            "A1:E4",
            expected_values,
        ),
    ),
)
print('Done')