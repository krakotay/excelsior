# excelsior
[![Rust](https://github.com/krakotay/excelsior/actions/workflows/rust.yml/badge.svg)](https://github.com/krakotay/excelsior/actions/workflows/rust.yml)
[![Build maturin wheels](https://github.com/krakotay/excelsior/actions/workflows/release.yml/badge.svg?branch=master)](https://github.com/krakotay/excelsior/actions/workflows/release.yml)

A project for fast and predictable `.xlsx` editing from Rust and Python.
It consists of two crates:

* **rust-core** – the core library that works directly with spreadsheet XML.
* **python-bindings** – Python wrapper built with `pyo3` and `maturin`.

The library supports:

* opening existing workbooks,
* creating an empty workbook from scratch,
* editing cells/tables/styles/column widths,
* saving changes back to disk without loading the full workbook model.

For detailed usage examples see [the usage guide](https://github.com/krakotay/excelsior/blob/master/docs/usage.md).

## Create and write a new XLSX

```python
from excelsior import Editor, create_excel

editor = Editor.create("new.xlsx", "Data")
editor.append_row(["Hello", "42"])
editor.save()

create_excel("empty.xlsx")
```

CPython 3.10–3.15 is supported. The base package has no Python runtime
dependencies. Creation replaces an existing destination and creates missing
parent directories.

## Optional Polars support

`Editor.with_polars(frame, start_cell=None)` runs entirely in Python using
`frame.columns` and `frame.iter_rows()`. It neither imports Polars nor converts
through Arrow. Use your existing Polars installation or install
`excelsior-fast[polars]`; there is no upper Polars version pin.

```python
import polars as pl
from excelsior import Editor

editor = Editor.create("report.xlsx")
editor = editor.add_worksheet("Data")
editor.with_polars(pl.DataFrame({"value": [1, 2]}))
editor.save()
```

The package exposes `excelsior.__version__` to identify the loaded build.

## Python method contracts in 0.15

Methods returning `Editor` return an independent snapshot, including unsaved
changes. Always keep the result:

```python
editor = editor.add_worksheet("Data")
editor = editor.set_column_width("A", 24)
editor = editor.set_number_format("B:", "0.00")
branch = editor.with_worksheet("Sheet1")  # editor still selects Data
```

This applies to worksheet operations, styling, merges, column widths,
`remove_style`, and `copy`. Ignoring their return value leaves the input
unchanged. Copying a large workbook costs memory proportional to its edited XML.
Rust methods explicitly taking `&mut self` retain Rust's mutating contract.

Methods returning `None` modify the current editor: `set_cell`, `append_row`,
`append_table_at`, and `with_polars`. `save()` writes to the file passed to create/open. `save(other_path)` writes a
copy, retaining the original path for subsequent `save()` calls. Queries `last_row_index`/`last_rows_index` return values without editing.

## Dataframe types and column formats

`with_polars(frame, start_cell=None, *, force_column_formats=True)` preserves
cell types. Strings such as `"00123"`, `"111111119999999"`, and `"=1+1"` stay
text, including headers. Integers, floats, and decimals become numbers; nulls
become empty cells. NaN and infinity are stored as text because Excel has no
corresponding numeric values.

By default dataframe dtypes force the number format of **each entire destination
column**, including cells outside the written area and column defaults. Integer
columns use `0`, floats use `0.###############`, strings/categorical/enum use
`@`, and decimals use their declared scale. Dates, datetimes, times, and durations
use their corresponding Excel formats; both 1900 and 1904 date systems work.
Explicit cell styling and column widths are retained.

```python
editor.with_polars(frame)  # mutate; force formats for all affected columns
editor.with_polars(frame, "B3", force_column_formats=False)  # keep template formats
```

Disabling forcing preserves template formats while values still retain their
explicit types. Excel numeric cells have about 15 significant decimal digits;
store identifiers or longer exact numbers in Polars String columns. The fixed
float display shows up to 15 fractional digits; use `set_number_format` after
writing when more suitable scientific/precision formatting is needed.
Nanosecond temporal precision is limited by Polars `iter_rows()` conversion to
Python datetime values. Nested values are stored as text; object columns use their Python value types.

The legacy string APIs `set_cell`, `append_row`, and `append_table_at` retain
numeric/formula inference. Use dataframe String columns to write numeric-looking
text through the typed path.

## Saving a dataframe across worksheets

```python
from excelsior import Editor

MAX_ROWS = 1_000_000  # must be <= 1_048_575: one Excel row is used by headers
editor = Editor.create(f"out/{name}.xlsx")
for i, offset in enumerate(range(0, df.height, MAX_ROWS), start=1):
    if i > 1:
        editor = editor.add_worksheet(f"Sheet{i}")
    editor.with_polars(df.slice(offset, MAX_ROWS))
editor.save()
# editor.save("out/another.xlsx") writes a copy; the original file is unchanged.
```

Saving after every sheet rewrites the entire ZIP each time. Save once after the
loop unless intermediate on-disk checkpoints are needed. Both `save()` and
`save(path)` replace the destination atomically using a temporary file beside it.
