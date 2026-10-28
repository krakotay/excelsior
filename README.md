# excelsior
[![Rust](https://github.com/krakotay/excelsior/actions/workflows/rust.yml/badge.svg)](https://github.com/krakotay/excelsior/actions/workflows/rust.yml)
[![Build maturin wheels](https://github.com/krakotay/excelsior/actions/workflows/release.yml/badge.svg?branch=master)](https://github.com/krakotay/excelsior/actions/workflows/release.yml)

[pypi link](https://pypi.org/project/excelsior-fast/)

Just 
```
pip install excelsior-fast
```
or 
```
uv pip install excelsior-fast
```

A small project for quickly modifying `.xlsx` workbooks from Rust or Python. Can be 200+ times faster [than openpyxl](/python-bindings/speed_tests/speed-test-polars-openpyxl-excelsior-0.11.3.md) without [openpyxl errors](/python-bindings/speed_tests/about-speed-test.md). 
It consists of two crates:

* **rust-core** – the core library that works directly with spreadsheet XML.
* **python-bindings** – Python wrapper built with `pyo3` and `maturin`.

The library lets you append rows or tables, modify individual cells and
save the workbook back to disk without loading the entire file into memory.

## Create a workbook from scratch

CPython 3.10–3.15 is supported. The base package has no Python runtime
dependencies. Polars integration uses Python's public `columns` and
`iter_rows()` APIs and does not need PyArrow or a specific Polars build.

```python
from excelsior import Editor, create_excel

# Create, write, and save to the same path; no template is needed.
editor = Editor.create("new.xlsx", "Data")
editor.append_table_at([["name", "value"], ["alpha", "42"]], "A1")
editor.save()

# Or create an empty XLSX without opening an editor:
create_excel("empty.xlsx")
```

Creation replaces an existing destination. Missing parent directories are
created automatically. `Editor.open(path)` and `Scanner(path).open_editor()`
open the first sheet if no sheet name is supplied.

## Write a Polars DataFrame

Use your existing Polars installation, or install the optional extra with
`pip install "excelsior-fast[polars]"`. No upper version pin is imposed.

```python
import polars as pl
from excelsior import Editor

editor = Editor.create("report.xlsx", "Index")
frame = pl.DataFrame({"name": ["alpha", "beta"], "value": [10, 20]})
editor = editor.add_worksheet("Data")
editor.with_polars(frame.slice(0, 2), "B3")
editor.save()
```

Headers are included. Dates and datetimes receive Excel number formats;
booleans become native Excel boolean cells. XLSX does not
store timezones, so datetimes retain their displayed local date and time.

## Edit an existing workbook

```python
from excelsior import Scanner

scanner = Scanner("input.xlsx")
editor = scanner.open_editor()
editor.append_table_at([["Hello", "World"]], "B4")
editor.save("output.xlsx")
```

The Rust crate also exports `rust_core::create_excel(path, sheet_name)` and
`XlsxEditor::create(path, sheet_name)`. See [docs/usage.md](docs/usage.md) for more.

## Local development

Install `just` and `uv`, then run from the repository root:

```sh
just build
just dev
just test
```

`uv` creates/uses `python-bindings/.venv` and installs maturin automatically.
No environment activation or PowerShell is required. These same recipes are
available from `python-bindings/`. Pass extra maturin options as one quoted
argument, for example `just build "--interpreter python3.15"`. To choose the
Python used for local development, set `UV_PYTHON`, for example
`UV_PYTHON=3.14 just dev` in a POSIX shell. Benchmark recipes in
`python-bindings/justfile` additionally require `hyperfine` on PATH.

## Build and release

GitHub Actions builds and tests wheels for Python 3.10–3.15 on Linux,
Windows, and macOS, plus one source distribution. The **Build and publish to
PyPI** workflow publishes through PyPI Trusted Publishing when run manually
with version `0.15.0`, or when the matching `v0.15.0` tag is pushed. Ordinary
branch pushes only run CI. **Build maturin wheels** creates a GitHub release
without publishing to PyPI. Release inputs and tags must match `Cargo.toml`.

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

## Local installation across Python versions

A wheel tagged `cp314-cp314` only supports CPython 3.14. `uv add` resolves the
whole `requires-python` range, so that wheel cannot satisfy a project which also
supports 3.13 or 3.15. Use the local source directory instead:

```sh
uv add /path/to/excelsior/python-bindings
```

Or build a source archive with `just sdist` and `uv add /path/to/archive.tar.gz`.
Both allow uv to build for the selected interpreter (Rust is required locally).
Published releases provide separate wheels for supported Python versions and
platforms. `uv pip install ./file.whl` only installs into the selected environment;
it does not create a portable project lock.

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
