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
editor.save("new.xlsx")

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
editor.add_worksheet("Data").with_polars(frame.slice(0, 2), "B3")
editor.save("report.xlsx")
```

Headers are included. Dates and datetimes receive Excel number formats;
booleans retain the existing `TRUE`/`FALSE` text representation. XLSX does not
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

## Build and release

GitHub Actions builds and tests wheels for Python 3.10–3.15 on Linux,
Windows, and macOS, plus one source distribution. The **Build and publish to
PyPI** workflow publishes through PyPI Trusted Publishing when run manually
with version `0.14.0`, or when the matching `v0.14.0` tag is pushed. Ordinary
branch pushes only run CI. **Build maturin wheels** creates a GitHub release
without publishing to PyPI. Release inputs and tags must match `Cargo.toml`.
