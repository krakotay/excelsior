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
editor.save("new.xlsx")

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
editor.add_worksheet("Data").with_polars(pl.DataFrame({"value": [1, 2]}))
editor.save("report.xlsx")
```

The package exposes `excelsior.__version__` to identify the loaded build.
