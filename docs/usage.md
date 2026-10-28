# Usage Guide

This document describes how to use the `excelsior` project.

## Rust library

The core functionality resides in the `rust-core` crate. It exposes the
`XlsxEditor` type which can open an existing `.xlsx` file and modify its
contents.

### Creating a workbook

```rust
use rust_core::{create_excel, XlsxEditor};

create_excel("empty.xlsx", "Sheet1")?;
let mut editor = XlsxEditor::create("new.xlsx", "Data")?;
editor.append_row(["Hello", "42"])?;
editor.save("new.xlsx")?;
```

Creation replaces an existing file and creates missing parent directories.

### Opening a sheet
```rust
use rust_core::{XlsxEditor, scan};

let sheet_names = scan("test.xlsx")?;
let mut editor = XlsxEditor::open("test.xlsx", &sheet_names[0])?;
```

### Adding data
Append a row to the end of the current worksheet:
```rust
editor.append_row(["Hello", "World"])?;
```

Insert a table starting from a specific cell:
```rust
let rows = vec![vec!["1"], vec!["2"]];
editor.append_table_at("C1", rows)?;
```

Set the value of an individual cell:
```rust
editor.set_cell("A1", "Some text")?;
```

### Saving
Write the modified workbook to a new file:
```rust
editor.save("output.xlsx")?;
```

## Python bindings

Bindings are provided in the `python-bindings` crate. After building with
`maturin`, the package `excelsior` offers an interface similar to the
Rust API.

### Creating and writing

```python
from excelsior import create_excel, Editor

create_excel("empty.xlsx", "Data")
editor = Editor.create("new.xlsx", "Data")
editor.append_row(["Hello", "42"])
editor.save("new.xlsx")
```

`create_excel` is exported from the package root. `Editor.create` creates the
workbook and opens its first sheet. Both accept strings or `pathlib.Path`.
`Editor.open(path)` opens the first sheet by default, as does
`Scanner(path).open_editor()`.

### Polars

```python
import polars as pl
from excelsior import Editor

editor = Editor.create("data.xlsx")
editor.add_worksheet("Data").with_polars(pl.DataFrame({"count": [1, 2]}), "B3")
editor.save("data.xlsx")
```

The Python adapter only uses `columns` and `iter_rows()`. It has no compile-time
Polars dependency, never calls `to_arrow`, and does not import Polars or
PyArrow. Install your preferred Polars version or use the optional
`excelsior-fast[polars]` extra. Empty frames with a schema write headers;
frames without columns leave the sheet unchanged. Dates/datetimes are written
as Excel serials with number formats; booleans remain `TRUE`/`FALSE` text.
The adapter validates XLSX row/column limits before editing the sheet.

### Existing workbooks

```python
from excelsior import scan_excel, Editor

sheets = scan_excel("tests/test_sum.xlsx")
editor = Editor("tests/test_sum.xlsx", sheets[0])
editor.append_row(["Hello", "World"])
editor.save("tests/result.xlsx")
```

Refer to `python-bindings/tests` for more examples.
