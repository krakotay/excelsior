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
editor.save()?;
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
editor.save()
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
editor = editor.add_worksheet("Data")
editor.with_polars(pl.DataFrame({"count": [1, 2]}), "B3")
editor.save()
```

The Python adapter uses `columns`, `dtypes`, and `iter_rows()` (with a fallback for objects without dtypes). It has no compile-time
Polars dependency, never calls `to_arrow`, and does not import Polars or
PyArrow. Install your preferred Polars version or use the optional
`excelsior-fast[polars]` extra. Empty frames with a schema write headers;
frames without columns leave the sheet unchanged. Dates/datetimes are written
as Excel serials with number formats; booleans become native boolean cells.
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
unchanged. Snapshots share XML and metadata through Arc; writing detaches the affected component.
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

## Streaming creation with bounded memory (0.15)

Use `WorkbookWriter` for new files whose rows only move forward. It spools the
current sheet as typed records on disk and writes each closed sheet directly
into ZIP. Worksheet XML is not retained in memory. Temporary disk space is
required; the destination is replaced atomically when saving succeeds.

```python
from excelsior import WorkbookWriter

with WorkbookWriter("out/report.xlsx", "Sheet1", compression_level=3) as wb:
    for i, offset in enumerate(range(0, df.height, 1_000_000), start=1):
        if i > 1:
            wb.add_worksheet(f"Sheet{i}")  # mutates, returns None
        wb.with_polars(df.slice(offset, 1_000_000))
```

All writer operations mutate and return `None`; factories return a new writer.
The context manager saves on successful exit and discards temporary files on
an exception. `save()`/`close()` finalizes the writer; subsequent writes are
rejected. Repeated `save(other_path)` copies the completed archive without
changing the original destination. Closed sheets cannot be revisited, edited,
renamed, or copied; use `Editor` for templates, random edits, or snapshots.

`with_polars` appends headers and rows after the previous rows by default.
Explicit start cells must follow previous rows. Column formats and widths can
be set on the active sheet before it closes. `force_column_formats=False`
preserves its configured formats. Python-only adaptation is shared with Editor;
Polars and Arrow are not imported by the package.

`write_rows(iterator)` and `append_row(values)` preserve Python value types:
strings remain text, including numeric-looking identifiers and formula-looking
strings. `set_number_format("A:C", "0.00")` and
`set_column_width("A:C", 24)` affect the current sheet's entire columns.
A failing row iterator rolls back that write call; earlier rows remain usable.

## ZIP compression

```python
editor.save()                       # compression_level=3 by default
editor.save(compression_level=0)    # store every ZIP part without compression
editor.save("copy.xlsx", compression_level=9)
```

Levels are `0..9`; invalid levels fail before changing the destination.
Every ZIP part uses the requested level, including unchanged parts of a template;
the former small-XML threshold is removed. Compression may change file size and
saving time. On `WorkbookWriter` choose the level in its constructor, because
worksheets are compressed before the final `save()` call.

Editor snapshots share their XML buffers, caches, dimension metadata, and the
original open ZIP source. Each branch has an independent ZIP cursor. Atomic
saving by one branch does not replace the source read by the other branches.
An open editor can still save if its original path has been removed.


## Parallel exports and ZIP compression (0.15.1)

Independent workbooks can be written from Python threads. The adapter selects
value converters once per dataframe dtype and hands bounded batches to Rust;
XML generation, row spooling, float formatting, and saving run outside the GIL.
Python iteration/conversion still uses the GIL, so scaling is workload-dependent.
The package still imports neither Polars nor Arrow.

```python
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from excelsior import WorkbookWriter

MAX_ROWS = 1_000_000
output = Path("out")

def export(item):
    name, frame = item
    with WorkbookWriter(output / f"{name}.xlsx", "Index",
                        compression_level=1, compression_threads=2) as wb:
        for i, offset in enumerate(range(0, max(1, frame.height), MAX_ROWS), 1):
            if i > 1:
                wb.add_worksheet(f"Sheet{i}")
            wb.with_polars(frame.slice(offset, MAX_ROWS))

with ThreadPoolExecutor(max_workers=4) as pool:
    list(pool.map(export, files_list))
```

Each worker owns its workbook; dataframes are only read. WorkbookWriter keeps
worksheet XML out of RAM, which is useful when several large files are written
at once. No multiprocessing/pickling or additional dataframe copies are needed.
Its compression workers can close previous sheets while subsequent rows are
spooled. Exceptions from `pool.map` propagate to the caller.

Editor can also be used in the worker, retaining its normal mutation contracts:

```python
wb = Editor.create("out/report.xlsx", "Index")
# wb.with_polars(...); wb = wb.add_worksheet(...)
wb.save(compression_level=1, compression_threads=4)
```

`compression_threads` must be positive and defaults to 1. Each worksheet is a
separate ZIP member, so several sheets can compress concurrently. One member's
Deflate stream is not split across threads; a one-sheet workbook generally
benefits from exporting independent files in parallel instead. Combining many
file workers and compression workers does not guarantee extra speed.
Parallel compression uses a bounded queue and temporary files beside the output,
without buffering full compressed worksheets in RAM. Atomic replacement and
cleanup on success/error/abort apply to both sequential and parallel modes.

`Editor.create` is a classmethod: calling `origin.create(...)` creates a new
book rather than copying `origin` or inheriting its sheet name. Pass `"Index"`
explicitly as shown above; a separate blank.xlsx is unnecessary for this export.
