# Project overview

`excelsior` is organised as a Cargo workspace containing two crates:

- `rust-core` implements the logic for editing XLSX files. It works at the XML
  level inside the ZIP archive without requiring heavy dependencies.
- `python-bindings` exposes the same functionality to Python via the `pyo3`
  ecosystem and is published as the `excelsior` package.

The `rust-core` crate defines the `XlsxEditor` type and helper function
`scan` for listing sheet names. The editor can append rows or tables to a
worksheet, insert a table starting from an arbitrary cell and update individual
cells. Modified content can be saved to a new file or overwrite the
original workbook.

The Python bindings mirror this API in the `Editor` class and a
`scan_excel` function. Example scripts can be found in
[`python-bindings/tests`](../python-bindings/tests).

A sample spreadsheet for tests is located in the `test/` directory.
