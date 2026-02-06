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

For detailed usage examples see [docs/usage.md](docs/usage.md).
