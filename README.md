# excelsior
[![Rust](https://github.com/krakotay/excelsior/actions/workflows/rust.yml/badge.svg)](https://github.com/krakotay/excelsior/actions/workflows/rust.yml)
[![Build maturin wheels](https://github.com/krakotay/excelsior/actions/workflows/release.yml/badge.svg?branch=master)](https://github.com/krakotay/excelsior/actions/workflows/release.yml)

[pypi link](https://pypi.org/project/excelsior-fast/)

A small project for quickly updating `.xlsx` workbooks from Rust or Python.
It consists of two crates:

* **rust-core** – the core library that works directly with spreadsheet XML.
* **python-bindings** – Python wrapper built with `pyo3` and `maturin`.

The library lets you append rows or tables, modify individual cells and
save the workbook back to disk without loading the entire file into memory.

