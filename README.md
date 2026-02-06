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

## Example
```python
scanner = Scanner(file_path)
editor = scanner.open_editor(scanner.get_sheets()[0])

editor.append_table_at([[str(k) for k in list(range(50))] for _k in list(range(5))], "B4")
editor.save(out_path)

```

![Before](image_before.png)
![After](image_after.png)
