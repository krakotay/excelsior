# Changelog

## 0.14.0

- Build and test CPython 3.10–3.15 wheels for Linux, Windows, and macOS;
  update PyO3 for Python 3.15, and validate release versions before building.
- Move Polars integration entirely into Python using `columns` and `iter_rows`;
  remove Arrow/native Polars coupling and the `polars==1.35.*` runtime pin.
- Make Python dependencies optional, keeping XLSX creation/editing standalone.
- Export XLSX creation from the Rust crate root and document `create_excel`
  and `Editor.create` in the Python package root.
- Include typing markers/stubs and expose the loaded package version.
- Run Python tests without a required LibreOffice installation, and publish
  only on manual release requests or version tags.
- Make both justfiles portable through uv-managed environments without shell activation.
