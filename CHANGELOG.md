# Changelog

## 0.15.1

- Convert dataframe values with converters selected once per dtype; pass finite
  floats as numbers and format them in Rust without changing their cell types.
- Process bounded row batches with native XML/spool work outside the GIL so
  independent workbooks can be exported from Python threads.
- Add optional `compression_threads` (default 1) to Editor.save and
  WorkbookWriter construction. Compress independent ZIP members concurrently
  using temporary files and copy compressed bytes into the final ZIP in order.
- Buffer streaming worksheet output before compression to avoid tiny writes
  to the compressor for every cell.
- Preserve whole-call iterator rollback and exception order across batches;
  join workers and remove their temporary files on save, error, and abort.

## 0.15.0

- Share Editor XML, caches, and dimensions with Arc/copy-on-write; retain an open
  original ZIP with independent cursors so branches survive atomic saves.
- Add `WorkbookWriter` for bounded-memory creation using a typed disk spool,
  direct ZIP output, iterator rollback, context management, and atomic save.
- Replace byte-pattern XML editing with one namespace-aware span parser,
  preserving unknown extensions and handling quote/prefix/CDATA variations.
- Add `save(compression_level=3)` with levels 0..9 for all ZIP entries;
  configure streaming compression at writer construction.
- Cache parsed row spans for repeated scalar edits, invalidating offsets after
  structural changes and keeping branch caches independent.
- Cache generated sheet dimensions and assign initial column styles while
  writing dataframe rows, avoiding subsequent whole-sheet format passes.

- Allow `save()` to use the original create/open path; `save(other_path)` writes a
  copy without rebinding the editor. Remember absolute paths across cwd changes.

- Breaking Python API: methods returning Editor now return independent copies;
  methods returning None mutate. Update fluent calls to assign their result.
- Preserve dataframe cell types, including long numeric-looking strings and
  formula-looking text, without importing Polars or converting through Arrow.
- Force whole-column number formats from dataframe dtypes by default;
  `force_column_formats=False` preserves template formats.
- Keep repeated appends efficient by inserting new row XML at the tail.
- Stream rows into one XML merge instead of materializing multiple full tables;
  apply all dataframe column formats in one cell pass.
- Handle NaN/infinity as text, native booleans, decimal scale, time/duration,
  and the workbook's 1904 date system.
- Fix columns beyond Z, coordinate/selector validation, failed sheet selection,
  self-closing rows, and use atomic replacement when saving workbooks.
- Adapt XML name/value handling to quick-xml 0.42, including normalized escaped attributes.
- Add a portable source-distribution recipe and document local uv installation.

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
