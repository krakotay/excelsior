# Performance and architecture work in 0.15

Implemented:

- [x] Row-wise dataframe ingestion without a full Python string table.
- [x] A single namespace-aware OOXML span parser based on quick-xml 0.42.
- [x] Attribute order, quote styles, XML comments, CDATA, and empty elements handled by the parser.
- [x] Shared XML buffers and workbook caches with Arc/copy-on-write.
- [x] Independent cursors over the shared original ZIP file for Editor branches.
- [x] Borrowed sheet output during save; dimension changes write slices instead of cloning entire XML.
- [x] Cached dimensions for generated/validated sheets.
- [x] A lazy row-span index for repeated scalar replacements, shared with copy-on-write.
- [x] Column defaults established before writing a new dataframe sheet.
- [x] WorkbookWriter with a temporary typed row spool and direct ZIP output for completed sheets.
- [x] Configurable ZIP compression, default level 3, levels 0..9.
- [x] GIL released during saving and streaming sheet finalization.

Still separate work:

- General batch patches and indexes for sparse insertions or many unrelated edits.
- A workbook/worksheet-view API if the public ownership model needs further simplification.
- More application-produced OOXML fixtures and visual checks in Microsoft Excel.

The streaming writer holds no worksheet XML, but requires temporary disk space.
Editor remains an in-memory XML editor; copy-on-write removes eager snapshot
copies without making arbitrary modifications free. Performance comparisons
must use matching input, compression settings, and Python/Rust builds.

See [the 0.15 measurements](performance-015.md) for timings, memory use, and reproducible commands.
