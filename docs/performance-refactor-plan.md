# Performance Refactor Plan

This document describes a practical refactor plan for the main performance
and architecture bottlenecks in `excelsior`.

## Progress

- [x] Phase 1 scope defined.
- [x] Rust benchmark harness added.
- [x] Baseline benchmark report recorded.
- [ ] Worksheet bulk-write refactor started.
- [x] Styles cache refactor started.
- [ ] Save/sheet clone reduction started.
- [ ] Workbook metadata cache started.
- [ ] Python streaming path started.

## Recent Progress

- 2026-04-06: Added `rust-core/src/bin/perf_bench.rs` for repeatable local
  performance checks.
- 2026-04-06: Recorded first release baseline in
  `docs/performance-baseline.md`.
- 2026-04-06: Switched `read_style_parts` to use `StyleIndex.xfs` cache instead
  of reparsing `styles.xml` on every style merge.
- 2026-04-06: Rebuilt benchmark into a stress tool with streaming row
  generation, million-row presets, hardcore style waves, persistent `.xlsx`
  outputs, and auto-written markdown reports under `test/perf_outputs`.

## Goals

- Remove repeated full scans of worksheet XML for common write operations.
- Remove repeated full parses of `styles.xml`.
- Reduce cloning of large XML buffers during sheet switching and save.
- Avoid building full in-memory tables in Python when streaming data would do.
- Keep the public API stable where possible.

## Guiding Rule

Current design strong like rock for small files, but for large sheets it do
same work many times. Refactor should move from:

- "search whole XML again"
- "parse whole styles.xml again"
- "clone big buffer again"

toward:

- "scan once"
- "cache structure"
- "write once"

## Phase 1: Measure First

Before changing internals, add repeatable benchmarks and profiling scenarios.

### Add benchmark cases

- Append 10k, 100k, 1M cells to an empty sheet.
- Insert table into existing populated sheet with overlap.
- Apply style to a large rectangular range.
- Apply number format to many columns after `with_polars`.
- Switch between multiple sheets many times, then save.

### Measure

- Total wall time.
- Peak memory.
- Time split for open, mutate, save.
- Time split for style-heavy vs data-heavy operations.

### Deliverable

- Bench harness in Rust.
- Simple benchmark script for Python bindings.
- Baseline markdown report checked into `docs`.

## Phase 2: Stop Re-Scanning Worksheet XML

Biggest hot path now is repeated linear search over `sheet_xml`.

### Problem

- `append_table_at` calls `set_cell` for overlapping rows.
- `set_cell` searches row and cell again for each write.
- Single-cell style application also searches by string pattern again.

### Refactor

Build one-pass worksheet patch engine for data writes, similar in spirit to the
better style patch path already present.

### Steps

1. Add a row-oriented patch path for `append_table_at`.
2. For overlapping writes, scan `sheet_xml` once and patch rows in streaming
   fashion.
3. Insert missing rows and missing cells during that same pass.
4. Keep `set_cell` as convenience API, but make bulk APIs avoid it internally.

### Target outcome

- `append_table_at` becomes near-linear in sheet size plus inserted data size.
- No per-cell full-buffer search for bulk operations.

## Phase 3: Make Styles Cache Actually Complete

Styles subsystem already has a cache, but cache not deep enough. Code still
re-parses `styles.xml` many times.

### Problem

- `StyleIndex::build` does several full passes.
- `read_style_parts` triggers more full passes through `styles.xml`.
- `xf_components`, `xf_border`, `xf_alignment` all walk same XML again.

### Refactor

Expand `StyleIndex` into the single source of truth for parsed style metadata.

### Steps

1. During `StyleIndex::build`, parse `styles.xml` once or at worst one logical
   pass per top-level block, not many disconnected passes.
2. Store full `XfParts` in cache for every style id.
3. Make `read_style_parts` read from cached `xfs` instead of reparsing XML.
4. Make `ensure_style_from_parts` and style merge logic use cache only.
5. Invalidate or incrementally update cache only when style XML changes.

### Target outcome

- Style reads become O(1) after cache build.
- Style-heavy operations stop paying repeated XML parse tax.

## Phase 4: Kill Big Buffer Clones

Large `Vec<u8>` clones cost both time and memory.

### Problem

- Current sheet flushed by cloning whole `sheet_xml`.
- Save path clones changed file contents before writing.
- Sheet switching clones buffers from caches.
- Save opens source ZIP more than once.

### Refactor

Change ownership model so sheet buffers move when possible and borrow when safe.

### Steps

1. Replace clone-based `flush_current_sheet` with move-based storage.
2. Keep current active sheet separate, but move old sheet buffer into
   `new_files` on switch.
3. In save path, avoid `content.clone()` when only dimension update is needed.
   Prefer:
   - mutate temporary only for sheet outputs that truly need rewrite
   - or write from borrowed content directly
4. Read `[Content_Types].xml` from already-open archive instead of opening ZIP
   second time.
5. Consider storing `loaded_files` and `new_files` behind `Arc<[u8]>` or
   move-only `Vec<u8>` ownership if that simplifies switching.

### Target outcome

- Lower peak memory on large workbooks.
- Less allocator churn.

## Phase 5: Stop Reading Workbook State From Disk Again

Some operations still ask original file for workbook state even though editor
already holds workbook metadata in memory.

### Problem

- `open` does `scan`, then `open_sheet`.
- worksheet add/rename validation re-reads workbook from disk.

### Refactor

Keep workbook-level metadata cached in memory.

### Steps

1. Add parsed workbook metadata struct:
   - sheet name to rid
   - rid to target path
   - sheet order
   - max sheet file number
2. Build it once on open.
3. Update it on add/rename/delete/switch operations.
4. Replace `scan(&self.src_path)` calls inside editor methods with metadata
   lookup from memory.

### Target outcome

- Fewer ZIP opens.
- Less duplicated workbook parsing.
- Cleaner architecture for sheet management.

## Phase 6: Stream Python `with_polars`

Python path now duplicates data too much.

### Problem

- `with_polars` builds full `Vec<Vec<String>>`.
- Then Rust writes whole table.
- Then code loops columns again to set formats.

### Refactor

Move to chunked or streaming transfer from Python to Rust.

### Steps

1. Add chunked row ingestion API in Rust core.
2. In Python binding, iterate `iter_rows()` in chunks instead of collecting all.
3. Detect column format while streaming first chunk set or whole stream.
4. Batch style/format application by contiguous columns with same format.
5. If possible, add a fast path for primitive Polars columns rather than Python
   object per cell conversion.

### Target outcome

- Lower peak memory.
- Better throughput for large DataFrames.
- Less Python/Rust boundary overhead.

## Phase 7: Clean API and Internal Duplication

There are several repeated helper implementations and duplicated logic.

### Work items

- Centralize column name/index conversion helpers.
- Centralize XML attribute parsing helpers on byte slices.
- Remove duplicate `add_new_xf` style-building paths if both old and new code
  still exist.
- Remove `println!` debug output from open path.
- Audit old fallback code paths that silently reparse whole XML.

### Target outcome

- Smaller maintenance surface.
- Fewer accidental slow paths.

## Implementation Order

Best order for good speed gain with low risk:

1. Add benchmarks and baseline reports.
2. Refactor styles cache.
3. Refactor `append_table_at` bulk write path.
4. Reduce sheet/save clones.
5. Add workbook metadata cache.
6. Stream `with_polars`.
7. Cleanup and dedup.

## Risk Notes

- XLSX is picky. Any refactor that touches XML writing needs regression tests
  against Excel-openable files.
- Style refactor can cause subtle corruption if `count` attributes or ids drift.
- Streaming writes must preserve row order and valid sheet dimension updates.

## Test Plan

- Keep existing unit and Python tests.
- Add golden-file tests for:
  - append into empty sheet
  - append into populated sheet
  - style merge on existing styled cells
  - worksheet add/rename/delete/switch
  - save and reopen in Excel-compatible reader
- Add large-file stress tests for save and style operations.

## Success Criteria

Refactor done when:

- Large `append_table_at` workloads no longer degrade sharply with overlap.
- Style-heavy operations stop doing repeated full XML parses.
- Peak memory during save materially drops.
- `with_polars` no longer needs full duplicate table in memory.
- Bench reports show clear win on both Rust core and Python path.
