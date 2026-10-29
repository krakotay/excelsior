# Architecture measurements for 0.15.0

Measured on Linux x86_64, CPython 3.14.6, Polars 2.1.0, Rust 1.99.1,
release builds. Before is commit `59efe57` (the earlier 0.15.0), after is the
copy-on-write/parser/streaming change. These are local measurements, not CI or
production-workbook results.

Input: 100,000 rows × 8 columns, two strings and six integers, including a
20-character zero-padded identifier. Identical column formats are enabled.
Times exclude dataframe preparation and workbook construction, include row
writing and save, and are medians of three fresh sequential processes per case.
Peak RSS includes Python, the input dataframe, and Polars. Files use disk-backed
ext4 temporary storage; no tmpfs was used. Compression is shown explicitly.

| Case | Write + save (s) | Peak RSS (MiB) | XLSX bytes |
| --- | ---: | ---: | ---: |
| Before: Editor, legacy level 1 / small parts stored | 0.998 | 172.1 | 6,148,538 |
| After: Editor, level 1 for all parts | 0.854 | 101.0 | 5,743,542 |
| After: Editor, default level 3 | 0.932 | 100.5 | 4,092,741 |
| After: WorkbookWriter, level 1 | 1.394 | 96.8 | 5,742,285 |
| After: WorkbookWriter, default level 3 | 1.188 | 96.8 | 4,092,767 |

Editor uses roughly 42% less peak memory here. Default level 3 produces a
smaller file and is slightly faster than the previous implementation in this
scenario, but this comparison combines algorithm, serialization, and compression
changes. The streaming writer is slower than the in-memory Editor here: typed
rows are written to disk and later read into ZIP. Its benefit is bounded
worksheet memory, especially across multiple sheets, rather than guaranteed
speed.

Five consecutive `copy()` calls, replacing the previous temporary branch:
median 0.132 s before versus 0.000022 s after. Peak process RSS was 278.5 versus
100.9 MiB. These are cheap snapshots; modifying a shared XML component still
requires allocating/rebuilding that component.

Single-run scaling checks (default level 3):

| Case | Write + save (s) | Peak RSS (MiB) |
| --- | ---: | ---: |
| Editor: 4 sheets × 100,000 rows | 3.327 | 207.0 |
| WorkbookWriter: same dataframe and 4 sheets | 4.438 | 96.8 |
| WorkbookWriter: generator, 100,000 rows, no Polars import | 1.234 | 21.1 |
| WorkbookWriter: generator, 1,000,000 rows, no Polars import | 11.537 | 20.9 |

For both generator cases a streaming Expat read of the saved XLSX verified row
counts and the final string/numeric values after timing and RSS collection.
No hidden spool/archive files remained after successful saves. Temporary disk
space is still needed for typed rows and the compressed archive; the input
DataFrame's own memory is outside the writer's control.

## Reproduction

Use `python-bindings/speed_tests/benchmark_015.py` with the desired installed
release build. Run each case in a fresh Linux process and repeat it at least
three times, without concurrent CPU-heavy work. Select a disk-backed temp
directory explicitly if the system default temp directory is tmpfs.

```sh
python python-bindings/speed_tests/benchmark_015.py --mode editor --compression-level 1 --temp-dir ./bench-tmp
python python-bindings/speed_tests/benchmark_015.py --mode stream --compression-level 3 --temp-dir ./bench-tmp
python python-bindings/speed_tests/benchmark_015.py --mode copy --temp-dir ./bench-tmp
python python-bindings/speed_tests/benchmark_015.py --mode stream --sheets 4 --temp-dir ./bench-tmp
python python-bindings/speed_tests/benchmark_015.py --mode stream --generator --rows 1000000 --verify --temp-dir ./bench-tmp
```

The script detects the earlier save signature when testing an installed
`59efe57` wheel. `--mode stream` requires the new implementation. Both builds
report 0.15.0, so keep their wheels/environments separate and check the reported
package path. Copy time is reported separately from write + save.

## Repeated scalar updates

A final single-process spot check used existing sheets of 1,000 / 2,000 / 4,000
rows with a numeric cell and a text payload per row. It updated every existing
cell in column A, then overwrote the same column with one table. These are
single runs, not medians; use `benchmark_edits_015.py` for reproduction.

| Rows | Before: set_cell loop (s) | After: indexed set_cell loop (s) | After: one table (s) |
| --- | ---: | ---: | ---: |
| 1,000 | 0.051 | 0.007 | 0.004 |
| 2,000 | 0.190 | 0.009 | 0.007 |
| 4,000 | 0.750 | 0.016 | 0.013 |

The lazy row index is built by the common parser and shared with copy-on-write.
Only the selected row is reparsed; replacing equal-length XML does not shift
later offsets. Width/format changes, table writes, and sheet switching invalidate
the index. A branch's edits do not alter another branch's cached offsets.
The general overlapping table path is slower here than the old byte scanner
(4,000 rows: 0.013 versus 0.004 s); the new blank-sheet dataframe path uses a
separate append path. No universal speedup is claimed.

## Remaining limits

- Row offsets still need shifting when replacement length changes; moving XML
  bytes and a cold index build cost O(sheet size). Repeated sparse row insertions
  and large batches of unrelated edits need a more general batch/index design.
- Editor intentionally retains edited worksheet XML. Use WorkbookWriter when
  creating new books sequentially; closed sheets cannot be revisited and saving
  finalizes the writer.
- The common parser handles namespace prefixes, quote styles, BOM, comments,
  CDATA, and unknown extension elements. This is not evidence of complete
  support for every OOXML feature, Strict OOXML, or every third-party template.
- Saving an edited template recompresses unchanged ZIP parts at the requested
  level. This costs CPU on large embedded media; reuse of matching compressed
  entries is a possible later optimization.
