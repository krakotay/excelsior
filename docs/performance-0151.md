# 0.15.1: parallel OSV dataframe exports

Measured on Linux x86_64 with 16 logical CPUs, CPython 3.14.6, Polars 2.1.0,
Rust 1.99.1 release builds, and disk-backed ext4 temporary storage.
The baseline is commit `a0bcbfc`, version 0.15.0.

## Real input and validation

The CSV contains 3,919,358 rows and 18 columns. Applying the notebook's footer,
currency, identifier cleanup, numeric casts and partition filters produces:

| Frame | Data rows |
| --- | ---: |
| Zero | 759,743 |
| Small | 2,342,870 |
| Working | 816,738 |
| All | 3,919,351 |

The four files therefore write 7,838,702 data rows (141,096,636 cells), plus
headers, across 1 / 3 / 1 / 4 worksheets using 1,000,000-row chunks.
The deprecated Polars `is_in(Series)` call was reproduced with `.implode()`
without changing the resulting counts. The actual CSV and existing output
files were only read. Benchmark XLSX files used a separate temporary directory
and were removed after verification; no financial row values are logged.

Both full runs streamed every saved ZIP member to EOF to verify decompression
and CRC, counted every worksheet's row tags, and compared all last-row values
and their stored text/number/bool types against the input slice. This is not a
cell-by-cell comparison of the whole workbook or a visual Microsoft Excel test.

## Full export results

| Workflow | Wall time (s) | Process CPU time (s) | Peak RSS (MiB) |
| --- | ---: | ---: | ---: |
| 0.15.0 Editor, sequential files, compression level 1 | 138.92 | 138.96 | 6,162.5 |
| 0.15.1 WorkbookWriter, 4 file threads × 2 compression threads, level 1 | 48.75 | 129.45 | 4,928.8 |

This measured workflow is 2.85× faster. CPU/wall increased from about 1.00 to
2.65, confirming that native work actually ran concurrently. These are single
full-size runs on this machine, not statistical guarantees. The comparison
combines the writer choice, batching, buffering, dtype conversion and parallelism;
it does not isolate one change. It also does not promise that the unchanged
sequential Editor loop alone becomes 2.85× faster.

Times include workbook creation, all row writes, compression, flushing and
atomic saving. CSV preparation and post-save verification are excluded from
export time. Peak RSS includes input dataframes, their preparation, Python and
Polars, so it is not the writer's incremental memory usage. Four parallel
Editors would retain more worksheet XML; WorkbookWriter is the preferred fit
for these large new files.

Profiling the 0.15.0 adapter showed repeated Python type checks/conversions as
the main hot path. Saving/compression was a smaller part of the original
sequential workflow. Profiling overhead is not used as an export timing result.

## Implementation

- Select common value converters once per dtype instead of repeating all scalar
  type checks for every cell. Keep generic conversion for dates/objects and
  unusual Python values; finite floats are formatted in Rust with a round-trip
  representation that preserves extreme exponents and signed zero.
- Fetch bounded batches (at most 256 rows, stopping after 16,384 cells) while
  attached to Python. Generate XML, spool rows and format floats outside the GIL.
  Python iteration/conversion still takes the GIL, so scaling is not unlimited.
- Add `compression_threads=1` to Editor.save and WorkbookWriter construction.
  A larger value compresses independent ZIP members/sheets concurrently into
  temporary mini archives; compressed bytes are copied into the final ZIP
  without another compression pass. Queue size is bounded by the thread count.
- Buffer streaming worksheet XML before calling the compressor, avoiding tiny
  compression writes for each cell.
- Keep iterator rollback and original exception order across batches. Worker
  errors/abort join all outstanding jobs and remove temporary files. Successful
  saves atomically replace the destination.

The Polars adapter remains Python-only: no native Polars/Arrow dependency was
added. Each Python file worker owns a separate Editor/WorkbookWriter and only
reads its input dataframe. A single worksheet's Deflate stream remains serial;
several sheets or separate files provide the independent work for threads.
Compression threads are opt-in and default to 1.

## Reproduce

Use a release build and run each workflow in its own process, without other
heavy benchmarks running concurrently. The CSV must use the OSV column schema
from the notebook. No private data files or machine-specific paths are bundled.
The benchmark deletes generated XLSX files after collecting results.

```sh
python python-bindings/speed_tests/benchmark_osv_parallel.py --csv /path/to/OSV.csv --rows 0 --mode editor --workers 1 --verify --result before.json
python python-bindings/speed_tests/benchmark_osv_parallel.py --csv /path/to/OSV.csv --rows 0 --mode stream --workers 4 --compression-threads 2 --verify --result after.json
```

Run the first command with the 0.15.0 wheel installed and the second with
0.15.1. `--rows 200000` selects a smaller real-data workload. `--work-dir` can
select a disk-backed filesystem if the system temp directory is tmpfs.
The JSON records the loaded package path, version, workflow options and timings.
