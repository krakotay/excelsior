"""Small local workloads for the architecture review; no timing assertions."""

import json
from pathlib import Path
from time import perf_counter
import tempfile

import polars as pl
from excelsior import Editor, __version__

results = {"version": __version__, "updates": [], "saves": [], "copies": []}
with tempfile.TemporaryDirectory() as directory:
    root = Path(directory)
    for size in (1000, 2000, 4000):
        editor = Editor.create(root / "updates.xlsx")
        editor.append_table_at([[str(i), "payload"] for i in range(size)], "A1")
        begin = perf_counter()
        for row in range(1, size + 1):
            editor.set_cell(f"A{row}", str(row + 1))
        individual = perf_counter() - begin
        begin = perf_counter()
        editor.append_table_at([[str(i + 2)] for i in range(size)], "A1")
        batch = perf_counter() - begin
        results["updates"].append(
            {
                "rows": size,
                "set_cell_loop_seconds": individual,
                "one_table_seconds": batch,
            }
        )

    frame = pl.DataFrame(
        {
            "id": [f"{i:020d}" for i in range(30_000)],
            "n": range(30_000),
            "value": range(30_000),
        }
    )
    for checkpoints in (False, True):
        editor = Editor.create(root / "sheets.xlsx")
        begin = perf_counter()
        save_time = 0.0
        for sheet in range(1, 5):
            if sheet > 1:
                editor = editor.add_worksheet(f"Sheet{sheet}")
            editor.with_polars(frame)
            if checkpoints:
                started = perf_counter()
                editor.save()
                save_time += perf_counter() - started
        if not checkpoints:
            started = perf_counter()
            editor.save()
            save_time += perf_counter() - started
        results["saves"].append(
            {
                "sheets": 4,
                "rows_per_sheet": 30_000,
                "save_after_each_sheet": checkpoints,
                "save_seconds": save_time,
                "total_seconds": perf_counter() - begin,
            }
        )

    for after_save in (False, True):
        editor = Editor.create(root / "copies.xlsx")
        editor.with_polars(
            pl.DataFrame({f"n{col}": range(100_000) for col in range(8)})
        )
        if after_save:
            editor.save()
        begin = perf_counter()
        for _ in range(5):
            branch = editor.copy()
        results["copies"].append(
            {
                "rows": 100_000,
                "columns": 8,
                "copies": 5,
                "after_save": after_save,
                "copy_seconds": perf_counter() - begin,
            }
        )

print(json.dumps(results, indent=2))
