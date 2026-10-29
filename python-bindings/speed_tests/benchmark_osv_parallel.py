"""Local OSV export benchmark. Logs dimensions/timings, never financial rows."""

import argparse, inspect, json, os, resource, tempfile, time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import polars as pl

MONEY = [
    "Вх остаток А",
    "Вх остаток П",
    "Оборот ДТ",
    "Оборот КТ",
    "Исх. остаток А",
    "Исх. остаток П",
]


def frames(source, limit=None):
    df = pl.read_csv(source, separator=";", infer_schema_length=0)
    df = df.slice(0, -3).filter(pl.col("Валюта").is_null())
    if limit is not None:
        df = df.head(limit)
    df = (
        df.with_columns(
            pl.col("Лицевой счет", "Сводный счет", "Инн клиента")
            .str.slice(2)
            .str.strip_suffix('"')
        )
        .with_columns(pl.all().str.replace_all(",", "", literal=True))
        .with_columns(pl.col(MONEY).cast(pl.Float64))
    )
    zero = df.filter(pl.all_horizontal(pl.col(MONEY) == 0.0))
    small = df.filter(
        pl.any_horizontal(pl.col(MONEY) > 0)
        & pl.all_horizontal(pl.col(df.columns[6:12]) < 50_000)
    )
    big = df.filter(
        ~pl.col("Номер строки").is_in(
            pl.concat([zero, small])["Номер строки"].implode()
        )
    )
    return [
        ("нулевой", zero),
        ("малозначительный", small),
        ("рабочий", big),
        ("все данные", df),
    ]


import excelsior

p = argparse.ArgumentParser()
p.add_argument("--csv", type=Path, required=True)
p.add_argument("--work-dir", type=Path, default=Path(tempfile.gettempdir()))
p.add_argument("--result", type=Path)
p.add_argument("--rows", type=int, default=200_000)
p.add_argument("--mode", choices=["editor", "stream"], default="editor")
p.add_argument("--workers", type=int, default=1)
p.add_argument("--compression-threads", type=int, default=1)
p.add_argument("--sheet-rows", type=int, default=1_000_000)
p.add_argument("--verify", action="store_true")
p.add_argument("--label", default="current")
a = p.parse_args()
root = a.work_dir
root.mkdir(parents=True, exist_ok=True)
start = time.perf_counter()
jobs = frames(a.csv, a.rows if a.rows else None)
prepared = time.perf_counter() - start
print(
    json.dumps(
        {
            "event": "prepared",
            "label": a.label,
            "seconds": prepared,
            "shapes": [(name, df.shape) for name, df in jobs],
        },
        ensure_ascii=False,
    ),
    flush=True,
)


def check(path, df):
    with ZipFile(path) as z:
        sheets = sorted(
            (n for n in z.namelist() if n.startswith("xl/worksheets/")),
            key=lambda name: int(name.rsplit("sheet", 1)[1].split(".")[0]),
        )
        assert len(sheets) == max(1, (df.height + a.sheet_rows - 1) // a.sheet_rows)
        for i, name in enumerate(sheets):
            expected = df.slice(i * a.sheet_rows, a.sheet_rows)
            count = 0
            carry = b""
            tail = b""
            with z.open(name) as f:
                while block := f.read(1024 * 1024):
                    count += (carry + block).count(b"<row ")
                    carry = block[-4:]
                    tail = (tail + block)[-2 * 1024 * 1024 :]
            assert count == expected.height + 1, (name, count, expected.height)
            begin = tail.rfind(b"<row ")
            end = tail.find(b"</row>", begin) + 6
            row = ET.fromstring(tail[begin:end])
            cells = list(row)
            wanted = expected.row(-1) if expected.height else expected.columns
            assert len(cells) == len(wanted)
            for cell, value in zip(cells, wanted):
                if value is None:
                    assert cell.find("v") is None and cell.find("is") is None
                elif isinstance(value, str):
                    assert (
                        cell.attrib["t"] == "inlineStr"
                        and cell.findtext("is/t") == value
                    )
                elif isinstance(value, bool):
                    assert cell.attrib["t"] == "b" and cell.findtext("v") == str(
                        int(value)
                    )
                else:
                    assert float(cell.findtext("v")) == float(value)
        for name in z.namelist():
            if name not in sheets:
                z.read(name)


def export(item, directory):
    name, df = item
    path = Path(directory) / f"{name}.xlsx"
    started = time.perf_counter()
    if a.mode == "editor":
        wb = excelsior.Editor.create(path)
    else:
        wb = (
            excelsior.WorkbookWriter(
                path, compression_level=1, compression_threads=a.compression_threads
            )
            if hasattr(excelsior, "WorkbookWriter")
            and "compression_threads"
            in inspect.signature(excelsior.WorkbookWriter).parameters
            else excelsior.WorkbookWriter(path, compression_level=1)
        )
    for i, offset in enumerate(range(0, max(1, df.height), a.sheet_rows), 1):
        if i > 1:
            if a.mode == "editor":
                wb = wb.add_worksheet(f"Sheet{i}")
            else:
                wb.add_worksheet(f"Sheet{i}")
        wb.with_polars(df.slice(offset, a.sheet_rows))
    written = time.perf_counter()
    if a.mode == "editor":
        kw = {"compression_level": 1}
        if "compression_threads" in inspect.signature(wb.save).parameters:
            kw["compression_threads"] = a.compression_threads
        wb.save(**kw)
    else:
        wb.save()
    saved = time.perf_counter()
    record = {
        "name": name,
        "rows": df.height,
        "cols": df.width,
        "write_s": written - started,
        "save_s": saved - written,
        "bytes": path.stat().st_size,
    }
    print(json.dumps({"event": "exported", **record}, ensure_ascii=False), flush=True)
    return record, path, df


with tempfile.TemporaryDirectory(prefix="osv-", dir=root) as directory:
    before = resource.getrusage(resource.RUSAGE_SELF)
    begin = time.perf_counter()
    if a.workers == 1:
        exported = [export(job, directory) for job in jobs]
    else:
        with ThreadPoolExecutor(max_workers=a.workers) as pool:
            exported = list(pool.map(lambda job: export(job, directory), jobs))
    wall = time.perf_counter() - begin
    after = resource.getrusage(resource.RUSAGE_SELF)
    result = {
        "label": a.label,
        "version": excelsior.__version__,
        "package": excelsior.__file__,
        "mode": a.mode,
        "file_workers": a.workers,
        "compression_threads": a.compression_threads,
        "sheet_rows": a.sheet_rows,
        "total_data_rows": sum(df.height for _, df in jobs),
        "wall_s": wall,
        "cpu_s": after.ru_utime + after.ru_stime - before.ru_utime - before.ru_stime,
        "peak_rss_mib": after.ru_maxrss / 1024,
        "preparation_s": prepared,
        "files": [x[0] for x in exported],
    }
    out = a.result
    if out is not None:
        out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    if a.verify:
        for _, path, df in exported:
            check(path, df)
        result["verified_crc_row_counts_last_values"] = True
        if out is not None:
            out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"event": "result", **result}, ensure_ascii=False), flush=True)
