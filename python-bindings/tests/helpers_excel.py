from __future__ import annotations

import shutil
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


def assert_with_libreoffice(path: Path) -> None:
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if executable is None:
        # The openpyxl and ZIP assertions still run without an office install.
        return
    path = Path(path).resolve()
    with tempfile.TemporaryDirectory(prefix="excelsior-lo-") as directory:
        root = Path(directory)
        proc = subprocess.run(
            [
                executable,
                "--headless",
                "--nologo",
                "--nodefault",
                "--nolockcheck",
                "--nofirststartwizard",
                f"-env:UserInstallation={(root / 'profile').as_uri()}",
                "--convert-to",
                "csv",
                "--outdir",
                str(root),
                str(path),
            ],
            check=False,
            capture_output=True,
            text=True,
            timeout=30,
        )
        assert (root / path.with_suffix(".csv").name).exists(), (
            f"LibreOffice produced no CSV (exit {proc.returncode})\n"
            f"stdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
        )


def get_sheet(path: Path, sheet_name: str):
    wb = load_workbook(path)
    return wb[sheet_name]


def cell_text(v: Any) -> str | None:
    if v is None:
        return None
    return str(v)


def column_width_map(path: Path, sheet_index: int = 1) -> dict[str, float]:
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    out: dict[str, float] = {}
    with zipfile.ZipFile(path) as zf:
        xml = zf.read(f"xl/worksheets/sheet{sheet_index}.xml")
    root = ET.fromstring(xml)
    cols = root.find("x:cols", ns)
    if cols is None:
        return out
    for col in cols.findall("x:col", ns):
        min_idx = int(col.attrib["min"])
        max_idx = int(col.attrib["max"])
        width = float(col.attrib["width"])
        for idx in range(min_idx, max_idx + 1):
            out[index_to_excel_col(idx - 1)] = width
    return out


def index_to_excel_col(idx: int) -> str:
    col = ""
    idx += 1
    while idx > 0:
        rem = (idx - 1) % 26
        col = chr(ord("A") + rem) + col
        idx = (idx - 1) // 26
    return col
