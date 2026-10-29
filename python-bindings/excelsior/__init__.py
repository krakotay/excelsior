"""Fast XLSX creation and editing. Polars is an optional Python integration."""

from .excelsior import (
    AlignSpec,
    HorizAlignment,
    VertAlignment,
    __version__,
    create_excel,
    scan_excel,
)
from ._editor import Editor, Scanner
from ._writer import WorkbookWriter
from . import multithreading

__all__ = [
    "Scanner",
    "Editor",
    "WorkbookWriter",
    "multithreading",
    "AlignSpec",
    "HorizAlignment",
    "VertAlignment",
    "scan_excel",
    "create_excel",
    "__version__",
]
