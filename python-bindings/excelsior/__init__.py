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

__all__ = [
    "Scanner", "Editor", "WorkbookWriter", "AlignSpec", "HorizAlignment", "VertAlignment",
    "scan_excel", "create_excel", "__version__",
]
