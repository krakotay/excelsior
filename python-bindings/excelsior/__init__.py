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

__all__ = [
    "Scanner", "Editor", "AlignSpec", "HorizAlignment", "VertAlignment",
    "scan_excel", "create_excel", "__version__",
]
