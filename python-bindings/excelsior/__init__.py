from importlib import import_module as _import_module
from .excelsior import Scanner, Editor, AlignSpec, HorizAlignment, VertAlignment, scan_excel

_ext = _import_module(".excelsior", package=__name__)  # бинарник: excelsior.excelsior

# ЯВНЫЕ реэкспорты — чтобы статике, IDE и людям было ясно

__all__ = ["Scanner", "Editor", "AlignSpec", "HorizAlignment", "VertAlignment", "scan_excel"]

del _import_module, _ext
