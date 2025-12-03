import win32com.client as win32
from pathlib import Path
from typing import Callable, Iterable, List, Sequence, Tuple, Union
import shutil
import tempfile
import os
import sys

# Константы Excel (чтобы не тянуть win32com.client.constants)
XL_CENTER = -4108  # xlCenter
XL_THIN = 2  # xlThin
XL_EDGE_LEFT = 7  # xlEdgeLeft
XL_EDGE_TOP = 8  # xlEdgeTop
XL_EDGE_BOTTOM = 9  # xlEdgeBottom
XL_EDGE_RIGHT = 10  # xlEdgeRight

try:
    temp_dir = os.environ.get("TEMP") or tempfile.gettempdir()
    target = Path(temp_dir) / "gen_py"
    if target.exists():
        shutil.rmtree(target, ignore_errors=True)
except Exception as e:
    # Silently ignore cleanup errors
    print(e)
    pass


def open_excel():
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # чтобы не маячил
    excel.DisplayAlerts = False
    return excel


def check_font_range(ws, addr, name, size, bold, italic, horiz=None):
    rng = ws.Range(addr)
    errors = []

    for cell in rng:
        font = cell.Font
        if name is not None and font.Name != name:
            errors.append(
                f"{addr} [{cell.Address}] Font.Name: '{font.Name}' != '{name}'"
            )
        if size is not None and abs(font.Size - size) > 0.01:
            errors.append(f"{addr} [{cell.Address}] Font.Size: {font.Size} != {size}")
        if bold is not None and bool(font.Bold) != bool(bold):
            errors.append(f"{addr} [{cell.Address}] Font.Bold: {font.Bold} != {bold}")
        if italic is not None and bool(font.Italic) != bool(italic):
            errors.append(
                f"{addr} [{cell.Address}] Font.Italic: {font.Italic} != {italic}"
            )
        if horiz is not None and cell.HorizontalAlignment != horiz:
            errors.append(
                f"{addr} [{cell.Address}] HorizontalAlignment: "
                f"{cell.HorizontalAlignment} != {horiz} (xlCenter)"
            )

    return errors


def check_font_not_range(ws, addr, name, size, bold, italic, horiz=None):
    """Проверяет, что ни одна ячейка диапазона не имеет указанных стилей.

    Используем для проверки, что стили были успешно сняты: если хотя бы одна
    ячейка точно совпадает по всем переданным параметрам, это считается ошибкой.
    """
    rng = ws.Range(addr)
    errors = []

    for cell in rng:
        font = cell.Font
        same_name = name is None or font.Name == name
        same_size = size is None or abs(font.Size - size) <= 0.01
        same_bold = bold is None or bool(font.Bold) == bool(bold)
        same_italic = italic is None or bool(font.Italic) == bool(italic)
        same_horiz = horiz is None or cell.HorizontalAlignment == horiz

        if same_name and same_size and same_bold and same_italic and same_horiz:
            errors.append(
                f"{addr} [{cell.Address}] still has style "
                f"name={font.Name!r}, size={font.Size}, bold={font.Bold}, "
                f"italic={font.Italic}, horiz={cell.HorizontalAlignment}"
            )

    return errors


def check_borders_thin(ws, addr):
    rng = ws.Range(addr)
    errors = []

    # Проверим по границам диапазона
    # Если стили заданы равномерно, этого достаточно
    borders = rng.Borders
    for edge, name in [
        (XL_EDGE_LEFT, "Left"),
        (XL_EDGE_TOP, "Top"),
        (XL_EDGE_BOTTOM, "Bottom"),
        (XL_EDGE_RIGHT, "Right"),
    ]:
        b = borders(edge)
        if b.Weight != XL_THIN:
            errors.append(
                f"{addr} Border {name}: Weight {b.Weight} != {XL_THIN} (xlThin)"
            )
        # Можно дополнительно проверить LineStyle (например, что не xlLineStyleNone)
        if b.LineStyle == 0:
            errors.append(f"{addr} Border {name}: LineStyle == 0 (похоже, границы нет)")

    return errors


def safe_merge_area(rng):
    try:
        return rng.MergeArea.Address
    except Exception:
        return None


def check_merge(ws, addr, expected_addr=None):
    rng = ws.Range(addr)
    errors = []

    if not rng.MergeCells:
        errors.append(f"{addr}: диапазон не объединён")
        return errors

    try:
        real_addr = rng.MergeArea.Address
    except Exception as e:
        errors.append(f"{addr}: Excel COM взорвался при доступе к MergeArea: {e}")
        return errors

    if expected_addr and real_addr != expected_addr:
        errors.append(
            f"{addr}: MergeArea.Address = {real_addr}, ожидается {expected_addr}"
        )

    return errors


def _normalize_range_value(value):
    if isinstance(value, tuple):
        return [list(row) for row in value]
    return [[value]]


def check_range_values(ws, addr, expected: Sequence[Sequence[object]]):
    rng = ws.Range(addr)
    actual = _normalize_range_value(rng.Value)
    expected_list = [list(row) for row in expected]

    def _normalize_scalar(v):
        if v is None:
            return None
        if isinstance(v, float):
            if v.is_integer():
                v = int(v)
        return str(v)

    normalized_actual = [[_normalize_scalar(v) for v in row] for row in actual]
    normalized_expected = [[_normalize_scalar(v) for v in row] for row in expected_list]

    if normalized_actual != normalized_expected:
        return [f"{addr}: значения {actual} != {expected_list}"]
    return []


CheckFunc = Callable[[object], Iterable[str]]
SheetSelector = Union[int, str]
CheckSpec = Union[CheckFunc, Tuple[SheetSelector, CheckFunc]]


def on_sheet(sheet: SheetSelector, check: CheckFunc) -> Tuple[SheetSelector, CheckFunc]:
    """Helper to explicitly bind a check to a worksheet."""

    return (sheet, check)


def _resolve_check_spec(spec: CheckSpec, default_sheet: SheetSelector) -> Tuple[SheetSelector, CheckFunc]:
    if callable(spec):
        return default_sheet, spec
    if isinstance(spec, tuple) and len(spec) == 2:
        sheet, check = spec
        if not callable(check):
            raise TypeError("Check must be callable")
        return sheet, check
    raise TypeError(
        "Check specification must be a callable or a tuple of (sheet, callable)"
    )


def run_excel_checks(
    out_filename: Union[str, Path], *check_specs: CheckSpec, sheet: SheetSelector = 1
) -> List[str]:
    """Open a workbook in real Excel and run a set of validation callbacks.

    Each callback receives the worksheet object and should return an iterable of
    error strings (or None/empty iterable if everything is fine).
    """

    if not check_specs:
        raise ValueError("Нужно передать хотя бы одну функцию проверки")

    file_path = Path(out_filename).resolve()
    if not file_path.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    excel = open_excel()
    wb = None
    try:
        wb = excel.Workbooks.Open(str(file_path))
        errors: List[str] = []
        for spec in check_specs:
            target_sheet, check = _resolve_check_spec(spec, sheet)
            ws = wb.Worksheets(target_sheet)
            result = check(ws)
            if result:
                errors.extend(result)
        return errors
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.Quit()


def assert_excel_file(
    out_filename: Union[str, Path], *check_specs: CheckSpec, sheet: SheetSelector = 1
) -> None:
    """PyTest-friendly helper that fails fast when Excel validation finds errors."""

    errors = run_excel_checks(out_filename, *check_specs, sheet=sheet)
    if errors:
        formatted = "\n".join(f" - {err}" for err in errors)
        raise AssertionError(
            f"Excel validation failed:\n{formatted}\nFilename {out_filename}"
        )


def main():
    base_dir = Path(__file__).resolve().parent

    # Можно поменять путь при необходимости
    out_filename = (base_dir / "../../test/style_test_out_py.xlsx").resolve()

    try:
        errors = run_excel_checks(
            out_filename,
            lambda ws: check_font_range(ws, "D4:D7", "Arial", 12.0, True, False),
            lambda ws: check_font_range(
                ws,
                "A1:C2",
                "Calibri",
                10.0,
                False,
                True,
                horiz=XL_CENTER,
            ),
            lambda ws: check_borders_thin(ws, "A1:C2"),
            lambda ws: check_merge(ws, "B12:D12", expected_addr="$B$12:$D$12"),
        )
    except FileNotFoundError as exc:
        print(exc)
        sys.exit(1)

    if errors:
        print("Проверка НЕ пройдена. Найдены несоответствия:")
        for err in errors:
            print(" -", err)
        sys.exit(1)

    print("OK: все проверенные стили соответствуют ожиданиям.")
    sys.exit(0)


if __name__ == "__main__":
    main()
