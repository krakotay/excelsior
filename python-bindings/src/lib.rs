use pyo3::exceptions::{PyRuntimeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyIterator, PyTuple};
use rust_core::style::{AlignSpec, HorizAlignment, VertAlignment};
use rust_core::{XlsxEditor, scan};
use std::path::PathBuf;

fn index_to_excel_col(mut idx: usize) -> String {
    let mut col = String::new();
    idx += 1;
    while idx > 0 {
        let rem = (idx - 1) % 26;
        col.insert(0, (b'A' + rem as u8) as char);
        idx = (idx - 1) / 26;
    }
    col
}

fn excel_col_to_index(col: &str) -> PyResult<usize> {
    rust_core::data::parse_column(col.trim()).map_err(|e| PyRuntimeError::new_err(e.to_string()))
}

fn normalize_column_letter(col: &str) -> PyResult<String> {
    excel_col_to_index(col)?;
    Ok(col.trim().to_ascii_uppercase())
}

fn expand_column_selector(selector: &str) -> PyResult<Vec<String>> {
    let trimmed = selector.trim();
    if let Some((start, end)) = trimmed.split_once(':') {
        let start_idx = excel_col_to_index(start)?;
        let end_idx = excel_col_to_index(end)?;
        if start_idx > end_idx {
            return Err(PyRuntimeError::new_err(format!(
                "Invalid column range: {selector}"
            )));
        }
        Ok((start_idx..=end_idx).map(index_to_excel_col).collect())
    } else {
        Ok(vec![normalize_column_letter(trimmed)?])
    }
}

fn validate_width(width: f64) -> PyResult<f64> {
    if !width.is_finite() || width <= 0.0 || width > 255.0 {
        return Err(PyRuntimeError::new_err(
            "Column width must be in range (0, 255]",
        ));
    }
    Ok(width)
}

fn normalize_sheet_name(sheet_name: &str) -> PyResult<String> {
    rust_core::workbook::normalize_sheet_name(sheet_name)
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))
}

fn open_editor_with_optional_sheet(path: PathBuf, sheet_name: Option<&str>) -> PyResult<Editor> {
    let sheet = match sheet_name {
        Some(name) => normalize_sheet_name(name)?,
        None => scan(&path)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?
            .into_iter()
            .next()
            .ok_or_else(|| PyRuntimeError::new_err("Workbook contains no worksheets"))?,
    };

    let opened =
        XlsxEditor::open(path, &sheet).map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
    Ok(Editor { editor: opened })
}

#[pyclass(name = "HorizAlignment", from_py_object)]
#[derive(Clone)]
struct PyHorizAlignment(HorizAlignment);

#[pyclass(name = "VertAlignment", from_py_object)]
#[derive(Clone)]
struct PyVertAlignment(VertAlignment);

#[pyclass(name = "AlignSpec", from_py_object)]
#[derive(Clone)]
struct PyAlignSpec(AlignSpec);

#[pymethods]
impl PyAlignSpec {
    #[new]
    #[pyo3(signature = (horiz = None, vert = None, wrap = false))]
    fn new(
        py: Python<'_>,
        horiz: Option<Py<PyAny>>,
        vert: Option<Py<PyAny>>,
        wrap: bool,
    ) -> PyResult<Self> {
        let h_opt = if let Some(h_obj) = horiz {
            let h_any = h_obj.bind(py);
            let h_value = h_any.getattr("value")?;
            let py_h: PyRef<PyHorizAlignment> = h_value.extract()?;
            Some(py_h.0.clone())
        } else {
            None
        };

        let v_opt = if let Some(v_obj) = vert {
            let v_any = v_obj.bind(py);
            let v_value = v_any.getattr("value")?;
            let py_v: PyRef<PyVertAlignment> = v_value.extract()?;
            Some(py_v.0.clone())
        } else {
            None
        };

        Ok(Self(AlignSpec {
            horiz: h_opt,
            vert: v_opt,
            wrap,
        }))
    }
}

#[pyfunction]
fn scan_excel(path: PathBuf) -> PyResult<Vec<String>> {
    scan(&path).map_err(|e| PyRuntimeError::new_err(e.to_string()))
}

#[pyfunction]
#[pyo3(signature = (path, sheet_name = "Sheet1"))]
fn create_excel(path: PathBuf, sheet_name: &str) -> PyResult<()> {
    rust_core::create_excel(&path, sheet_name).map_err(|e| PyRuntimeError::new_err(e.to_string()))
}

// Batches bound Python object lifetimes and let row serialization/compression
// run outside the GIL. No Polars or Arrow objects cross the native boundary.
enum InputCell {
    Value(rust_core::data::CellValue),
    Float(f64),
}
impl FromPyObject<'_, '_> for InputCell {
    type Error = PyErr;
    fn extract(cell: Borrowed<'_, '_, PyAny>) -> PyResult<Self> {
        use rust_core::data::CellValue;
        let cell = cell.cast::<PyTuple>()?;
        if cell.len() != 2 {
            return Err(PyValueError::new_err("Expected a tagged cell pair"));
        }
        let kind: u8 = cell.get_borrowed_item(0)?.extract()?;
        let value = cell.get_borrowed_item(1)?;
        Ok(match kind {
            0 => InputCell::Value(CellValue::Blank),
            1 => InputCell::Value(CellValue::Text(value.extract()?)),
            2 => InputCell::Value(CellValue::Number(value.extract()?)),
            3 => InputCell::Value(CellValue::Boolean(value.extract::<String>()? == "1")),
            5 => InputCell::Float(value.extract()?),
            _ => return Err(PyValueError::new_err("Unknown cell type")),
        })
    }
}
fn ingest_rows<F>(py: Python<'_>, rows: &Bound<'_, PyAny>, write: F) -> PyResult<()>
where
    F: FnOnce(
            &mut dyn Iterator<Item = anyhow::Result<Vec<rust_core::data::CellValue>>>,
        ) -> anyhow::Result<()>
        + Send,
{
    let iterator: Py<PyIterator> = rows.try_iter()?.unbind();
    let mut python_error = None;
    let result = py.detach(|| {
        let mut batch = Vec::new().into_iter();
        let mut done = false;
        let mut converted = std::iter::from_fn(|| {
            if batch.len() == 0 && !done {
                let next = Python::attach(|py| {
                    let mut values = Vec::with_capacity(256);
                    let mut cells = 0;
                    let mut iterator = iterator.bind(py).clone();
                    for _ in 0..256 {
                        match iterator.next() {
                            Some(item) => {
                                match item.and_then(|row| row.extract::<Vec<InputCell>>()) {
                                    Ok(row) => {
                                        cells += row.len();
                                        values.push(Ok(row));
                                        if cells >= 16_384 {
                                            break;
                                        }
                                    }
                                    Err(error) => {
                                        values.push(Err(error));
                                        done = true;
                                        break;
                                    }
                                }
                            }
                            None => {
                                done = true;
                                break;
                            }
                        }
                    }
                    values
                });
                batch = next.into_iter();
            }
            batch.next().map(|row| match row {
                Ok(row) => Ok(row
                    .into_iter()
                    .map(|cell| match cell {
                        InputCell::Value(value) => value,
                        InputCell::Float(value) => {
                            rust_core::data::CellValue::Number(format!("{value:?}"))
                        }
                    })
                    .collect()),
                Err(error) => {
                    python_error = Some(error);
                    Err(anyhow::anyhow!("Data row conversion failed"))
                }
            })
        });
        write(&mut converted)
    });
    if let Some(error) = python_error {
        return Err(error);
    }
    result.map_err(|e| PyRuntimeError::new_err(e.to_string()))
}
fn compression_level(level: i32) -> PyResult<u8> {
    if !(0..=9).contains(&level) {
        return Err(PyValueError::new_err(
            "compression_level must be between 0 and 9",
        ));
    }
    Ok(level as u8)
}

fn compression_threads(threads: i32) -> PyResult<usize> {
    if threads < 1 {
        return Err(PyValueError::new_err(
            "compression_threads must be positive",
        ));
    }
    Ok(threads as usize)
}

#[pyclass(subclass)]
struct Editor {
    editor: XlsxEditor,
}

#[pymethods]
impl Editor {
    #[new]
    #[pyo3(signature = (path, sheet_name = None))]
    fn new(path: PathBuf, sheet_name: Option<&str>) -> PyResult<Self> {
        open_editor_with_optional_sheet(path, sheet_name)
    }

    #[staticmethod]
    #[pyo3(signature = (path, sheet_name = "Sheet1"))]
    fn create(path: PathBuf, sheet_name: &str) -> PyResult<Self> {
        rust_core::create_excel(&path, sheet_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        open_editor_with_optional_sheet(path, Some(sheet_name))
    }

    #[staticmethod]
    #[pyo3(signature = (path, sheet_name = None))]
    fn open(path: PathBuf, sheet_name: Option<&str>) -> PyResult<Self> {
        open_editor_with_optional_sheet(path, sheet_name)
    }

    fn clone(&self) -> Self {
        Self {
            editor: self.editor.clone(),
        }
    }

    fn is_empty_sheet(&self) -> bool {
        self.editor.is_empty_sheet()
    }

    fn uses_1904_dates(&self) -> PyResult<bool> {
        self.editor
            .uses_1904_dates()
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    fn write_typed_rows(
        &mut self,
        py: Python<'_>,
        start: &str,
        rows: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        ingest_rows(py, rows, |rows| {
            self.editor.write_typed_table_at(start, rows)
        })
    }

    fn force_column_formats(
        &mut self,
        py: Python<'_>,
        formats: Vec<(usize, String)>,
    ) -> PyResult<()> {
        py.detach(|| self.editor.force_column_formats(&formats))
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    fn add_worksheet(&mut self, sheet_name: &str) -> PyResult<()> {
        let normalized = normalize_sheet_name(sheet_name)?;
        self.editor
            .add_worksheet(&normalized)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(())
    }

    fn add_worksheet_at(&mut self, sheet_name: &str, index: usize) -> PyResult<()> {
        let normalized = normalize_sheet_name(sheet_name)?;
        self.editor
            .add_worksheet_at(&normalized, index)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(())
    }

    fn with_worksheet(&mut self, sheet_name: &str) -> PyResult<()> {
        let normalized = normalize_sheet_name(sheet_name)?;
        self.editor
            .with_worksheet(&normalized)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(())
    }

    fn rename_worksheet(&mut self, old_name: &str, new_name: &str) -> PyResult<()> {
        let old_name = normalize_sheet_name(old_name)?;
        let new_name = normalize_sheet_name(new_name)?;
        self.editor
            .rename_worksheet(&old_name, &new_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(())
    }

    fn delete_worksheet(&mut self, sheet_name: &str) -> PyResult<()> {
        let normalized = normalize_sheet_name(sheet_name)?;
        self.editor
            .delete_worksheet(&normalized)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(())
    }

    fn set_cell(&mut self, coords: &str, cell: String) -> PyResult<()> {
        self.editor
            .set_cell(coords, cell)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    fn append_row(&mut self, cells: Vec<String>) -> PyResult<()> {
        self.editor
            .append_row(cells)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    fn append_table_at(&mut self, cells: Vec<Vec<String>>, start_cell: &str) -> PyResult<()> {
        self.editor
            .append_table_at(start_cell, cells)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    fn last_row_index(&mut self, col_name: String) -> PyResult<u32> {
        self.editor
            .get_last_row_index(&col_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    fn last_rows_index(&mut self, col_name: String) -> PyResult<Vec<u32>> {
        self.editor
            .get_last_roww_index(&col_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    #[pyo3(signature = (path = None, *, compression_level = 3, compression_threads = 1))]
    fn save(
        &mut self,
        py: Python<'_>,
        path: Option<PathBuf>,
        compression_level: i32,
        compression_threads: i32,
    ) -> PyResult<()> {
        if !(0..=9).contains(&compression_level) {
            return Err(PyValueError::new_err(
                "compression_level must be between 0 and 9",
            ));
        }
        let threads = crate::compression_threads(compression_threads)?;
        let destination = path.unwrap_or_else(|| self.editor.source_path().to_path_buf());
        py.detach(|| {
            self.editor
                .save_with_options(destination, compression_level as u8, threads)
        })
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    fn set_number_format(&mut self, range: &str, fmt: &str) -> PyResult<()> {
        self.editor
            .set_number_format(range, fmt)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(())
    }

    fn set_fill(&mut self, range: &str, fmt: &str) -> PyResult<()> {
        self.editor
            .set_fill(range, fmt)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(())
    }

    #[pyo3(signature = (range, name, size, bold = false, italic = false, align = None))]
    fn set_font(
        &mut self,
        range: &str,
        name: &str,
        size: f32,
        bold: bool,
        italic: bool,
        align: Option<PyAlignSpec>,
    ) -> PyResult<()> {
        let editor = &mut self.editor;

        if let Some(py_align_spec) = align {
            editor
                .set_font_with_alignment(range, name, size, bold, italic, &py_align_spec.0)
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        } else {
            editor
                .set_font(range, name, size, bold, italic)
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        }
        Ok(())
    }

    fn set_alignment(&mut self, range: &str, spec: PyAlignSpec) -> PyResult<()> {
        self.editor
            .set_alignment(range, &spec.0)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(())
    }

    fn merge_cells(&mut self, range: &str) -> PyResult<()> {
        self.editor
            .merge_cells(range)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(())
    }

    fn set_border(&mut self, range: &str, style: &str) -> PyResult<()> {
        self.editor
            .set_border(range, style)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(())
    }

    fn set_column_width(&mut self, col_letter: &str, width: f64) -> PyResult<()> {
        let normalized = normalize_column_letter(col_letter)?;
        let width = validate_width(width)?;
        self.editor
            .set_column_width(&normalized, width)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(())
    }

    fn set_column_width_range(&mut self, col_range: &str, width: f64) -> PyResult<()> {
        let width = validate_width(width)?;
        for col in expand_column_selector(col_range)? {
            self.editor
                .set_column_width(&col, width)
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        }
        Ok(())
    }

    fn set_columns_width(&mut self, col_letters: Vec<String>, width: f64) -> PyResult<()> {
        let width = validate_width(width)?;
        for selector in &col_letters {
            for col in expand_column_selector(selector)? {
                self.editor
                    .set_column_width(&col, width)
                    .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
            }
        }
        Ok(())
    }

    fn set_column_widths(&mut self, widths: &Bound<'_, PyDict>) -> PyResult<()> {
        for (key, value) in widths.iter() {
            let selector: String = key.extract()?;
            let width: f64 = value.extract()?;
            let width = validate_width(width)?;
            for col in expand_column_selector(&selector)? {
                self.editor
                    .set_column_width(&col, width)
                    .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
            }
        }
        Ok(())
    }

    fn remove_style(&mut self, range: &str) -> PyResult<()> {
        self.editor
            .remove_style(range)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(())
    }
}

#[pyclass]
struct WorkbookWriter {
    writer: rust_core::WorkbookWriter,
}
#[pymethods]
impl WorkbookWriter {
    #[new]
    #[pyo3(signature=(path,sheet_name="Sheet1",*,compression_level=3,compression_threads=1))]
    fn new(
        path: PathBuf,
        sheet_name: &str,
        compression_level: i32,
        compression_threads: i32,
    ) -> PyResult<Self> {
        Ok(Self {
            writer: rust_core::WorkbookWriter::create_with_options(
                path,
                sheet_name,
                crate::compression_level(compression_level)?,
                crate::compression_threads(compression_threads)?,
            )
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?,
        })
    }
    fn next_row(&self) -> PyResult<u32> {
        self.writer
            .next_row()
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }
    fn add_worksheet(&mut self, py: Python<'_>, sheet_name: &str) -> PyResult<()> {
        py.detach(|| self.writer.add_worksheet(sheet_name))
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }
    fn write_typed_rows(
        &mut self,
        py: Python<'_>,
        start: &str,
        rows: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        ingest_rows(py, rows, |rows| self.writer.write_typed_rows(start, rows))
    }
    fn force_column_formats(&mut self, formats: Vec<(usize, String)>) -> PyResult<()> {
        for (col, format) in formats {
            self.writer
                .set_column_format(col, &format)
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        }
        Ok(())
    }
    fn set_number_format(&mut self, col_range: &str, fmt: &str) -> PyResult<()> {
        for col in expand_column_selector(col_range.trim_end_matches(':'))? {
            self.writer
                .set_column_format(excel_col_to_index(&col)?, fmt)
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        }
        Ok(())
    }
    fn set_column_width(&mut self, col_range: &str, width: f64) -> PyResult<()> {
        let width = validate_width(width)?;
        for col in expand_column_selector(col_range)? {
            self.writer
                .set_column_width(excel_col_to_index(&col)?, width)
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        }
        Ok(())
    }
    #[pyo3(signature=(path=None))]
    fn save(&mut self, py: Python<'_>, path: Option<PathBuf>) -> PyResult<()> {
        py.detach(|| self.writer.save(path.as_deref()))
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }
    fn abort(&mut self, py: Python<'_>) {
        py.detach(|| self.writer.abort());
    }
}

#[pyclass]
struct Scanner {
    path: PathBuf,
}

#[pymethods]
impl Scanner {
    #[new]
    fn new(path: PathBuf) -> PyResult<Self> {
        Ok(Scanner { path })
    }

    fn get_sheets(&self) -> PyResult<Vec<String>> {
        scan_excel(self.path.clone()).map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }

    #[pyo3(signature = (sheet_name = None))]
    fn open_editor(&self, sheet_name: Option<&str>) -> PyResult<Editor> {
        open_editor_with_optional_sheet(self.path.clone(), sheet_name)
    }
}

#[pymodule]
fn excelsior(py: Python<'_>, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add_class::<Editor>()?;
    m.add_class::<WorkbookWriter>()?;
    m.add_class::<Scanner>()?;
    m.add_function(wrap_pyfunction!(scan_excel, m)?)?;
    m.add_function(wrap_pyfunction!(create_excel, m)?)?;

    m.add_class::<PyAlignSpec>()?;

    let horiz_enum = py.import("enum")?.getattr("Enum")?;
    let horiz_members = PyDict::new(py);
    horiz_members.set_item("Left", PyHorizAlignment(HorizAlignment::Left))?;
    horiz_members.set_item("Center", PyHorizAlignment(HorizAlignment::Center))?;
    horiz_members.set_item("Right", PyHorizAlignment(HorizAlignment::Right))?;
    horiz_members.set_item("Fill", PyHorizAlignment(HorizAlignment::Fill))?;
    horiz_members.set_item("Justify", PyHorizAlignment(HorizAlignment::Justify))?;
    let horiz_cls = horiz_enum.call1(("HorizAlignment", horiz_members))?;
    m.add("HorizAlignment", horiz_cls)?;

    let vert_enum = py.import("enum")?.getattr("Enum")?;
    let vert_members = PyDict::new(py);
    vert_members.set_item("Top", PyVertAlignment(VertAlignment::Top))?;
    vert_members.set_item("Center", PyVertAlignment(VertAlignment::Center))?;
    vert_members.set_item("Bottom", PyVertAlignment(VertAlignment::Bottom))?;
    vert_members.set_item("Justify", PyVertAlignment(VertAlignment::Justify))?;
    let vert_cls = vert_enum.call1(("VertAlignment", vert_members))?;
    m.add("VertAlignment", vert_cls)?;

    Ok(())
}
