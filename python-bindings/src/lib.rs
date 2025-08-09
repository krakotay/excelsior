use pyo3::exceptions::PyRuntimeError;
use pyo3::prelude::*;

use pyo3::PyRefMut;
use pyo3::types::PyDict;
use rust_core::{XlsxEditor, scan};
use std::path::PathBuf;

#[cfg(feature = "polars")]
use pyo3_polars::PyDataFrame;
#[cfg(feature = "polars")]
fn index_to_excel_col(mut idx: usize) -> String {
    let mut col = String::new();
    idx += 1; // 1-based
    while idx > 0 {
        let rem = (idx - 1) % 26;
        col.insert(0, (b'A' + rem as u8) as char);
        idx = (idx - 1) / 26;
    }
    col
}
// Импортируем типы из rust_core
use rust_core::style::{AlignSpec, HorizAlignment, VertAlignment};

// --- ОБЕРТКИ ДЛЯ ENUM-ОВ ---

#[pyclass(name = "HorizAlignment")]
#[derive(Clone)]
struct PyHorizAlignment(HorizAlignment);

#[pyclass(name = "VertAlignment")]
#[derive(Clone)]
struct PyVertAlignment(VertAlignment);

#[pyclass(name = "AlignSpec")]
#[derive(Clone)]
struct PyAlignSpec(AlignSpec);

#[pymethods]
impl PyAlignSpec {
    #[new]
    #[pyo3(signature = (horiz = None, vert = None, wrap = false))]
    fn new(
        py: Python<'_>,          // <--- Запрашиваем доступ к GIL
        horiz: Option<PyObject>, // <--- Принимаем PyObject
        vert: Option<PyObject>,  // <--- Принимаем PyObject
        wrap: bool,
    ) -> PyResult<Self> {
        // Извлекаем .value из горизонтального выравнивания, если оно есть
        let h_opt = if let Some(h_obj) = horiz {
            // "Привязываем" PyObject к GIL, чтобы работать с ним
            let h_any = h_obj.bind(py);
            // Получаем атрибут .value
            let h_value = h_any.getattr("value")?;
            // Извлекаем из .value нашу Rust-структуру
            let py_h: PyRef<PyHorizAlignment> = h_value.extract()?;
            // Клонируем внутренние данные
            Some(py_h.0.clone())
        } else {
            None
        };

        // То же самое для вертикального
        let v_opt = if let Some(v_obj) = vert {
            let v_any = v_obj.bind(py);
            let v_value = v_any.getattr("value")?;
            let py_v: PyRef<PyVertAlignment> = v_value.extract()?;
            Some(py_v.0.clone())
        } else {
            None
        };

        // Создаем и возвращаем финальную структуру
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
#[pyclass]
struct Editor {
    editor: XlsxEditor,
}

#[pymethods]
impl Editor {
    #[new]
    #[pyo3(signature = (path, sheet_name))]
    fn new(path: PathBuf, sheet_name: &str) -> PyResult<Self> {
        let openned = XlsxEditor::open(path, sheet_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(Editor { editor: openned })
    }
    fn add_worksheet<'py>(
        mut slf: PyRefMut<'py, Self>,
        sheet_name: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .add_worksheet(sheet_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }
    fn add_worksheet_at<'py>(
        mut slf: PyRefMut<'py, Self>,
        sheet_name: &str,
        index: usize,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .add_worksheet_at(sheet_name, index)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }
    fn with_worksheet<'py>(
        mut slf: PyRefMut<'py, Self>,
        sheet_name: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .with_worksheet(sheet_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }
    fn rename_worksheet<'py>(
        mut slf: PyRefMut<'py, Self>,
        old_name: &str,
        new_name: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .rename_worksheet(old_name, new_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }
    fn delete_worksheet<'py>(
        mut slf: PyRefMut<'py, Self>,
        sheet_name: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .delete_worksheet(sheet_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
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

    fn save(&mut self, path: PathBuf) -> PyResult<()> {
        self.editor
            .save(path)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))
    }
    #[cfg(feature = "polars")]
    #[pyo3(signature = (py_df, start_cell = None, default_width = 15.0))]
    fn with_polars(
        &mut self,
        py_df: PyDataFrame,
        start_cell: Option<String>,
        default_width: f64,
    ) -> PyResult<()> {
        let df = py_df.into();
        let start = start_cell.as_deref();
        self.editor
            .with_polars(&df, start)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;

        // --- Вот тут автоприменяем ширину к столбцам ---
        // Определяем имена столбцов из DataFrame (через polars)
        let columns: Vec<String> = df
            .get_column_names()
            .iter()
            .map(|s| s.to_string())
            .collect();

        // Вставляем ширину для каждого столбца
        for col in &columns {
            // Можно сделать функцию для конвертации индекса столбца в Excel-букву, если нужно
            let col_letter = index_to_excel_col(columns.iter().position(|c| c == col).unwrap());
            self.editor
                .set_column_width(&col_letter, default_width)
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        }

        Ok(())
    }
    fn set_number_format<'py>(
        mut slf: PyRefMut<'py, Self>,
        range: &str,
        fmt: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .set_number_format(range, fmt)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }

    fn set_fill<'py>(
        mut slf: PyRefMut<'py, Self>,
        range: &str,
        fmt: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .set_fill(range, fmt)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }
    #[pyo3(signature = (range, name, size, bold = false, italic = false, align = None))]
    fn set_font<'py>(
        mut slf: PyRefMut<'py, Self>,
        range: &str,
        name: &str,
        size: f32,
        bold: bool,
        italic: bool,
        align: Option<PyAlignSpec>, // <--- ИЗМЕНЕНО: принимаем PyAlignSpec
    ) -> PyResult<PyRefMut<'py, Self>> {
        let editor = &mut slf.editor;

        // Конвертируем PyAlignSpec в rust_core::AlignSpec вручную
        if let Some(py_align_spec) = align {
            editor
                .set_font_with_alignment(range, name, size, bold, italic, &py_align_spec.0) // <--- ИЗМЕНЕНО: используем .0
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        } else {
            editor
                .set_font(range, name, size, bold, italic)
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        }
        Ok(slf)
    }

    fn set_alignment<'py>(
        mut slf: PyRefMut<'py, Self>,
        range: &str,
        spec: PyAlignSpec, // <--- ИЗМЕНЕНО: принимаем PyAlignSpec
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .set_alignment(range, &spec.0) // <--- ИЗМЕНЕНО: используем .0
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }
    fn merge_cells<'py>(
        mut slf: PyRefMut<'py, Self>,
        range: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .merge_cells(range)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }
    fn set_border<'py>(
        mut slf: PyRefMut<'py, Self>,
        range: &str,
        style: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .set_border(range, style)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }
    fn set_column_width<'py>(
        mut slf: PyRefMut<'py, Self>,
        col_letter: &str,
        width: f64,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .set_column_width(col_letter, width)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }
    fn set_columns_width<'py>(
        mut slf: PyRefMut<'py, Self>,
        col_letters: Vec<String>,
        width: f64,
    ) -> PyResult<PyRefMut<'py, Self>> {
        for col_letter in col_letters.iter() {
            slf.editor
                .set_column_width(col_letter, width)
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        }
        Ok(slf)
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
    fn open_editor(&self, sheet_name: String) -> PyResult<Editor> {
        let openned = XlsxEditor::open(self.path.clone(), &sheet_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(Editor { editor: openned })
    }
}

#[pymodule]
fn excelsior(py: Python<'_>, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Editor>()?;
    m.add_class::<Scanner>()?;
    m.add_function(wrap_pyfunction!(scan_excel, m)?)?;

    // --- РЕГИСТРАЦИЯ НОВЫХ КЛАССОВ И ENUM-ОВ ---

    // 1. Добавляем класс AlignSpec
    m.add_class::<PyAlignSpec>()?;

    // 2. Создаем Python Enum для HorizAlignment
    let horiz_enum = py.import("enum")?.getattr("Enum")?;
    let horiz_members = PyDict::new(py);
    horiz_members.set_item("Left", PyHorizAlignment(HorizAlignment::Left))?;
    horiz_members.set_item("Center", PyHorizAlignment(HorizAlignment::Center))?;
    horiz_members.set_item("Right", PyHorizAlignment(HorizAlignment::Right))?;
    horiz_members.set_item("Fill", PyHorizAlignment(HorizAlignment::Fill))?;
    horiz_members.set_item("Justify", PyHorizAlignment(HorizAlignment::Justify))?;
    let horiz_cls = horiz_enum.call1(("HorizAlignment", horiz_members))?;
    m.add("HorizAlignment", horiz_cls)?;

    // 3. Создаем Python Enum для VertAlignment
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
