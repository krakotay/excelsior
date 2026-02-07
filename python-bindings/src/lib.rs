use pyo3::PyRefMut;
use pyo3::exceptions::PyRuntimeError;
use pyo3::prelude::*;
use pyo3::types::PyDict;
use rust_core::style::{AlignSpec, HorizAlignment, VertAlignment};
use rust_core::{XlsxEditor, scan};
use std::fs::File;
use std::path::{Path, PathBuf};
use zip::write::FileOptions;

fn py_value_to_excel_string(value: &Bound<'_, PyAny>) -> PyResult<String> {
    if value.is_none() {
        return Ok(String::new());
    }
    Ok(value.str()?.to_str()?.to_owned())
}

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
    let mut idx = 0usize;
    let up = col.trim().to_ascii_uppercase();
    if up.is_empty() || !up.bytes().all(|b| b.is_ascii_uppercase()) {
        return Err(PyRuntimeError::new_err(format!(
            "Invalid Excel column: {col}"
        )));
    }
    for b in up.bytes() {
        idx = idx * 26 + (b - b'A' + 1) as usize;
    }
    Ok(idx - 1)
}

fn normalize_column_letter(col: &str) -> PyResult<String> {
    let normalized = col.trim().to_ascii_uppercase();
    if normalized.is_empty() || !normalized.bytes().all(|b| b.is_ascii_uppercase()) {
        return Err(PyRuntimeError::new_err(format!(
            "Invalid Excel column: {col}"
        )));
    }
    Ok(normalized)
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
    let normalized = sheet_name.trim();
    if normalized.is_empty() {
        return Err(PyRuntimeError::new_err("Sheet name cannot be empty"));
    }
    if normalized.len() > 31 {
        return Err(PyRuntimeError::new_err(
            "Sheet name cannot be longer than 31 characters",
        ));
    }
    if normalized
        .chars()
        .any(|c| matches!(c, ':' | '\\' | '/' | '?' | '*' | '[' | ']'))
    {
        return Err(PyRuntimeError::new_err(
            "Sheet name contains forbidden Excel characters (: \\ / ? * [ ])",
        ));
    }
    Ok(normalized.to_owned())
}

fn xml_escape_attr(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('\'', "&apos;")
}

fn create_empty_excel_file(path: &Path, sheet_name: &str) -> PyResult<()> {
    let sheet_name = normalize_sheet_name(sheet_name)?;
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent).map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
    }

    let file = File::create(path).map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
    let mut zip = zip::ZipWriter::new(file);
    let options: FileOptions<'_, ()> =
        FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#;

    let rels_root = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
  <sheets>\
    <sheet name=\"{}\" sheetId=\"1\" r:id=\"rId1\"/>\
  </sheets>\
</workbook>",
        xml_escape_attr(&sheet_name)
    );

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

    let sheet1 = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData></sheetData>
</worksheet>"#;

    let styles = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>"#;

    zip.start_file("[Content_Types].xml", options)
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
    std::io::Write::write_all(&mut zip, content_types.as_bytes())
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;

    zip.start_file("_rels/.rels", options)
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
    std::io::Write::write_all(&mut zip, rels_root.as_bytes())
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;

    zip.start_file("xl/workbook.xml", options)
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
    std::io::Write::write_all(&mut zip, workbook.as_bytes())
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
    std::io::Write::write_all(&mut zip, workbook_rels.as_bytes())
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
    std::io::Write::write_all(&mut zip, sheet1.as_bytes())
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;

    zip.start_file("xl/styles.xml", options)
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
    std::io::Write::write_all(&mut zip, styles.as_bytes())
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;

    zip.finish()
        .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
    Ok(())
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
    create_empty_excel_file(&path, sheet_name)
}

#[pyclass]
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
        create_empty_excel_file(&path, sheet_name)?;
        open_editor_with_optional_sheet(path, Some(sheet_name))
    }

    #[staticmethod]
    #[pyo3(signature = (path, sheet_name = None))]
    fn open(path: PathBuf, sheet_name: Option<&str>) -> PyResult<Self> {
        open_editor_with_optional_sheet(path, sheet_name)
    }

    fn add_worksheet<'py>(
        mut slf: PyRefMut<'py, Self>,
        sheet_name: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let normalized = normalize_sheet_name(sheet_name)?;
        slf.editor
            .add_worksheet(&normalized)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }

    fn add_worksheet_at<'py>(
        mut slf: PyRefMut<'py, Self>,
        sheet_name: &str,
        index: usize,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let normalized = normalize_sheet_name(sheet_name)?;
        slf.editor
            .add_worksheet_at(&normalized, index)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }

    fn with_worksheet<'py>(
        mut slf: PyRefMut<'py, Self>,
        sheet_name: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let normalized = normalize_sheet_name(sheet_name)?;
        slf.editor
            .with_worksheet(&normalized)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }

    fn rename_worksheet<'py>(
        mut slf: PyRefMut<'py, Self>,
        old_name: &str,
        new_name: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let old_name = normalize_sheet_name(old_name)?;
        let new_name = normalize_sheet_name(new_name)?;
        slf.editor
            .rename_worksheet(&old_name, &new_name)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }

    fn delete_worksheet<'py>(
        mut slf: PyRefMut<'py, Self>,
        sheet_name: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let normalized = normalize_sheet_name(sheet_name)?;
        slf.editor
            .delete_worksheet(&normalized)
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

    #[pyo3(signature = (py_df, start_cell = None))]
    fn with_polars(
        &mut self,
        py: Python<'_>,
        py_df: &Bound<'_, PyAny>,
        start_cell: Option<String>,
    ) -> PyResult<()> {
        let columns: Vec<String> = py_df
            .getattr("columns")
            .map_err(|_| PyRuntimeError::new_err("Expected polars.DataFrame with .columns"))?
            .extract()?;

        let rows_iter = py_df
            .call_method0("iter_rows")
            .map_err(|_| PyRuntimeError::new_err("Expected polars.DataFrame with .iter_rows()"))?;

        let mut table: Vec<Vec<String>> = Vec::new();
        table.push(columns);

        for row in rows_iter.try_iter()? {
            let row_any = row?;
            let row_values: Vec<Py<PyAny>> = row_any.extract()?;
            let mut out_row = Vec::with_capacity(row_values.len());
            for value in row_values {
                out_row.push(py_value_to_excel_string(value.bind(py).as_any())?);
            }
            table.push(out_row);
        }

        let start = start_cell.as_deref().unwrap_or("A1");
        self.editor
            .append_table_at(start, table)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
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
        align: Option<PyAlignSpec>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let editor = &mut slf.editor;

        if let Some(py_align_spec) = align {
            editor
                .set_font_with_alignment(range, name, size, bold, italic, &py_align_spec.0)
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
        spec: PyAlignSpec,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .set_alignment(range, &spec.0)
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
        let normalized = normalize_column_letter(col_letter)?;
        let width = validate_width(width)?;
        slf.editor
            .set_column_width(&normalized, width)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        Ok(slf)
    }

    fn set_column_width_range<'py>(
        mut slf: PyRefMut<'py, Self>,
        col_range: &str,
        width: f64,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let width = validate_width(width)?;
        for col in expand_column_selector(col_range)? {
            slf.editor
                .set_column_width(&col, width)
                .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
        }
        Ok(slf)
    }

    fn set_columns_width<'py>(
        mut slf: PyRefMut<'py, Self>,
        col_letters: Vec<String>,
        width: f64,
    ) -> PyResult<PyRefMut<'py, Self>> {
        let width = validate_width(width)?;
        for selector in &col_letters {
            for col in expand_column_selector(selector)? {
                slf.editor
                    .set_column_width(&col, width)
                    .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
            }
        }
        Ok(slf)
    }

    fn set_column_widths<'py>(
        mut slf: PyRefMut<'py, Self>,
        widths: &Bound<'_, PyDict>,
    ) -> PyResult<PyRefMut<'py, Self>> {
        for (key, value) in widths.iter() {
            let selector: String = key.extract()?;
            let width: f64 = value.extract()?;
            let width = validate_width(width)?;
            for col in expand_column_selector(&selector)? {
                slf.editor
                    .set_column_width(&col, width)
                    .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
            }
        }
        Ok(slf)
    }

    fn remove_style<'py>(
        mut slf: PyRefMut<'py, Self>,
        range: &str,
    ) -> PyResult<PyRefMut<'py, Self>> {
        slf.editor
            .remove_style(range)
            .map_err(|e| PyRuntimeError::new_err(e.to_string()))?;
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

    #[pyo3(signature = (sheet_name = None))]
    fn open_editor(&self, sheet_name: Option<&str>) -> PyResult<Editor> {
        open_editor_with_optional_sheet(self.path.clone(), sheet_name)
    }
}

#[pymodule]
fn excelsior(py: Python<'_>, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Editor>()?;
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
