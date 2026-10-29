// use mimalloc::MiMalloc;

// #[global_allocator]
// static GLOBAL: MiMalloc = MiMalloc;
mod archive;
pub mod data;
mod parallel;
mod rows;
mod shared;
mod xml;
use shared::{Bytes, Shared};
pub mod files_part;
pub mod streaming;
pub mod workbook;
pub use streaming::WorkbookWriter;
pub use workbook::create_excel;
mod read_part;
pub mod style;
mod test;
use std::{
    collections::{HashMap, HashSet},
    fs::File,
    io::Read,
    path::{Path, PathBuf},
};

use anyhow::{Context, Result};

use crate::{
    files_part::sheet_path_for_name,
    style::{AlignSpec, HorizAlignment, VertAlignment},
};
// use tempfile::NamedTempFile;
// use zip::{ZipArchive, ZipWriter, write::FileOptions};

/// `XlsxEditor` provides functionality to open, modify, and save XLSX files.
/// It allows appending rows and tables to a specified sheet within an XLSX file.

#[derive(Hash, Eq, PartialEq, Clone)]
struct FontKey {
    name: String,
    size_100: u32,
    bold: bool,
    italic: bool,
}
#[derive(Hash, Eq, PartialEq, Clone)]
struct StyleKey {
    num_fmt_id: u32,
    font_id: Option<u32>,
    fill_id: Option<u32>,
    border_id: Option<u32>,
    align: Option<(Option<HorizAlignment>, Option<VertAlignment>, bool)>, // wrap
}
#[allow(dead_code)]
#[derive(Clone)]
struct XfParts {
    num_fmt_id: u32,
    font_id: Option<u32>,
    fill_id: Option<u32>,
    border_id: Option<u32>,
    align: Option<AlignSpec>,
}

#[derive(Clone)]
struct StyleIndex {
    xfs: Vec<XfParts>, // index == style_id

    numfmt_by_code: HashMap<String, u32>,
    numfmt_code_by_id: HashMap<u32, String>,
    next_custom_numfmt: u32, // >=164

    font_by_key: HashMap<FontKey, u32>,
    fill_by_rgb: HashMap<String, u32>,   // RGB в верхнем регистре
    border_by_key: HashMap<String, u32>, // единый style для всех сторон

    xf_by_key: HashMap<StyleKey, u32>,

    fonts_count: u32,
    fills_count: u32,
    borders_count: u32,
}

#[derive(Clone)]
pub struct XlsxEditor {
    src_path: PathBuf,
    source: archive::Source,
    sheet_path: String,
    sheet_xml: Bytes,
    last_row: u32,
    sheet_data_end: Option<usize>,
    row_index: Option<Shared<rows::RowIndex>>,
    dimensions: Shared<HashMap<String, (usize, u32)>>,
    styles_xml: Bytes,   // содержимое styles.xml
    workbook_xml: Bytes, // содержимое workbook.xml (может изменяться)
    rels_xml: Bytes,     // содержимое workbook.xml.rels
    new_files: Shared<HashMap<String, Bytes>>,
    styles_index: Option<Shared<StyleIndex>>,
    loaded_files: Shared<HashMap<String, Bytes>>,
    removed_files: Shared<HashSet<String>>, // ← НОВОЕ: пути внутри ZIP, которые надо выкинуть
}

impl XlsxEditor {
    pub fn is_empty_sheet(&self) -> bool {
        self.last_row == 0
    }
    pub fn source_path(&self) -> &Path {
        &self.src_path
    }

    /// Creates an empty workbook on disk and opens its first sheet for editing.
    /// An existing destination is replaced, just as with `create_excel`.
    pub fn create<P: AsRef<Path>>(path: P, sheet_name: &str) -> Result<Self> {
        create_excel(&path, sheet_name)?;
        Self::open(path, &workbook::normalize_sheet_name(sheet_name)?)
    }

    /// Opens an XLSX file and prepares a specific sheet for editing by its name.
    ///
    /// Resolve the selected sheet through workbook relationships, then open its
    /// actual package path (which need not match its numeric sheet ID).
    ///
    /// # Arguments
    /// * `src` - The path to the XLSX file.
    /// * `sheet_name` - The name of the sheet to open (e.g., "Sheet1").
    ///
    /// # Returns
    /// A `Result` containing an `XlsxEditor` instance if successful, or an `anyhow::Error` otherwise.
    pub fn open<P: AsRef<Path>>(src: P, sheet_name: &str) -> Result<Self> {
        let sheet_path = sheet_path_for_name(src.as_ref(), sheet_name)?;
        Self::open_sheet_path(src, sheet_path)
    }

    /// Appends a single row of cells to the end of the current sheet.
    ///
    /// Each item in the `cells` iterator will be converted to a string and written as a cell.
    /// The cell type (number or inline string) is inferred based on whether the value can be parsed as a float.
    ///
    /// # Arguments
    /// * `cells` - An iterator over values that can be converted to strings, representing the cells in the new row.
    ///
    /// # Returns
    /// A `Result` indicating success or an `anyhow::Error` if the operation fails.
    pub fn append_row<I, S>(&mut self, cells: I) -> anyhow::Result<()>
    where
        I: IntoIterator<Item = S>,
        S: ToString,
    {
        self.append_table(std::iter::once(cells))
    }

    /// Append rows using the same validated writer as positioned tables.
    pub fn append_table<R, I, S>(&mut self, rows: R) -> Result<()>
    where
        R: IntoIterator<Item = I>,
        I: IntoIterator<Item = S>,
        S: ToString,
    {
        let start = self
            .last_row
            .checked_add(1)
            .context("XLSX row limit exceeded")?;
        self.append_table_at(&format!("A{start}"), rows)
    }

    /// Appends multiple rows (a table) starting at a specified coordinate in the current sheet.
    ///
    /// This function allows inserting a table at a specific cell coordinate (e.g., "A1", "C5").
    /// If the target rows already exist, their cells will be updated. If the target rows are beyond
    /// the current last row, new rows will be appended.
    ///
    /// # Arguments
    /// * `start_coord` - The starting cell coordinate (e.g., "A1") where the table should begin.
    /// * `rows` - An iterator over iterators of values that can be converted to strings, representing the rows and cells of the table.
    ///
    /// # Returns
    /// A `Result` indicating success or an `anyhow::Error` if the operation fails.
    pub fn append_table_at<R, I, S>(&mut self, start_coord: &str, rows: R) -> anyhow::Result<()>
    where
        R: IntoIterator<Item = I>,
        I: IntoIterator<Item = S>,
        S: ToString,
    {
        let rows = rows.into_iter().map(|row| {
            Ok(row
                .into_iter()
                .map(|v| data::CellValue::infer(v.to_string()))
                .collect())
        });
        self.write_typed_table_at(start_coord, rows)
    }

    /// Sets the value of a specific cell in the sheet.
    ///
    /// This function allows updating an existing cell or creating a new one if it doesn't exist.
    /// The cell type (number or inline string) is inferred based on whether the value can be parsed as a float.
    ///
    /// # Arguments
    /// * `coord` - The cell coordinate (e.g., "A1", "B2").
    /// * `value` - The value to set for the cell, which can be converted to a string.
    ///
    /// # Returns
    /// A `Result` indicating success or an `anyhow::Error` if the operation fails.
    pub fn set_cell<S: ToString>(&mut self, coord: &str, value: S) -> Result<()> {
        let (col, row) = data::parse_coord(coord)?;
        let value = data::CellValue::infer(value.to_string());
        if row <= self.last_row && self.replace_indexed_cell(col, row, &value)? {
            return Ok(());
        }
        self.write_typed_table_at(coord, std::iter::once(Ok(vec![value])))
    }
}

fn build_row_xml(
    row_num: u32,
    start_col_idx: usize,
    values: &[data::CellValue],
) -> Result<Vec<u8>> {
    build_styled_row_xml(
        row_num,
        start_col_idx,
        values,
        &std::collections::BTreeMap::new(),
    )
}
fn build_styled_row_xml(
    row_num: u32,
    start_col_idx: usize,
    values: &[data::CellValue],
    styles: &std::collections::BTreeMap<usize, u32>,
) -> Result<Vec<u8>> {
    let mut row = Vec::new();
    row.extend_from_slice(format!(r#"<row r="{row_num}">"#).as_bytes());
    for (offset, value) in values.iter().enumerate() {
        let coord = format!("{}{}", col_idx_to_letters(start_col_idx + offset), row_num);
        row.extend_from_slice(&data::build_typed_cell_xml(
            &coord,
            value,
            styles.get(&(start_col_idx + offset)).copied(),
        )?);
    }
    row.extend_from_slice(b"</row>");
    Ok(row)
}

fn patch_row_xml(
    row: &xml::Element<'_>,
    row_num: u32,
    start_col_idx: usize,
    values: &[data::CellValue],
) -> Result<Vec<u8>> {
    let mut cells = std::collections::BTreeMap::new();
    let mut others = Vec::new();
    for element in row.children() {
        let element = element?;
        if element.is("c", xml::SHEET_NS) {
            let coord = element.attr("r")?.context("Cell coordinate missing")?;
            let (col, _) = data::parse_coord(&coord)?;
            anyhow::ensure!(
                cells.insert(col, element).is_none(),
                "Duplicate cell coordinate"
            );
        } else {
            others.push(element);
        }
    }
    let mut content = Vec::new();
    let mut replacements = values.iter().enumerate().peekable();
    for (col, cell) in cells {
        while replacements
            .peek()
            .is_some_and(|(offset, _)| start_col_idx + offset < col)
        {
            let (offset, value) = replacements.next().unwrap();
            let coord = format!("{}{}", col_idx_to_letters(start_col_idx + offset), row_num);
            content.extend_from_slice(&xml::child_fragment(
                row,
                &data::build_typed_cell_xml(&coord, value, None)?,
            )?);
        }
        if replacements
            .peek()
            .is_some_and(|(offset, _)| start_col_idx + offset == col)
        {
            let (_, value) = replacements.next().unwrap();
            let sid = cell.attr("s")?.map(|v| v.parse::<u32>()).transpose()?;
            let coord = format!("{}{}", col_idx_to_letters(col), row_num);
            content.extend_from_slice(&xml::child_fragment(
                row,
                &data::build_typed_cell_xml(&coord, value, sid)?,
            )?);
        } else {
            content.extend_from_slice(cell.bytes());
        }
    }
    for (offset, value) in replacements {
        let coord = format!("{}{}", col_idx_to_letters(start_col_idx + offset), row_num);
        content.extend_from_slice(&xml::child_fragment(
            row,
            &data::build_typed_cell_xml(&coord, value, None)?,
        )?);
    }
    for element in others {
        content.extend_from_slice(element.bytes());
    }
    row.with_content(&content)
}

fn col_idx_to_letters(mut idx: usize) -> String {
    let mut s = String::new();
    loop {
        let rem = idx % 26;
        s.insert(0, (b'A' + rem as u8) as char);
        if idx < 26 {
            break;
        }
        idx = idx / 26 - 1;
    }
    s
}

pub fn scan<P: AsRef<Path>>(source: P) -> Result<Vec<String>> {
    let mut archive = zip::ZipArchive::new(File::open(source)?)?;
    let mut part = archive.by_name("xl/workbook.xml")?;
    let mut bytes = Vec::new();
    part.read_to_end(&mut bytes)?;
    let sheets = xml::require(&bytes, "sheets")?;
    let mut names = Vec::new();
    for sheet in sheets.children() {
        let sheet = sheet?;
        if sheet.is("sheet", xml::SHEET_NS) {
            names.push(
                sheet
                    .attr("name")?
                    .context("Sheet name missing")?
                    .into_owned(),
            );
        }
    }
    Ok(names)
}

impl XlsxEditor {
    pub fn merge_cells(&mut self, range: &str) -> Result<()> {
        let (start, end) = range
            .split_once(':')
            .context("Merge range must have two cell coordinates")?;
        let (c0, r0) = data::parse_coord(start)?;
        let (c1, r1) = data::parse_coord(end)?;
        anyhow::ensure!(c0 <= c1 && r0 <= r1, "Invalid merge range order");
        let normalized = format!(
            "{}{}:{}{}",
            col_idx_to_letters(c0),
            r0,
            col_idx_to_letters(c1),
            r1
        );
        let root = xml::root(&self.sheet_xml)?;
        let tag = format!(r#"<mergeCell ref="{normalized}"/>"#);
        let updated = if let Some(merges) = root.child("mergeCells")? {
            let mut content = self.sheet_xml[merges.content.clone()].to_vec();
            content.extend_from_slice(&xml::child_fragment(&merges, tag.as_bytes())?);
            let count = merges.children().count() + 1;
            let replacement = merges.with_content(&content)?;
            let element = xml::root(&replacement)?;
            let opening = element.set_attr("count", Some(&count.to_string()))?;
            let mut replacement = replacement.clone();
            replacement.splice(element.open, opening);
            xml::replace_element(&self.sheet_xml, &merges, &replacement)
        } else {
            let data = root.require_child("sheetData")?;
            let block = format!(r#"<mergeCells count="1">{tag}</mergeCells>"#);
            let mut out = self.sheet_xml.to_vec();
            out.splice(
                data.full.end..data.full.end,
                xml::child_fragment(&root, block.as_bytes())?
                    .iter()
                    .copied(),
            );
            out
        };
        self.sheet_xml = updated.into();
        self.sheet_data_end = None;
        self.row_index = None;
        Ok(())
    }
}
