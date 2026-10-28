// use mimalloc::MiMalloc;

// #[global_allocator]
// static GLOBAL: MiMalloc = MiMalloc;
pub mod data;
pub mod files_part;
pub mod workbook;
use memchr::memmem;
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
use quick_xml::{Reader, events::Event};

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
    sheet_path: String,
    sheet_xml: Vec<u8>,
    last_row: u32,
    styles_xml: Vec<u8>,   // содержимое styles.xml
    workbook_xml: Vec<u8>, // содержимое workbook.xml (может изменяться)
    rels_xml: Vec<u8>,     // содержимое workbook.xml.rels
    new_files: HashMap<String, Vec<u8>>,
    styles_index: Option<StyleIndex>,
    loaded_files: std::collections::HashMap<String, Vec<u8>>,
    removed_files: HashSet<String>, // ← НОВОЕ: пути внутри ZIP, которые надо выкинуть
}

impl XlsxEditor {
    /// Creates an empty workbook on disk and opens its first sheet for editing.
    /// An existing destination is replaced, just as with `create_excel`.
    pub fn create<P: AsRef<Path>>(path: P, sheet_name: &str) -> Result<Self> {
        create_excel(&path, sheet_name)?;
        Self::open(path, &workbook::normalize_sheet_name(sheet_name)?)
    }

    /// Opens an XLSX file and prepares a specific sheet for editing by its name.
    ///
    /// This function first scans the workbook to find the sheet ID corresponding to the given sheet name,
    /// then calls `open_sheet` with the found ID.
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
        self.write_typed_table_at(
            coord,
            std::iter::once(Ok(vec![data::CellValue::infer(value.to_string())])),
        )
    }
}

fn build_row_xml(
    row_num: u32,
    start_col_idx: usize,
    values: &[data::CellValue],
) -> Result<Vec<u8>> {
    let mut row = Vec::new();
    row.extend_from_slice(format!(r#"<row r="{row_num}">"#).as_bytes());
    for (offset, value) in values.iter().enumerate() {
        let coord = format!("{}{}", col_idx_to_letters(start_col_idx + offset), row_num);
        row.extend_from_slice(&data::build_typed_cell_xml(&coord, value, None)?);
    }
    row.extend_from_slice(b"</row>");
    Ok(row)
}

fn patch_row_xml(
    row_xml: &[u8],
    row_num: u32,
    start_col_idx: usize,
    values: &[data::CellValue],
) -> Result<Vec<u8>> {
    let row_slice = if row_xml.ends_with(b"/>") {
        let mut expanded = row_xml[..row_xml.len() - 2].to_vec();
        expanded.extend_from_slice(b"></row>");
        expanded
    } else {
        row_xml.to_vec()
    };
    let row_tag_end = find_bytes(&row_slice, b">").context("malformed row tag")?;
    let content_start = row_tag_end + 1;
    let content_end = row_slice.len() - "</row>".len();

    let find_c = memmem::Finder::new(b"<c");
    let find_gt = memmem::Finder::new(b">");

    let mut cell_ranges: Vec<(usize, usize, usize)> = Vec::new();
    let mut j = content_start;
    while j < content_end {
        let Some(cpos_rel) = find_c.find(&row_slice[j..content_end]) else {
            break;
        };
        let cpos = j + cpos_rel;
        let next = *row_slice.get(cpos + 2).unwrap_or(&b'>');
        let is_cell = matches!(next, b' ' | b'>' | b'/' | b'r' | b's' | b't');
        let tag_end = find_gt
            .find(&row_slice[cpos..content_end])
            .context("cell tag end")?
            + cpos;
        if !is_cell {
            j = tag_end + 1;
            continue;
        }

        let self_closing = tag_end > cpos && row_slice[tag_end - 1] == b'/';
        let cell_end = if self_closing {
            tag_end + 1
        } else {
            let close =
                find_bytes_from(&row_slice, b"</c>", tag_end + 1).context("</c> missing")?;
            close + 4
        };

        if let Some(cell_col) = extract_cell_col_idx(&row_slice[cpos..cell_end]) {
            cell_ranges.push((cell_col, cpos, cell_end));
        }
        j = cell_end;
    }

    let mut replacements: Vec<(usize, usize, Vec<u8>)> = Vec::new();
    let mut inserts: Vec<(usize, Vec<u8>)> = Vec::new();

    for (offset, value) in values.iter().enumerate() {
        let col_idx = start_col_idx + offset;
        let coord = format!("{}{}", col_idx_to_letters(col_idx), row_num);
        match cell_ranges.binary_search_by_key(&col_idx, |(cell_col, _, _)| *cell_col) {
            Ok(ix) => {
                let (_, start, end) = cell_ranges[ix];
                let style_id = extract_style_id(&row_slice[start..end]);
                replacements.push((
                    start,
                    end,
                    data::build_typed_cell_xml(&coord, value, style_id)?,
                ));
            }
            Err(ix) => {
                let insert_at = if ix < cell_ranges.len() {
                    cell_ranges[ix].1
                } else {
                    content_end
                };
                inserts.push((insert_at, data::build_typed_cell_xml(&coord, value, None)?));
            }
        }
    }

    let mut out = Vec::with_capacity(row_slice.len() + values.len() * 32);
    let mut cursor = 0usize;
    replacements.sort_by_key(|(start, _, _)| *start);
    inserts.sort_by_key(|(pos, _)| *pos);
    let mut insert_ix = 0usize;

    for (start, end, replacement) in replacements {
        while insert_ix < inserts.len() && inserts[insert_ix].0 <= start {
            let (pos, ref insert) = inserts[insert_ix];
            out.extend_from_slice(&row_slice[cursor..pos]);
            out.extend_from_slice(insert);
            cursor = pos;
            insert_ix += 1;
        }
        out.extend_from_slice(&row_slice[cursor..start]);
        out.extend_from_slice(&replacement);
        cursor = end;
    }

    while insert_ix < inserts.len() {
        let (pos, ref insert) = inserts[insert_ix];
        out.extend_from_slice(&row_slice[cursor..pos]);
        out.extend_from_slice(insert);
        cursor = pos;
        insert_ix += 1;
    }

    out.extend_from_slice(&row_slice[cursor..]);
    Ok(out)
}

fn extract_style_id(cell_xml: &[u8]) -> Option<u32> {
    let style_pos = find_bytes(cell_xml, b" s=\"")?;
    let val_start = style_pos + 4;
    let val_end = find_bytes_from(cell_xml, b"\"", val_start)?;
    std::str::from_utf8(&cell_xml[val_start..val_end])
        .ok()?
        .parse()
        .ok()
}

fn extract_cell_col_idx(cell_xml: &[u8]) -> Option<usize> {
    let coord_pos = find_bytes(cell_xml, b" r=\"")?;
    let val_start = coord_pos + 4;
    let val_end = find_bytes_from(cell_xml, b"\"", val_start)?;
    let coord = std::str::from_utf8(&cell_xml[val_start..val_end]).ok()?;
    let letters_len = coord.find(|c: char| c.is_ascii_digit())?;
    Some(letters_to_col_idx(&coord[..letters_len]))
}

fn parse_row_number(src: &[u8], row_start: usize, row_tag_end: usize) -> Option<u32> {
    let rpos = find_bytes_from(src, b" r=\"", row_start)?;
    if rpos >= row_tag_end {
        return None;
    }
    let value_start = rpos + 4;
    let value_end = find_bytes_from(src, b"\"", value_start)?;
    std::str::from_utf8(&src[value_start..value_end])
        .ok()?
        .parse()
        .ok()
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

fn letters_to_col_idx(s: &str) -> usize {
    s.bytes().fold(0, |acc, b| {
        acc * 26 + (b.to_ascii_uppercase() - b'A' + 1) as usize
    }) - 1
}

pub fn scan<P: AsRef<Path>>(src: P) -> Result<Vec<String>> {
    let mut zip = zip::ZipArchive::new(File::open(src)?)?;
    let mut wb = zip
        .by_name("xl/workbook.xml")
        .context("workbook.xml not found")?;

    let mut wb_xml = Vec::with_capacity(wb.size() as usize);
    wb.read_to_end(&mut wb_xml)?;

    let mut reader = Reader::from_reader(wb_xml.as_slice());
    reader.config_mut().trim_text(true);

    let mut names = Vec::new();

    while let Ok(ev) = reader.read_event() {
        match ev {
            Event::Empty(ref e) | Event::Start(ref e)
                if e.name().as_ref().as_bytes() == b"sheet" =>
            {
                if let Some(attribute) = e
                    .attributes()
                    .with_checks(false)
                    .flatten()
                    .find(|a| a.key.as_ref().as_bytes() == b"name")
                {
                    names.push(
                        attribute
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)?
                            .into_owned(),
                    );
                }
            }
            Event::Eof => break,
            _ => {}
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
        let range = normalized.as_str();
        // 1. позиция после </sheetData>
        let sd_end = find_bytes(&self.sheet_xml, b"</sheetData>")
            .context("</sheetData> not found")?
            + "</sheetData>".len();

        let (insert_pos, created) = if let Some(pos) = find_bytes(&self.sheet_xml, b"<mergeCells") {
            // уже есть блок
            bump_count(&mut self.sheet_xml, b"<mergeCells", b"count=\"")?;
            let end = find_bytes_from(&self.sheet_xml, b"</mergeCells>", pos)
                .context("</mergeCells> not found")?;
            (end, false)
        } else {
            // нет блока – создаём
            let tpl = br#"<mergeCells count="0"></mergeCells>"#;
            self.sheet_xml.splice(sd_end..sd_end, tpl.iter().copied());
            (sd_end + tpl.len() - "</mergeCells>".len(), true)
        };

        // 2. сам <mergeCell>
        let tag = format!(r#"<mergeCell ref="{}"/>"#, range);
        self.sheet_xml
            .splice(insert_pos..insert_pos, tag.as_bytes().iter().copied());

        // 3. правим count (если блок создан только что)
        if created {
            bump_count(&mut self.sheet_xml, b"<mergeCells", b"count=\"")?;
        }
        Ok(())
    }
}

fn find_bytes(hay: &[u8], needle: &[u8]) -> Option<usize> {
    hay.windows(needle.len()).position(|w| w == needle)
}
fn find_bytes_from(hay: &[u8], needle: &[u8], start: usize) -> Option<usize> {
    hay[start..]
        .windows(needle.len())
        .position(|w| w == needle)
        .map(|p| p + start)
}

fn bump_count(xml: &mut Vec<u8>, tag: &[u8], attr: &[u8]) -> Result<()> {
    if let Some(pos) = find_bytes(xml, tag) {
        if let Some(a) = find_bytes_from(xml, attr, pos) {
            let start = a + attr.len();
            let end = find_bytes_from(xml, b"\"", start).unwrap();
            let mut num: u32 = std::str::from_utf8(&xml[start..end])?.parse()?;
            num += 1;
            xml.splice(start..end, num.to_string().as_bytes().iter().copied());
            return Ok(());
        }
    }
    Err(anyhow::anyhow!("attribute count not found"))
}

fn ensure_sheetdata_open_close(xml: &mut Vec<u8>) -> Result<()> {
    const SELF_CLOSING: &[u8] = b"<sheetData/>";
    if let Some(pos) = memchr::memmem::find(xml, SELF_CLOSING) {
        // заменяем на <sheetData></sheetData>
        let replacement = b"<sheetData></sheetData>";
        xml.splice(pos..pos + SELF_CLOSING.len(), replacement.iter().copied());
    }
    Ok(())
}
