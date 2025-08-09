// use mimalloc::MiMalloc;

// #[global_allocator]
// static GLOBAL: MiMalloc = MiMalloc;
pub mod files_part;
use memchr::memmem;
mod polars_part;
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
use quick_xml::{Reader, Writer, events::Event};

use crate::{
    files_part::needs_xml_space_preserve,
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
struct XfParts {
    num_fmt_id: u32,
    font_id: Option<u32>,
    fill_id: Option<u32>,
    border_id: Option<u32>,
    align: Option<AlignSpec>,
}

struct StyleIndex {
    xfs: Vec<XfParts>, // index == style_id

    numfmt_by_code: HashMap<String, u32>,
    next_custom_numfmt: u32, // >=164

    font_by_key: HashMap<FontKey, u32>,
    fill_by_rgb: HashMap<String, u32>,   // RGB в верхнем регистре
    border_by_key: HashMap<String, u32>, // единый style для всех сторон

    xf_by_key: HashMap<StyleKey, u32>,

    fonts_count: u32,
    fills_count: u32,
    borders_count: u32,
}

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

/// Polars

/// Main
impl XlsxEditor {
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
        let sheet_names = scan(src.as_ref())?;
        let sheet_id = sheet_names
            .iter()
            .position(|n| n == sheet_name)
            .context(format!("Sheet '{}' not found", sheet_name))?
            + 1;
        println!("Sheet ID: {} with name {}", sheet_id, sheet_name);
        Self::open_sheet(src, sheet_id)
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
        self.last_row += 1;
        let row_num = self.last_row;
        let mut writer = quick_xml::Writer::new(Vec::new());

        writer
            .create_element("row")
            .with_attribute(("r", row_num.to_string().as_str()))
            .write_inner_content(|w| {
                let mut col = b'A';
                for val in cells {
                    let coord = format!("{}{}", col as char, row_num);
                    let val_str = val.to_string();
                    let is_formula = val_str.starts_with('=');
                    let is_number = !is_formula && val_str.parse::<f64>().is_ok();

                    let mut c_elem = w.create_element("c").with_attribute(("r", coord.as_str()));
                    if !is_number && !is_formula {
                        c_elem = c_elem.with_attribute(("t", "inlineStr"));
                    }
                    c_elem.write_inner_content(|w2| {
                        use quick_xml::events::BytesText;
                        if is_formula {
                            w2.create_element("f")
                                .write_text_content(BytesText::new(&val_str[1..]))?;
                        } else if !is_number {
                            w2.create_element("is").write_inner_content(|w3| {
                                let mut t = w3.create_element("t");
                                if needs_xml_space_preserve(&val_str) {
                                    t = t.with_attribute(("xml:space", "preserve"));
                                }
                                t.write_text_content(BytesText::new(&val_str))?;
                                Ok(())
                            })?;
                        } else {
                            w2.create_element("v")
                                .write_text_content(BytesText::new(&val_str))?;
                        }
                        Ok(())
                    })?;
                    col += 1;
                }
                Ok(())
            })?;

        let new_row_xml = writer.into_inner();

        // ИСПРАВЛЕНО: не делаем mem::take; вставляем по позиции из текущего буфера
        let pos = memchr::memmem::rfind(&self.sheet_xml, b"</sheetData>")
            .context("</sheetData> tag not found")?;
        self.sheet_xml.splice(pos..pos, new_row_xml);

        Ok(())
    }

    /// Appends multiple rows (a table) to the end of the current sheet.
    ///
    /// This function iterates through the provided rows, and for each row, it iterates through its cells.
    /// Each cell's value is converted to a string, and its type (number or inline string) is inferred.
    /// The new rows are then appended to the sheet's XML content.
    ///
    /// # Arguments
    /// * `rows` - An iterator over iterators of values that can be converted to strings, representing the rows and cells of the table.
    ///
    /// # Returns
    /// A `Result` indicating success or an `anyhow::Error` if the operation fails.
    pub fn append_table<R, I, S>(&mut self, rows: R) -> anyhow::Result<()>
    where
        R: IntoIterator<Item = I>,
        I: IntoIterator<Item = S>,
        S: ToString,
    {
        ensure_sheetdata_open_close(&mut self.sheet_xml)?;

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

        let mut bulk_rows_xml = Vec::<u8>::new();

        for row in rows {
            self.last_row += 1;
            let row_num = self.last_row;

            let mut writer = quick_xml::Writer::new(Vec::new());
            writer
                .create_element("row")
                .with_attribute(("r", row_num.to_string().as_str()))
                .write_inner_content(|w| {
                    for (col_idx, val) in row.into_iter().enumerate() {
                        let coord = format!("{}{}", col_idx_to_letters(col_idx), row_num);
                        let val_str = val.to_string();
                        let is_formula = val_str.starts_with('=');
                        let is_number = !is_formula && val_str.parse::<f64>().is_ok();

                        let mut c_elem =
                            w.create_element("c").with_attribute(("r", coord.as_str()));
                        if !is_number && !is_formula {
                            c_elem = c_elem.with_attribute(("t", "inlineStr"));
                        }
                        c_elem.write_inner_content(|w2| {
                            use quick_xml::events::BytesText;
                            if is_formula {
                                w2.create_element("f")
                                    .write_text_content(BytesText::new(&val_str[1..]))?;
                            } else if !is_number {
                                w2.create_element("is").write_inner_content(|w3| {
                                    let mut t = w3.create_element("t");
                                    if needs_xml_space_preserve(&val_str) {
                                        t = t.with_attribute(("xml:space", "preserve"));
                                    }
                                    t.write_text_content(BytesText::new(&val_str))?;
                                    Ok(())
                                })?;
                            } else {
                                w2.create_element("v")
                                    .write_text_content(BytesText::new(&val_str))?;
                            }
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;

            bulk_rows_xml.extend_from_slice(&writer.into_inner());
        }

        // ИСПРАВЛЕНО: считаем pos по self.sheet_xml и туда же вставляем
        let pos = memchr::memmem::rfind(&self.sheet_xml, b"</sheetData>")
            .context("</sheetData> tag not found")?;
        self.sheet_xml.splice(pos..pos, bulk_rows_xml);

        Ok(())
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
        ensure_sheetdata_open_close(&mut self.sheet_xml)?;

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

        let row_start_pos = start_coord
            .find(|c: char| c.is_ascii_digit())
            .context("invalid start coordinate – no digits")?;
        let col_letters = &start_coord[..row_start_pos];
        let start_col_idx = letters_to_col_idx(col_letters);
        let current_row_num: u32 = start_coord[row_start_pos..]
            .parse()
            .context("invalid row in start coordinate")?;

        let mut bulk_rows_xml = Vec::<u8>::new();

        for (row_offset, row) in rows.into_iter().enumerate() {
            let abs_row = current_row_num + row_offset as u32;
            if abs_row <= self.last_row {
                for (col_offset, val) in row.into_iter().enumerate() {
                    let coord = format!(
                        "{}{}",
                        col_idx_to_letters(start_col_idx + col_offset),
                        abs_row
                    );
                    self.set_cell(&coord, val)?;
                }
            } else {
                let mut writer = quick_xml::Writer::new(Vec::new());
                writer
                    .create_element("row")
                    .with_attribute(("r", abs_row.to_string().as_str()))
                    .write_inner_content(|w| {
                        for (col_offset, val) in row.into_iter().enumerate() {
                            let coord = format!(
                                "{}{}",
                                col_idx_to_letters(start_col_idx + col_offset),
                                abs_row
                            );
                            let val_str = val.to_string();
                            let is_formula = val_str.starts_with('=');
                            let is_number = !is_formula && val_str.parse::<f64>().is_ok();

                            let mut c_elem =
                                w.create_element("c").with_attribute(("r", coord.as_str()));
                            if !is_number && !is_formula {
                                c_elem = c_elem.with_attribute(("t", "inlineStr"));
                            }
                            c_elem.write_inner_content(|w2| {
                                use quick_xml::events::BytesText;
                                if is_formula {
                                    w2.create_element("f")
                                        .write_text_content(BytesText::new(&val_str[1..]))?;
                                } else if !is_number {
                                    w2.create_element("is").write_inner_content(|w3| {
                                        let mut t = w3.create_element("t");
                                        if needs_xml_space_preserve(&val_str) {
                                            t = t.with_attribute(("xml:space", "preserve"));
                                        }
                                        t.write_text_content(BytesText::new(&val_str))?;
                                        Ok(())
                                    })?;
                                } else {
                                    w2.create_element("v")
                                        .write_text_content(BytesText::new(&val_str))?;
                                }
                                Ok(())
                            })?;
                        }
                        Ok(())
                    })?;

                bulk_rows_xml.extend_from_slice(&writer.into_inner());
                self.last_row = abs_row;
            }
        }

        // ИСПРАВЛЕНО: больше никакого mem::take здесь
        let pos = memchr::memmem::rfind(&self.sheet_xml, b"</sheetData>")
            .context("</sheetData> tag not found")?;
        self.sheet_xml.splice(pos..pos, bulk_rows_xml);

        Ok(())
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
        use crate::files_part::needs_xml_space_preserve;
        // row number
        let row_start = coord
            .find(|c: char| c.is_ascii_digit())
            .context("invalid cell coordinate – no digits found")?;
        let row_num: u32 = coord[row_start..]
            .parse()
            .context("invalid row number in cell coordinate")?;

        let val_str = value.to_string();
        let is_formula = val_str.starts_with('=');
        let is_number = !is_formula && val_str.parse::<f64>().is_ok();

        // → собрать XML ячейки
        let mut cell_writer = Writer::new(Vec::new());
        {
            let mut c = cell_writer.create_element("c").with_attribute(("r", coord));
            if !is_number && !is_formula {
                c = c.with_attribute(("t", "inlineStr"));
            }
            c.write_inner_content(|w2| {
                use quick_xml::events::BytesText;
                if is_formula {
                    w2.create_element("f")
                        .write_text_content(BytesText::new(&val_str[1..]))?;
                } else if !is_number {
                    w2.create_element("is").write_inner_content(|w3| {
                        let mut t = w3.create_element("t");
                        if needs_xml_space_preserve(&val_str) {
                            t = t.with_attribute(("xml:space", "preserve"));
                        }
                        t.write_text_content(BytesText::new(&val_str))?;
                        Ok(())
                    })?;
                } else {
                    w2.create_element("v")
                        .write_text_content(BytesText::new(&val_str))?;
                }
                Ok(())
            })?;
        }
        let new_cell_xml = cell_writer.into_inner();

        // ——— устойчивый поиск ряда r="row_num"
        let src = &self.sheet_xml;
        let find_row = memmem::Finder::new(b"<row ");
        let find_gt = memmem::Finder::new(b">");
        let find_row_close = memmem::Finder::new(b"</row>");

        let mut i = 0usize;
        let mut target_row: Option<(usize, usize, usize)> = None; // (row_start, row_tag_end, row_close_end)
        let mut insert_before_row_with_r_gt: Option<usize> = None;

        while let Some(off) = find_row.find(&src[i..]) {
            let rs = i + off;
            let tag_end = match find_gt.find(&src[rs..]) {
                Some(p) => rs + p,
                None => break,
            };
            // r=".."
            if let Some(rpos) = find_bytes_from(src, b" r=\"", rs) {
                if rpos < tag_end {
                    let v0 = rpos + 4;
                    if let Some(v1) = find_bytes_from(src, b"\"", v0) {
                        if let Ok(r) = std::str::from_utf8(&src[v0..v1])
                            .unwrap_or("")
                            .parse::<u32>()
                        {
                            if r == row_num {
                                let close_rel = match find_row_close.find(&src[tag_end + 1..]) {
                                    Some(p) => p,
                                    None => break,
                                };
                                let close_end = tag_end + 1 + close_rel + "</row>".len();
                                target_row = Some((rs, tag_end, close_end));
                                break;
                            } else if insert_before_row_with_r_gt.is_none() && r > row_num {
                                insert_before_row_with_r_gt = Some(rs);
                                // не break — возможно ещё встретим точное совпадение роу выше
                            }
                        }
                    }
                }
            }
            i = tag_end + 1;
        }

        // если ряд найден — меняем/вставляем ячейку внутри
        if let Some((row_start_pos, row_tag_end, row_close_end)) = target_row {
            // Найдём <c ... r="coord">
            let mut row_slice = self.sheet_xml[row_start_pos..row_close_end].to_vec();

            // границы содержимого
            let content_start = (row_tag_end - row_start_pos) + 1;
            let content_end = row_slice.len() - "</row>".len();

            let find_c = memmem::Finder::new(b"<c");
            let find_gt_local = memmem::Finder::new(b">");

            let mut j = content_start;
            let mut cell_found = false;

            while j < content_end {
                let Some(cpos_rel) = find_c.find(&row_slice[j..content_end]) else {
                    break;
                };
                let cpos = j + cpos_rel;

                // защита от <col>, <cfRule> и т.п.: следующий символ после "<c"
                let next = *row_slice.get(cpos + 2).unwrap_or(&b'>');
                let is_cell = matches!(next, b' ' | b'>' | b'/' | b'r' | b's' | b't');
                let tag_end = match find_gt_local.find(&row_slice[cpos..]) {
                    Some(p) => cpos + p,
                    None => break,
                };

                if !is_cell {
                    j = tag_end + 1;
                    continue;
                }

                // r="..."
                if let Some(rpos) = find_bytes_from(&row_slice, b" r=\"", cpos) {
                    if rpos < tag_end {
                        let v0 = rpos + 4;
                        if let Some(v1) = find_bytes_from(&row_slice, b"\"", v0) {
                            if &row_slice[v0..v1] == coord.as_bytes() {
                                // полные границы ячейки: self-closing или с </c>
                                let self_closing = tag_end > cpos && row_slice[tag_end - 1] == b'/';
                                let cell_end = if self_closing {
                                    tag_end + 1
                                } else {
                                    let cc = find_bytes_from(&row_slice, b"</c>", tag_end + 1)
                                        .context("</c> missing")?;
                                    cc + 4
                                };
                                row_slice.splice(cpos..cell_end, new_cell_xml.iter().copied());
                                cell_found = true;
                                break;
                            }
                        }
                    }
                }
                j = tag_end + 1;
            }

            if !cell_found {
                // вставляем перед </row>, поддерживая сортировку по колонке
                // let insert_at = content_end;
                // можно пробежать по существующим c и найти первую с колонкой > нашей
                // (упрощённо: просто в конец строки)
                row_slice.splice(content_end..content_end, new_cell_xml.iter().copied());
            }

            // заменяем назад
            self.sheet_xml
                .splice(row_start_pos..row_close_end, row_slice);
        } else {
            // ряда нет — создаём и вставляем в правильное место
            let mut new_row = Vec::<u8>::with_capacity(64 + new_cell_xml.len());
            new_row.extend_from_slice(b"<row r=\"");
            new_row.extend_from_slice(row_num.to_string().as_bytes());
            new_row.extend_from_slice(b"\">");
            new_row.extend_from_slice(&new_cell_xml);
            new_row.extend_from_slice(b"</row>");

            let pos = insert_before_row_with_r_gt.unwrap_or_else(|| {
                memmem::rfind(&self.sheet_xml, b"</sheetData>").expect("</sheetData> tag not found")
            });
            self.sheet_xml.splice(pos..pos, new_row);
        }

        if row_num > self.last_row {
            self.last_row = row_num;
        }

        Ok(())
    }
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
            Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"sheet" => {
                if let Some(n) = e.attributes().with_checks(false).flatten().find_map(|a| {
                    (a.key.as_ref() == b"name")
                        .then(|| String::from_utf8_lossy(&a.value).into_owned())
                }) {
                    names.push(n);
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
