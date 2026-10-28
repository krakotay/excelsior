//! Explicit cell types and a one-pass table writer. No dataframe dependencies.
use crate::{XlsxEditor, build_row_xml, parse_row_number, patch_row_xml};
use anyhow::{Context, Result, ensure};
use memchr::memmem;
use quick_xml::{Writer, events::BytesText};

pub const MAX_ROWS: u32 = 1_048_576;
pub const MAX_COLUMNS: usize = 16_384;

pub fn parse_column(column: &str) -> Result<usize> {
    ensure!(
        !column.is_empty() && column.bytes().all(|b| b.is_ascii_alphabetic()),
        "Invalid Excel column: {column}"
    );
    let mut n = 0usize;
    for b in column.bytes() {
        n = n
            .checked_mul(26)
            .and_then(|n| n.checked_add((b.to_ascii_uppercase() - b'A' + 1) as usize))
            .context("XLSX column limit exceeded")?;
        ensure!(n <= MAX_COLUMNS, "XLSX column limit exceeded");
    }
    Ok(n - 1)
}

pub fn parse_coord(coord: &str) -> Result<(usize, u32)> {
    let split = coord
        .find(|c: char| c.is_ascii_digit())
        .context("Invalid Excel coordinate")?;
    let col = parse_column(&coord[..split])?;
    let digits = &coord[split..];
    ensure!(
        digits.bytes().all(|b| b.is_ascii_digit()),
        "Invalid Excel coordinate"
    );
    let row: u32 = digits.parse().context("Invalid Excel row")?;
    ensure!((1..=MAX_ROWS).contains(&row), "XLSX row limit exceeded");
    Ok((col, row))
}

#[derive(Debug, Clone)]
pub enum CellValue {
    Blank,
    Text(String),
    Number(String),
    Boolean(bool),
    Formula(String),
}

impl CellValue {
    /// Legacy string API keeps its documented numeric/formula inference.
    pub fn infer(value: String) -> Self {
        if let Some(formula) = value.strip_prefix('=') {
            Self::Formula(formula.to_owned())
        } else if value.parse::<f64>().is_ok_and(f64::is_finite) {
            Self::Number(value)
        } else {
            Self::Text(value)
        }
    }
}

pub(crate) fn build_typed_cell_xml(
    coord: &str,
    value: &CellValue,
    sid: Option<u32>,
) -> Result<Vec<u8>> {
    match value {
        CellValue::Text(text) => ensure!(
            text.chars().all(|c| matches!(c, '\t' | '\n' | '\r')
                || (c >= '\u{20}' && c != '\u{fffe}' && c != '\u{ffff}')),
            "Text contains an invalid XML character"
        ),
        CellValue::Number(number) => ensure!(
            number.parse::<f64>().is_ok_and(f64::is_finite),
            "Excel numbers must be finite"
        ),
        _ => {}
    }
    let mut writer = Writer::new(Vec::new());
    let mut cell = writer.create_element("c").with_attribute(("r", coord));
    if let Some(sid) = sid {
        cell = cell.with_attribute(("s", sid.to_string().as_str()));
    }
    cell = match value {
        CellValue::Text(_) => cell.with_attribute(("t", "inlineStr")),
        CellValue::Boolean(_) => cell.with_attribute(("t", "b")),
        _ => cell,
    };
    cell.write_inner_content(|w| {
        match value {
            CellValue::Blank => {}
            CellValue::Text(text) => {
                // XML 1.0 control characters cannot be escaped by normal XML escaping.
                w.create_element("is").write_inner_content(|w| {
                    let mut t = w.create_element("t");
                    if crate::files_part::needs_xml_space_preserve(text) {
                        t = t.with_attribute(("xml:space", "preserve"));
                    }
                    t.write_text_content(BytesText::new(text))?;
                    Ok(())
                })?;
            }
            CellValue::Number(number) => {
                w.create_element("v")
                    .write_text_content(BytesText::new(number))?;
            }
            CellValue::Boolean(boolean) => {
                w.create_element("v")
                    .write_text_content(BytesText::new(if *boolean { "1" } else { "0" }))?;
            }
            CellValue::Formula(formula) => {
                w.create_element("f")
                    .write_text_content(BytesText::new(formula))?;
            }
        }
        Ok(())
    })?;
    Ok(writer.into_inner())
}

impl XlsxEditor {
    /// Consume one row at a time, merging with existing rows in a single pass.
    /// On an iterator/validation/XML error the original worksheet stays intact.
    pub fn write_typed_table_at<I>(&mut self, start: &str, rows: I) -> Result<()>
    where
        I: IntoIterator<Item = Result<Vec<CellValue>>>,
    {
        let (start_col, start_row) = parse_coord(start)?;
        // Appending must not scan or rebuild all previous rows on every call.
        if start_row > self.last_row {
            let mut tail = Vec::new();
            let mut last = self.last_row;
            for (offset, values) in rows.into_iter().enumerate() {
                let values = values?;
                let row_num = u32::try_from(offset)
                    .ok()
                    .and_then(|n| start_row.checked_add(n))
                    .context("XLSX row limit exceeded")?;
                ensure!(
                    row_num <= MAX_ROWS,
                    "XLSX row limit exceeded (including header)"
                );
                ensure!(
                    start_col + values.len() <= MAX_COLUMNS,
                    "XLSX column limit exceeded"
                );
                tail.extend_from_slice(&build_row_xml(row_num, start_col, &values)?);
                last = row_num;
            }
            if !tail.is_empty() {
                crate::ensure_sheetdata_open_close(&mut self.sheet_xml)?;
                let insert = memmem::rfind(&self.sheet_xml, b"</sheetData>")
                    .context("</sheetData> not found")?;
                self.sheet_xml.splice(insert..insert, tail);
                self.last_row = last;
            }
            return Ok(());
        }
        // Work on a temporary result, leaving the original buffer intact on error.
        let mut normalized = None;
        if memmem::find(&self.sheet_xml, b"<sheetData/>").is_some() {
            let mut xml = self.sheet_xml.clone();
            crate::ensure_sheetdata_open_close(&mut xml)?;
            normalized = Some(xml);
        }
        let src = normalized.as_deref().unwrap_or(&self.sheet_xml);
        let sd = memmem::find(src, b"<sheetData").context("<sheetData> not found")?;
        let content_start = memmem::find(&src[sd..], b">").context("Malformed sheetData")? + sd + 1;
        let content_end = memmem::find(&src[content_start..], b"</sheetData>")
            .context("</sheetData> not found")?
            + content_start;
        let mut out = Vec::with_capacity(src.len());
        out.extend_from_slice(&src[..content_start]);
        let mut cursor = content_start;
        let mut pending = next_existing_row(src, cursor, content_end)?;
        let mut max_row = self.last_row;
        let mut wrote = false;
        for (offset, values) in rows.into_iter().enumerate() {
            let values = values?;
            let row_num = u32::try_from(offset)
                .ok()
                .and_then(|n| start_row.checked_add(n))
                .context("XLSX row limit exceeded")?;
            ensure!(
                row_num <= MAX_ROWS,
                "XLSX row limit exceeded (including header)"
            );
            ensure!(
                start_col + values.len() <= MAX_COLUMNS,
                "XLSX column limit exceeded"
            );
            while let Some((number, _, end)) = pending {
                if number >= row_num {
                    break;
                }
                out.extend_from_slice(&src[cursor..end]);
                cursor = end;
                pending = next_existing_row(src, cursor, content_end)?;
            }
            if let Some((number, begin, end)) = pending {
                out.extend_from_slice(&src[cursor..begin]);
                cursor = begin;
                if number == row_num {
                    out.extend_from_slice(&patch_row_xml(
                        &src[begin..end],
                        row_num,
                        start_col,
                        &values,
                    )?);
                    cursor = end;
                    pending = next_existing_row(src, cursor, content_end)?;
                } else {
                    out.extend_from_slice(&build_row_xml(row_num, start_col, &values)?);
                }
            } else {
                out.extend_from_slice(&build_row_xml(row_num, start_col, &values)?);
            }
            wrote = true;
            max_row = max_row.max(row_num);
        }
        if wrote {
            out.extend_from_slice(&src[cursor..]);
            self.sheet_xml = out;
            self.last_row = max_row;
        }
        Ok(())
    }

    pub fn uses_1904_dates(&self) -> Result<bool> {
        let mut reader = quick_xml::Reader::from_reader(self.workbook_xml.as_slice());
        loop {
            match reader.read_event()? {
                quick_xml::events::Event::Start(e) | quick_xml::events::Event::Empty(e)
                    if e.name().as_ref().as_bytes() == b"workbookPr" =>
                {
                    for a in e.attributes() {
                        let a = a?;
                        if a.key.as_ref().as_bytes() == b"date1904" {
                            return Ok(matches!(a.value.as_ref(), "1" | "true"));
                        }
                    }
                }
                quick_xml::events::Event::Eof => return Ok(false),
                _ => {}
            }
        }
    }
}

fn next_existing_row(src: &[u8], from: usize, end: usize) -> Result<Option<(u32, usize, usize)>> {
    let Some(offset) = memmem::find(&src[from..end], b"<row ") else {
        return Ok(None);
    };
    let begin = from + offset;
    let tag_end = memmem::find(&src[begin..end], b">").context("Malformed row")? + begin;
    let number = parse_row_number(src, begin, tag_end).context("Missing/invalid row number")?;
    let finish = if src[tag_end - 1] == b'/' {
        tag_end + 1
    } else {
        memmem::find(&src[tag_end..end], b"</row>").context("</row> not found")? + tag_end + 6
    };
    Ok(Some((number, begin, finish)))
}
