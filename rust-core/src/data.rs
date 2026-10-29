//! Explicit cell types and a one-pass table writer. No dataframe dependencies.
use crate::{XlsxEditor, build_row_xml, patch_row_xml, xml};
use anyhow::{Context, Result, ensure};
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
    validate_cell(value)?;
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
        let mut dimensions = self.dimensions.get(&self.sheet_path).copied();
        // Appending must not scan or rebuild all previous rows on every call.
        if start_row > self.last_row {
            let empty_sheet = if self.last_row == 0 {
                Some(xml::require(&self.sheet_xml, "sheetData")?)
            } else {
                None
            };
            let mut tail = if let Some(data) = &empty_sheet {
                xml::begin_content(&self.sheet_xml, data)?
            } else {
                Vec::new()
            };
            let mut last = self.last_row;
            let mut wrote = false;
            let column_styles = self.column_style_map()?;
            let context = xml::root_context(&self.sheet_xml)?;
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
                let row_xml =
                    crate::build_styled_row_xml(row_num, start_col, &values, &column_styles)?;
                tail.extend_from_slice(&context.fragment(&row_xml)?);
                if let Some((columns, rows)) = &mut dimensions
                    && !values.is_empty()
                {
                    *columns = (*columns).max(start_col + values.len());
                    *rows = (*rows).max(row_num);
                }
                last = row_num;
                wrote = true;
            }
            if wrote && let Some(data) = &empty_sheet {
                xml::end_content(&self.sheet_xml, data, &mut tail)?;
                self.sheet_xml = tail.into();
                self.sheet_data_end = None;
                self.row_index = None;
                self.last_row = last;
                if let Some(stats) = dimensions {
                    self.dimensions.insert(self.sheet_path.clone(), stats);
                }
                return Ok(());
            }
            if wrote {
                let insert = self.sheet_data_insertion()?;
                let length = tail.len();
                self.sheet_xml.splice(insert..insert, tail);
                self.sheet_data_end = Some(insert + length);
                self.row_index = None;
                self.last_row = last;
                if let Some(stats) = dimensions {
                    self.dimensions.insert(self.sheet_path.clone(), stats);
                }
            }
            return Ok(());
        }
        let data = xml::require(&self.sheet_xml, "sheetData")?;
        let src = self.sheet_xml.as_slice();
        let mut rows_in_sheet = data.children();
        let mut pending = next_row(&mut rows_in_sheet)?;
        let mut content = xml::begin_content(src, &data)?;
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
            while pending.as_ref().is_some_and(|(n, _)| *n < row_num) {
                let (_, row) = pending.take().unwrap();
                content.extend_from_slice(row.bytes());
                pending = next_row(&mut rows_in_sheet)?;
            }
            if pending.as_ref().is_some_and(|(n, _)| *n == row_num) {
                let (_, row) = pending.take().unwrap();
                content.extend_from_slice(&patch_row_xml(&row, row_num, start_col, &values)?);
                pending = next_row(&mut rows_in_sheet)?;
            } else {
                content.extend_from_slice(&xml::child_fragment(
                    &data,
                    &build_row_xml(row_num, start_col, &values)?,
                )?);
            }
            if let Some((columns, rows)) = &mut dimensions
                && !values.is_empty()
            {
                *columns = (*columns).max(start_col + values.len());
                *rows = (*rows).max(row_num);
            }
            wrote = true;
            max_row = max_row.max(row_num);
        }
        if wrote {
            if let Some((_, row)) = pending {
                content.extend_from_slice(row.bytes());
            }
            for row in rows_in_sheet {
                content.extend_from_slice(row?.bytes());
            }
            xml::end_content(src, &data, &mut content)?;
            self.sheet_xml = content.into();
            self.sheet_data_end = None;
            self.row_index = None;
            self.last_row = max_row;
            if let Some(stats) = dimensions {
                self.dimensions.insert(self.sheet_path.clone(), stats);
            }
        }
        Ok(())
    }

    pub fn uses_1904_dates(&self) -> Result<bool> {
        let root = xml::root(&self.workbook_xml)?;
        let Some(properties) = root.child("workbookPr")? else {
            return Ok(false);
        };
        Ok(matches!(
            properties.attr("date1904")?.as_deref(),
            Some("1" | "true")
        ))
    }
}

fn next_row<'a>(elements: &mut xml::Elements<'a>) -> Result<Option<(u32, xml::Element<'a>)>> {
    if let Some(element) = elements.next() {
        let element = element?;
        anyhow::ensure!(
            element.is("row", xml::SHEET_NS),
            "Unexpected element in sheetData"
        );
        let number = element.attr("r")?.context("Row number missing")?.parse()?;
        return Ok(Some((number, element)));
    }
    Ok(None)
}

impl XlsxEditor {
    pub(crate) fn sheet_data_insertion(&mut self) -> Result<usize> {
        if let Some(end) = self.sheet_data_end {
            return Ok(end);
        }
        let data = xml::require(&self.sheet_xml, "sheetData")?;
        let end = if data.empty {
            let replacement = data.with_content(&[])?;
            let position = data.full.start + xml::root(&replacement)?.content.end;
            self.sheet_xml = xml::replace_element(&self.sheet_xml, &data, &replacement).into();
            position
        } else {
            data.content.end
        };
        self.sheet_data_end = Some(end);
        Ok(end)
    }
}

pub(crate) fn validate_cell(value: &CellValue) -> Result<()> {
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
    Ok(())
}
