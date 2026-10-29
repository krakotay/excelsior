//! Optional row spans for repeated scalar edits. The common XML parser builds
//! the index once; replacing a row shifts later spans without reparsing them.
use crate::{XlsxEditor, data, patch_row_xml, xml};
use anyhow::{Context, Result, ensure};
use std::ops::Range;

#[derive(Clone)]
pub(crate) struct RowIndex {
    offsets: Vec<(u32, Range<usize>)>,
    context: xml::RootContext,
}
impl RowIndex {
    fn build(bytes: &[u8]) -> Result<Self> {
        let data = xml::require(bytes, "sheetData")?;
        let mut offsets = Vec::new();
        let mut previous = 0;
        for row in data.children() {
            let row = row?;
            ensure!(
                row.is("row", xml::SHEET_NS),
                "Unexpected element in sheetData"
            );
            let number: u32 = row.attr("r")?.context("Row number missing")?.parse()?;
            ensure!(
                number > previous && number <= data::MAX_ROWS,
                "Invalid XLSX row order"
            );
            previous = number;
            offsets.push((number, row.full));
        }
        Ok(Self {
            offsets,
            context: data.context(),
        })
    }
}
impl XlsxEditor {
    /// Returns false for a missing row: the table writer handles its insertion.
    pub(crate) fn replace_indexed_cell(
        &mut self,
        column: usize,
        number: u32,
        value: &data::CellValue,
    ) -> Result<bool> {
        if self.row_index.is_none() {
            self.row_index = Some(RowIndex::build(&self.sheet_xml)?.into());
        }
        let index = self.row_index.as_ref().unwrap();
        let Ok(position) = index.offsets.binary_search_by_key(&number, |(n, _)| *n) else {
            return Ok(false);
        };
        let range = index.offsets[position].1.clone();
        let row = xml::root_in_root_context(&self.sheet_xml[range.clone()], &index.context)?;
        let replacement = patch_row_xml(&row, number, column, std::slice::from_ref(value))?;
        let delta = replacement.len() as isize - range.len() as isize;
        self.sheet_xml.splice(range.clone(), replacement);
        if delta != 0 {
            let index = self.row_index.as_mut().unwrap();
            index.offsets[position].1.end = range
                .end
                .checked_add_signed(delta)
                .context("Row index overflow")?;
            for (_, span) in &mut index.offsets[position + 1..] {
                span.start = span
                    .start
                    .checked_add_signed(delta)
                    .context("Row index overflow")?;
                span.end = span
                    .end
                    .checked_add_signed(delta)
                    .context("Row index overflow")?;
            }
        }
        self.sheet_data_end = self
            .sheet_data_end
            .and_then(|end| end.checked_add_signed(delta));
        if let Some((columns, rows)) = self.dimensions.get(&self.sheet_path).copied() {
            let stats = (columns.max(column + 1), rows.max(number));
            if stats != (columns, rows) {
                self.dimensions.insert(self.sheet_path.clone(), stats);
            }
        }
        Ok(true)
    }
}
