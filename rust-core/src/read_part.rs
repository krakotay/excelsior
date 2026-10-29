use crate::{XlsxEditor, data, xml};
use anyhow::{Context, Result, ensure};
impl XlsxEditor {
    pub fn get_last_row_index(&self, columns: &str) -> Result<u32> {
        let wanted = columns
            .split(',')
            .map(|c| data::parse_column(c.trim()))
            .collect::<Result<std::collections::HashSet<_>>>()?;
        let mut last = 0;
        self.visit_cell_coordinates(|col, row| {
            if wanted.contains(&col) {
                last = last.max(row);
            }
        })?;
        Ok(last)
    }
    pub fn get_last_roww_index(&self, range: &str) -> Result<Vec<u32>> {
        let (start, end) = range
            .split_once(':')
            .context("range must be in the form A:E")?;
        let (start, end) = (
            data::parse_column(start.trim())?,
            data::parse_column(end.trim())?,
        );
        ensure!(start <= end, "invalid range order");
        let mut last = vec![0; end - start + 1];
        self.visit_cell_coordinates(|col, row| {
            if col >= start && col <= end {
                last[col - start] = last[col - start].max(row);
            }
        })?;
        Ok(last)
    }
    pub(crate) fn visit_cell_coordinates(&self, mut visit: impl FnMut(usize, u32)) -> Result<()> {
        let sheet = xml::require(&self.sheet_xml, "sheetData")?;
        for row in sheet.children() {
            for cell in row?.children() {
                let cell = cell?;
                if cell.is("c", xml::SHEET_NS) {
                    let address = cell.attr("r")?.context("Cell coordinate missing")?;
                    let (col, row) = data::parse_coord(&address)?;
                    visit(col, row);
                }
            }
        }
        Ok(())
    }
}
