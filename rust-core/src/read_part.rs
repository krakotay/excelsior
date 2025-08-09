use crate::XlsxEditor;
use anyhow::{Result, bail};
use quick_xml::{Reader, events::Event};

impl XlsxEditor {
    /// Returns the last non-empty row index for the specified column or columns.
    ///
    /// The `columns` argument can be a single column such as "B" or multiple comma–separated
    /// columns such as "B,D". The function scans the sheet for the highest populated row
    /// across all specified columns and returns that 1-based row index. If no data is found
    /// in those columns, `Ok(0)` is returned.
    pub fn get_last_row_index(&self, columns: &str) -> Result<u32> {
        // Local helper to split coordinate like "C12" -> ("C", 12)
        fn split_coord(coord: &str) -> (String, u32) {
            let pos = coord
                .find(|c: char| c.is_ascii_digit())
                .unwrap_or(coord.len());
            let col = coord[..pos].to_ascii_uppercase();
            let row: u32 = coord[pos..].parse().unwrap_or(0);
            (col, row)
        }

        let targets: std::collections::HashSet<String> = columns
            .split(',')
            .map(|s| s.trim().to_ascii_uppercase())
            .collect();
        if targets.is_empty() {
            bail!("no columns supplied")
        }

        let mut reader = Reader::from_reader(self.sheet_xml.as_slice());
        reader.config_mut().trim_text(true);
        let mut last_row: u32 = 0;

        while let Ok(ev) = reader.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"c" => {
                    // locate the coordinate attribute r="A1"
                    if let Some(coord) = e.attributes().with_checks(false).flatten().find_map(|a| {
                        (a.key.as_ref() == b"r")
                            .then(|| String::from_utf8_lossy(&a.value).into_owned())
                    }) {
                        let (col, row) = split_coord(&coord);
                        if targets.contains(&col) && row > last_row {
                            last_row = row;
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(last_row)
    }

    /// Returns a vector with the last non-empty row indices for every column in the inclusive
    /// range like "A:E". The resulting vector has the same length as the number of columns in
    /// the range and is ordered left-to-right.
    ///
    /// Example: `get_last_roww_index("A:C")` might return `[10, 12, 7]`.
    pub fn get_last_roww_index(&self, range: &str) -> Result<Vec<u32>> {
        let parts: Vec<&str> = range.split(':').collect();
        if parts.len() != 2 {
            bail!("range must be in the form A:E")
        }
        // Reuse helpers from outer function
        fn letters_to_col_idx(s: &str) -> usize {
            s.bytes().fold(0, |acc, b| {
                acc * 26 + (b.to_ascii_uppercase() - b'A' + 1) as usize
            }) - 1
        }
        fn split_coord(coord: &str) -> (String, u32) {
            let pos = coord
                .find(|c: char| c.is_ascii_digit())
                .unwrap_or(coord.len());
            let col = coord[..pos].to_ascii_uppercase();
            let row: u32 = coord[pos..].parse().unwrap_or(0);
            (col, row)
        }

        let start = parts[0].trim().to_ascii_uppercase();
        let end = parts[1].trim().to_ascii_uppercase();

        let start_idx = letters_to_col_idx(&start);
        let end_idx = letters_to_col_idx(&end);
        if start_idx > end_idx {
            bail!("invalid range order")
        }
        let mut per_col_last: Vec<u32> = vec![0; end_idx - start_idx + 1];

        let mut reader = Reader::from_reader(self.sheet_xml.as_slice());
        reader.config_mut().trim_text(true);

        while let Ok(ev) = reader.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"c" => {
                    if let Some(coord) = e.attributes().with_checks(false).flatten().find_map(|a| {
                        (a.key.as_ref() == b"r")
                            .then(|| String::from_utf8_lossy(&a.value).into_owned())
                    }) {
                        let (col, row) = split_coord(&coord);
                        let idx = letters_to_col_idx(&col);
                        if idx >= start_idx && idx <= end_idx {
                            let vec_idx = idx - start_idx;
                            if row > per_col_last[vec_idx] {
                                per_col_last[vec_idx] = row;
                            }
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(per_col_last)
    }
}
