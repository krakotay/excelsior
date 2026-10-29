//! Bounded-memory creation: rows are spooled as typed records; each completed
//! worksheet is encoded directly into ZIP. Existing templates use XlsxEditor.
use crate::{
    data::{self, CellValue},
    files_part::compression_options,
    workbook, xml,
};
use anyhow::{Context, Result, ensure};
use std::{
    collections::{BTreeMap, HashMap},
    io::{BufReader, BufWriter, Read, Seek, SeekFrom, Write},
    path::{Path, PathBuf},
};
use tempfile::NamedTempFile;
use zip::ZipWriter;

struct Sheet {
    name: String,
    spool: Option<BufWriter<NamedTempFile>>,
    last_row: u32,
    max_col: usize,
    formats: BTreeMap<usize, String>,
    widths: BTreeMap<usize, f64>,
}
impl Sheet {
    fn new(name: String, directory: &Path) -> Result<Self> {
        Ok(Self {
            name,
            spool: Some(BufWriter::with_capacity(
                64 * 1024,
                tempfile::Builder::new()
                    .prefix(".excelsior-rows-")
                    .tempfile_in(directory)?,
            )),
            last_row: 0,
            max_col: 0,
            formats: BTreeMap::new(),
            widths: BTreeMap::new(),
        })
    }
}
#[derive(Default)]
struct Styles {
    codes: Vec<String>,
    by_code: HashMap<String, u32>,
}
impl Styles {
    fn id(&mut self, format: &str) -> Result<u32> {
        if format == "General" {
            return Ok(0);
        }
        if let Some(id) = self.by_code.get(format) {
            return Ok(*id);
        }
        ensure!(
            self.codes.len() < 65_000,
            "Too many distinct number formats"
        );
        let id = self.codes.len() as u32 + 1;
        self.codes.push(format.to_owned());
        self.by_code.insert(format.to_owned(), id);
        Ok(id)
    }
    fn xml(&self) -> Result<Vec<u8>> {
        let mut formats = String::new();
        let mut xfs = String::new();
        let mut custom = 164;
        for code in &self.codes {
            let num = if let Some(id) = crate::style::builtin_num_fmt_id(code) {
                id
            } else {
                let id = custom;
                custom += 1;
                formats.push_str(&format!(
                    r#"<numFmt numFmtId="{id}" formatCode="{}"/>"#,
                    xml::escape_attribute(code)
                ));
                id
            };
            xfs.push_str(&format!(r#"<xf numFmtId="{num}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#));
        }
        let mut bytes = workbook::BASE_STYLES.as_bytes().to_vec();
        if !xfs.is_empty() {
            bytes = xml::append_child(&bytes, "cellXfs", xfs.as_bytes())?;
            let cell_xfs = xml::require(&bytes, "cellXfs")?;
            let tag = cell_xfs.set_attr("count", Some(&(self.codes.len() + 1).to_string()))?;
            let range = cell_xfs.open;
            bytes.splice(range, tag);
        }
        if !formats.is_empty() {
            let root = xml::root(&bytes)?;
            let before = root.require_child("fonts")?.full.start;
            bytes.splice(
                before..before,
                format!(r#"<numFmts count="{}">{formats}</numFmts>"#, custom - 164).bytes(),
            );
        }
        Ok(bytes)
    }
}

/// A single mutable writer. Saving finalizes it; rows cannot be edited/reopened.
/// The destination is replaced atomically only by save().
pub struct WorkbookWriter {
    destination: PathBuf,
    compression: u8,
    zip: Option<ZipWriter<NamedTempFile>>,
    completed: Option<NamedTempFile>,
    published: Option<crate::archive::Source>,
    active: Option<Sheet>,
    names: Vec<String>,
    styles: Styles,
    poisoned: Option<String>,
    aborted: bool,
}
impl WorkbookWriter {
    pub fn create<P: AsRef<Path>>(
        path: P,
        sheet_name: &str,
        compression_level: u8,
    ) -> Result<Self> {
        ensure!(
            compression_level <= 9,
            "compression_level must be between 0 and 9"
        );
        let name = workbook::normalize_sheet_name(sheet_name)?;
        let destination = std::path::absolute(path)?;
        let parent = destination.parent().context("Output path has no parent")?;
        std::fs::create_dir_all(parent)?;
        let archive = tempfile::Builder::new()
            .prefix(".excelsior-stream-")
            .suffix(".xlsx")
            .tempfile_in(parent)?;
        let active = Sheet::new(name, parent)?;
        Ok(Self {
            destination,
            compression: compression_level,
            zip: Some(ZipWriter::new(archive)),
            completed: None,
            published: None,
            active: Some(active),
            names: Vec::new(),
            styles: Styles::default(),
            poisoned: None,
            aborted: false,
        })
    }
    fn writable(&self) -> Result<()> {
        ensure!(!self.aborted, "WorkbookWriter was aborted");
        ensure!(
            self.poisoned.is_none(),
            "WorkbookWriter is unusable after I/O error: {:?}",
            self.poisoned
        );
        ensure!(
            self.zip.is_some(),
            "WorkbookWriter is finalized; create a new writer to write more rows"
        );
        Ok(())
    }
    pub fn next_row(&self) -> Result<u32> {
        self.writable()?;
        Ok(self
            .active
            .as_ref()
            .context("No active worksheet")?
            .last_row
            + 1)
    }
    pub fn add_worksheet(&mut self, name: &str) -> Result<()> {
        self.writable()?;
        let name = workbook::normalize_sheet_name(name)?;
        ensure!(
            !self
                .names
                .iter()
                .chain(self.active.iter().map(|s| &s.name))
                .any(|n| n.to_lowercase() == name.to_lowercase()),
            "Sheet `{name}` already exists"
        );
        // Allocate the next spool before sealing the previous worksheet.
        let next = Sheet::new(
            name,
            self.destination.parent().context("Output parent missing")?,
        )?;
        self.seal_sheet()?;
        self.active = Some(next);
        Ok(())
    }
    pub fn set_column_format(&mut self, column: usize, format: &str) -> Result<()> {
        self.writable()?;
        ensure!(column < data::MAX_COLUMNS, "XLSX column limit exceeded");
        data::validate_cell(&CellValue::Text(format.to_owned()))?;
        self.active
            .as_mut()
            .context("No active worksheet")?
            .formats
            .insert(column, format.to_owned());
        Ok(())
    }
    pub fn set_column_width(&mut self, column: usize, width: f64) -> Result<()> {
        self.writable()?;
        ensure!(column < data::MAX_COLUMNS, "XLSX column limit exceeded");
        ensure!(
            width.is_finite() && width > 0.0 && width <= 255.0,
            "Column width must be in range (0, 255]"
        );
        self.active
            .as_mut()
            .context("No active worksheet")?
            .widths
            .insert(column, width);
        Ok(())
    }
    /// One call is transactional: iterator/validation failures roll back the spool.
    pub fn write_typed_rows<I>(&mut self, start: &str, rows: I) -> Result<()>
    where
        I: IntoIterator<Item = Result<Vec<CellValue>>>,
    {
        self.writable()?;
        let (col, row) = data::parse_coord(start)?;
        let active = self.active.as_mut().context("No active worksheet")?;
        ensure!(
            row > active.last_row,
            "Streaming rows must follow previously written rows; overlapping writes require Editor"
        );
        let writer = active.spool.as_mut().context("Row spool missing")?;
        writer.flush()?;
        let checkpoint = writer.get_mut().as_file_mut().stream_position()?;
        let old_row = active.last_row;
        let old_col = active.max_col;
        let result = (|| -> Result<()> {
            for (offset, values) in rows.into_iter().enumerate() {
                let values = values?;
                let number = u32::try_from(offset)
                    .ok()
                    .and_then(|n| row.checked_add(n))
                    .context("XLSX row limit exceeded")?;
                ensure!(number <= data::MAX_ROWS, "XLSX row limit exceeded");
                ensure!(
                    col + values.len() <= data::MAX_COLUMNS,
                    "XLSX column limit exceeded"
                );
                for value in &values {
                    data::validate_cell(value)?;
                }
                writer.write_all(&number.to_le_bytes())?;
                writer.write_all(&(col as u16).to_le_bytes())?;
                writer.write_all(&(values.len() as u16).to_le_bytes())?;
                for value in &values {
                    encode_cell(writer, value)?;
                }
                active.last_row = number;
                if !values.is_empty() {
                    active.max_col = active.max_col.max(col + values.len());
                }
            }
            writer.flush()?;
            Ok(())
        })();
        if result.is_err() {
            let (mut file, _buffer) = active.spool.take().unwrap().into_parts();
            let rollback = (|| -> std::io::Result<()> {
                file.as_file_mut().set_len(checkpoint)?;
                file.as_file_mut().seek(SeekFrom::Start(checkpoint))?;
                Ok(())
            })();
            active.spool = Some(BufWriter::with_capacity(64 * 1024, file));
            active.last_row = old_row;
            active.max_col = old_col;
            if let Err(error) = rollback {
                self.poisoned = Some(error.to_string());
            }
        }
        result
    }
    fn seal_sheet(&mut self) -> Result<()> {
        let Some(sheet) = self.active.take() else {
            return Ok(());
        };
        let result = self.encode_sheet(sheet);
        if let Err(error) = &result {
            self.poisoned = Some(error.to_string());
        }
        result
    }
    fn encode_sheet(&mut self, mut sheet: Sheet) -> Result<()> {
        let mut ids = BTreeMap::new();
        for (&col, fmt) in &sheet.formats {
            ids.insert(col, self.styles.id(fmt)?);
        }
        let mut spool = sheet
            .spool
            .take()
            .context("Row spool missing")?
            .into_inner()
            .map_err(|e| e.into_error())?;
        spool.as_file_mut().seek(SeekFrom::Start(0))?;
        let zip = self.zip.as_mut().context("Writer finalized")?;
        let index = self.names.len() + 1;
        zip.start_file(
            format!("xl/worksheets/sheet{index}.xml"),
            compression_options(self.compression),
        )?;
        write!(
            zip,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{}">"#,
            xml::SHEET_NS
        )?;
        let reference = if sheet.last_row == 0 || sheet.max_col == 0 {
            "A1".to_owned()
        } else {
            format!(
                "A1:{}{}",
                crate::col_idx_to_letters(sheet.max_col - 1),
                sheet.last_row
            )
        };
        write!(zip, r#"<dimension ref="{reference}"/>"#)?;
        if !sheet.formats.is_empty() || !sheet.widths.is_empty() {
            zip.write_all(b"<cols>")?;
            let columns: std::collections::BTreeSet<_> = sheet
                .formats
                .keys()
                .chain(sheet.widths.keys())
                .copied()
                .collect();
            for column in columns {
                let index = column + 1;
                let style = ids.get(&column).copied().unwrap_or(0);
                let width = sheet.widths.get(&column).copied().unwrap_or(8.43);
                write!(
                    zip,
                    r#"<col min="{index}" max="{index}" width="{width}" customWidth="1" style="{style}"/>"#
                )?;
            }
            zip.write_all(b"</cols>")?;
        }
        zip.write_all(b"<sheetData>")?;
        let mut reader = BufReader::with_capacity(64 * 1024, spool);
        loop {
            let mut header = [0u8; 8];
            let read = reader.read(&mut header[..1])?;
            if read == 0 {
                break;
            }
            reader.read_exact(&mut header[1..])?;
            let row = u32::from_le_bytes(header[0..4].try_into()?);
            let col = u16::from_le_bytes(header[4..6].try_into()?) as usize;
            let length = u16::from_le_bytes(header[6..8].try_into()?) as usize;
            write!(zip, r#"<row r="{row}">"#)?;
            for offset in 0..length {
                let value = decode_cell(&mut reader)?;
                let column = col + offset;
                let address = format!("{}{}", crate::col_idx_to_letters(column), row);
                let sid = ids.get(&column).copied().filter(|id| *id != 0);
                zip.write_all(&data::build_typed_cell_xml(&address, &value, sid)?)?;
            }
            zip.write_all(b"</row>")?;
        }
        zip.write_all(b"</sheetData></worksheet>")?;
        self.names.push(sheet.name);
        Ok(())
    }
    fn finalize(&mut self) -> Result<()> {
        if self.completed.is_some() || self.published.is_some() {
            return Ok(());
        }
        self.writable()?;
        self.seal_sheet()?;
        let result = (|| -> Result<NamedTempFile> {
            let zip = self.zip.as_mut().context("Writer finalized")?;
            let mut sheets = String::new();
            let mut rels = String::new();
            let mut overrides = String::new();
            for (i, name) in self.names.iter().enumerate() {
                let id = i + 1;
                sheets.push_str(&format!(
                    r#"<sheet name="{}" sheetId="{id}" r:id="rId{id}"/>"#,
                    xml::escape_attribute(name)
                ));
                rels.push_str(&format!(r#"<Relationship Id="rId{id}" Type="{}/worksheet" Target="worksheets/sheet{id}.xml"/>"#,xml::DOC_REL_NS));
                overrides.push_str(&format!(r#"<Override PartName="/xl/worksheets/sheet{id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#));
            }
            rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="{}/styles" Target="styles.xml"/>"#,
                self.names.len() + 1,
                xml::DOC_REL_NS
            ));
            let workbook = format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{}" xmlns:r="{}"><sheets>{sheets}</sheets></workbook>"#,
                xml::SHEET_NS,
                xml::DOC_REL_NS
            );
            let rels = format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{}">{rels}</Relationships>"#,
                xml::REL_NS
            );
            let types = format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="{}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>{overrides}</Types>"#,
                xml::CT_NS
            );
            let styles = self.styles.xml()?;
            for (name, bytes) in [
                ("_rels/.rels", workbook::ROOT_RELS.as_bytes()),
                ("xl/workbook.xml", workbook.as_bytes()),
                ("xl/_rels/workbook.xml.rels", rels.as_bytes()),
                ("[Content_Types].xml", types.as_bytes()),
                ("xl/styles.xml", styles.as_slice()),
            ] {
                zip.start_file(name, compression_options(self.compression))?;
                zip.write_all(bytes)?;
            }
            let mut complete = self.zip.take().context("Writer finalized")?.finish()?;
            complete.flush()?;
            complete.as_file().sync_all()?;
            Ok(complete)
        })();
        match result {
            Ok(file) => {
                self.completed = Some(file);
                Ok(())
            }
            Err(error) => {
                self.poisoned = Some(error.to_string());
                Err(error)
            }
        }
    }
    /// Finalize once; repeated save-as copies the completed archive without RAM buffering.
    pub fn save(&mut self, path: Option<&Path>) -> Result<()> {
        self.finalize()?;
        let path = path.unwrap_or(&self.destination);
        let parent = path
            .parent()
            .filter(|p| !p.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        let mut temp = tempfile::Builder::new()
            .prefix(".excelsior-save-")
            .suffix(".xlsx")
            .tempfile_in(parent)?;
        if let Some(source) = &self.published {
            std::io::copy(&mut source.reader(), temp.as_file_mut())?;
        } else {
            let source = self
                .completed
                .as_mut()
                .context("Completed archive missing")?;
            source.as_file_mut().seek(SeekFrom::Start(0))?;
            std::io::copy(source.as_file_mut(), temp.as_file_mut())?;
        }
        if let Ok(metadata) = std::fs::metadata(path) {
            temp.as_file().set_permissions(metadata.permissions())?;
        }
        temp.as_file().sync_all()?;
        let file = temp.persist(path).map_err(|e| e.error)?;
        self.published = Some(crate::archive::Source::from_file(file)?);
        self.completed = None;
        Ok(())
    }
    pub fn abort(&mut self) {
        self.active = None;
        self.zip = None;
        self.completed = None;
        self.published = None;
        self.aborted = true;
    }
}
fn encode_cell(writer: &mut impl Write, value: &CellValue) -> Result<()> {
    let (kind, text) = match value {
        CellValue::Blank => (0, ""),
        CellValue::Text(t) => (1, t.as_str()),
        CellValue::Number(t) => (2, t.as_str()),
        CellValue::Boolean(v) => (3, if *v { "1" } else { "0" }),
        CellValue::Formula(t) => (4, t.as_str()),
    };
    let length = u32::try_from(text.len()).context("Cell is too large for row spool")?;
    writer.write_all(&[kind])?;
    writer.write_all(&length.to_le_bytes())?;
    writer.write_all(text.as_bytes())?;
    Ok(())
}
fn decode_cell(reader: &mut impl Read) -> Result<CellValue> {
    let mut header = [0u8; 5];
    reader.read_exact(&mut header)?;
    let length = u32::from_le_bytes(header[1..5].try_into()?) as usize;
    let mut bytes = vec![0; length];
    reader.read_exact(&mut bytes)?;
    let text = String::from_utf8(bytes)?;
    Ok(match header[0] {
        0 => CellValue::Blank,
        1 => CellValue::Text(text),
        2 => CellValue::Number(text),
        3 => CellValue::Boolean(text == "1"),
        4 => CellValue::Formula(text),
        _ => anyhow::bail!("Invalid row spool cell type"),
    })
}
