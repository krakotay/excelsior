//! Workbook parts and sheet lifecycle. All XML addressing uses crate::xml.
use crate::{XlsxEditor, archive, data, shared::Bytes, workbook, xml};
use anyhow::{Context, Result, ensure};
use std::{
    fs::File,
    io::{Read, Write},
    path::Path,
};
use zip::{CompressionMethod, ZipArchive, ZipWriter, write::SimpleFileOptions};

pub(crate) fn needs_xml_space_preserve(text: &str) -> bool {
    text.as_bytes().first().is_some_and(u8::is_ascii_whitespace)
        || text.as_bytes().last().is_some_and(u8::is_ascii_whitespace)
        || text.contains(['\n', '\r', '\t'])
}
pub(crate) fn normalize_workbook_target(target: &str) -> String {
    let target = target.trim_start_matches('/');
    let path = if target.starts_with("xl/") {
        target.to_owned()
    } else {
        format!("xl/{target}")
    };
    let mut normalized = Vec::new();
    for part in path.split('/') {
        match part {
            "." | "" => {}
            ".." => {
                normalized.pop();
            }
            _ => normalized.push(part),
        }
    }
    normalized.join("/")
}
fn read_part<R: Read + std::io::Seek>(zip: &mut ZipArchive<R>, name: &str) -> Result<Vec<u8>> {
    let mut part = zip
        .by_name(name)
        .with_context(|| format!("{name} not found in XLSX"))?;
    let mut bytes = Vec::with_capacity(part.size() as usize);
    part.read_to_end(&mut bytes)?;
    Ok(bytes)
}
fn sheet_rid(workbook: &[u8], name: &str) -> Result<String> {
    for sheet in xml::require(workbook, "sheets")?.children() {
        let sheet = sheet?;
        if sheet.is("sheet", xml::SHEET_NS) && sheet.attr("name")?.as_deref() == Some(name) {
            return Ok(sheet
                .attr_ns("id", xml::DOC_REL_NS)?
                .context("Sheet relationship id missing")?
                .into_owned());
        }
    }
    anyhow::bail!("Sheet `{name}` not found")
}
fn relationship<'a>(rels: &'a [u8], id: &str) -> Result<xml::Element<'a>> {
    for rel in xml::root(rels)?.children() {
        let rel = rel?;
        if rel.is("Relationship", xml::REL_NS) && rel.attr("Id")?.as_deref() == Some(id) {
            return Ok(rel);
        }
    }
    anyhow::bail!("Relationship `{id}` not found")
}
fn sheet_path(workbook: &[u8], rels: &[u8], name: &str) -> Result<String> {
    let id = sheet_rid(workbook, name)?;
    let rel = relationship(rels, &id)?;
    ensure!(
        rel.attr("TargetMode")?.as_deref() != Some("External"),
        "External worksheets cannot be edited"
    );
    Ok(normalize_workbook_target(
        &rel.attr("Target")?.context("Worksheet target missing")?,
    ))
}
pub(crate) fn sheet_path_for_name<P: AsRef<Path>>(path: P, name: &str) -> Result<String> {
    let mut zip = ZipArchive::new(File::open(path)?)?;
    let workbook = read_part(&mut zip, "xl/workbook.xml")?;
    let rels = read_part(&mut zip, "xl/_rels/workbook.xml.rels")?;
    sheet_path(&workbook, &rels, name)
}
fn last_row(sheet: &[u8]) -> Result<u32> {
    let mut last = 0;
    for row in xml::require(sheet, "sheetData")?.children() {
        let row = row?;
        if row.is("row", xml::SHEET_NS) {
            let n: u32 = row.attr("r")?.context("Row number missing")?.parse()?;
            ensure!((1..=data::MAX_ROWS).contains(&n), "XLSX row limit exceeded");
            last = last.max(n);
        }
    }
    Ok(last)
}

impl XlsxEditor {
    pub fn open_sheet<P: AsRef<Path>>(source: P, sheet_id: usize) -> Result<Self> {
        Self::open_sheet_path(source, format!("xl/worksheets/sheet{sheet_id}.xml"))
    }
    pub(crate) fn open_sheet_path<P: AsRef<Path>, S: Into<String>>(
        source: P,
        sheet_path: S,
    ) -> Result<Self> {
        let src_path = std::fs::canonicalize(source)?;
        let source = archive::Source::open(&src_path)?;
        let mut zip = ZipArchive::new(source.reader())?;
        let sheet_path = sheet_path.into();
        let sheet_xml =
            xml::ensure_default_namespace(read_part(&mut zip, &sheet_path)?, xml::SHEET_NS)?;
        let styles_xml =
            xml::ensure_default_namespace(read_part(&mut zip, "xl/styles.xml")?, xml::SHEET_NS)?;
        let workbook_xml =
            xml::ensure_default_namespace(read_part(&mut zip, "xl/workbook.xml")?, xml::SHEET_NS)?;
        let rels_xml = xml::ensure_default_namespace(
            read_part(&mut zip, "xl/_rels/workbook.xml.rels")?,
            xml::REL_NS,
        )?;
        let last_row = last_row(&sheet_xml)?;
        let mut dimensions = std::collections::HashMap::new();
        if last_row == 0 {
            dimensions.insert(sheet_path.clone(), (0, 0));
        }
        Ok(Self {
            src_path,
            source,
            dimensions: dimensions.into(),
            sheet_path,
            sheet_xml: sheet_xml.into(),
            styles_xml: styles_xml.into(),
            workbook_xml: workbook_xml.into(),
            rels_xml: rels_xml.into(),
            last_row,
            sheet_data_end: None,
            row_index: None,
            new_files: Default::default(),
            loaded_files: Default::default(),
            removed_files: Default::default(),
            styles_index: None,
        })
    }
    fn load_sheet(&self, path: &str) -> Result<Bytes> {
        if path == self.sheet_path {
            return Ok(self.sheet_xml.clone());
        }
        if let Some(bytes) = self
            .new_files
            .get(path)
            .or_else(|| self.loaded_files.get(path))
        {
            return Ok(bytes.clone());
        }
        let mut zip = ZipArchive::new(self.source.reader())?;
        Ok(xml::ensure_default_namespace(read_part(&mut zip, path)?, xml::SHEET_NS)?.into())
    }
    pub fn with_worksheet(&mut self, name: &str) -> Result<&mut Self> {
        let path = sheet_path(&self.workbook_xml, &self.rels_xml, name)?;
        if path == self.sheet_path {
            return Ok(self);
        }
        let bytes = self.load_sheet(&path)?;
        let last = last_row(&bytes)?;
        let old_path = std::mem::replace(&mut self.sheet_path, path.clone());
        let old_bytes = std::mem::replace(&mut self.sheet_xml, bytes);
        self.new_files.insert(old_path, old_bytes);
        self.new_files.remove(&path);
        self.loaded_files.remove(&path);
        self.last_row = last;
        self.sheet_data_end = None;
        self.row_index = None;
        Ok(self)
    }
    pub fn add_worksheet(&mut self, name: &str) -> Result<&mut Self> {
        self.add_worksheet_at(name, usize::MAX)
    }
    pub fn add_worksheet_at(&mut self, name: &str, index: usize) -> Result<&mut Self> {
        let name = workbook::normalize_sheet_name(name)?;
        let root = xml::root(&self.workbook_xml)?;
        let sheets = root.require_child("sheets")?;
        let entries = sheets.children().collect::<Result<Vec<_>>>()?;
        for sheet in &entries {
            ensure!(
                sheet
                    .attr("name")?
                    .as_deref()
                    .is_none_or(|s| s.to_lowercase() != name.to_lowercase()),
                "Sheet `{name}` already exists"
            );
        }
        let new_sid = entries
            .iter()
            .map(|s| {
                s.attr("sheetId")
                    .and_then(|v| v.context("sheetId missing")?.parse().map_err(Into::into))
            })
            .collect::<Result<Vec<u32>>>()?
            .into_iter()
            .max()
            .unwrap_or(0)
            + 1;
        let rel_root = xml::root(&self.rels_xml)?;
        let mut rid = 1;
        let mut used = std::collections::HashSet::new();
        for rel in rel_root.children() {
            if let Some(id) = rel?.attr("Id")? {
                used.insert(id.into_owned());
            }
        }
        while used.contains(&format!("rId{rid}")) {
            rid += 1;
        }
        let mut zip = ZipArchive::new(self.source.reader())?;
        let mut max_file = 0;
        for i in 0..zip.len() {
            let name = zip.by_index_raw(i)?.name()?.into_owned();
            max_file = max_file.max(sheet_file_number(&name));
        }
        for path in self
            .new_files
            .keys()
            .chain(self.removed_files.iter())
            .chain(std::iter::once(&self.sheet_path))
        {
            max_file = max_file.max(sheet_file_number(path));
        }
        let path = format!("xl/worksheets/sheet{}.xml", max_file + 1);
        let sheet_tag = format!(
            r#"<sheet name="{}" sheetId="{new_sid}" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{}" r:id="rId{rid}"/>"#,
            xml::escape_attribute(&name),
            xml::DOC_REL_NS
        );
        let mut content = Vec::new();
        let position = index.min(entries.len());
        for (i, sheet) in entries.iter().enumerate() {
            if i == position {
                content.extend_from_slice(sheet_tag.as_bytes());
            }
            content.extend_from_slice(sheet.bytes());
        }
        if position == entries.len() {
            content.extend_from_slice(sheet_tag.as_bytes());
        }
        let replacement = sheets.with_content(&content)?;
        let wb = xml::replace_element(&self.workbook_xml, &sheets, &replacement);
        let rel_tag = format!(
            r#"<Relationship Id="rId{rid}" Type="{}/worksheet" Target="worksheets/sheet{}.xml"/>"#,
            xml::DOC_REL_NS,
            max_file + 1
        );
        let rels = xml::append_child(&self.rels_xml, "Relationships", rel_tag.as_bytes())?;
        let old_path = std::mem::replace(&mut self.sheet_path, path);
        self.new_files
            .insert(old_path, std::mem::take(&mut self.sheet_xml));
        self.sheet_xml = empty_sheet().into();
        self.last_row = 0;
        self.sheet_data_end = None;
        self.row_index = None;
        self.dimensions.insert(self.sheet_path.clone(), (0, 0));
        self.workbook_xml = wb.into();
        self.rels_xml = rels.into();
        Ok(self)
    }
    pub fn rename_worksheet(&mut self, old: &str, new: &str) -> Result<&mut Self> {
        let new = workbook::normalize_sheet_name(new)?;
        let sheets = xml::require(&self.workbook_xml, "sheets")?;
        let mut target = None;
        for sheet in sheets.children() {
            let sheet = sheet?;
            let name = sheet
                .attr("name")?
                .context("Sheet name missing")?
                .into_owned();
            if name == old {
                target = Some(sheet);
            } else {
                ensure!(
                    name.to_lowercase() != new.to_lowercase(),
                    "Sheet `{new}` already exists"
                );
            }
        }
        let target = target.with_context(|| format!("Sheet `{old}` not found"))?;
        let tag = target.set_attr("name", Some(&new))?;
        let range = target.open;
        self.workbook_xml.splice(range, tag);
        Ok(self)
    }
    pub fn delete_worksheet(&mut self, name: &str) -> Result<&mut Self> {
        let id = sheet_rid(&self.workbook_xml, name)?;
        let rel = relationship(&self.rels_xml, &id)?;
        let path = normalize_workbook_target(&rel.attr("Target")?.context("Sheet target missing")?);
        let rels = xml::replace_element(&self.rels_xml, &rel, &[]);
        let sheets = xml::require(&self.workbook_xml, "sheets")?;
        let mut content = Vec::new();
        let mut first = None;
        for sheet in sheets.children() {
            let sheet = sheet?;
            let title = sheet
                .attr("name")?
                .context("Sheet name missing")?
                .into_owned();
            if title != name {
                if first.is_none() {
                    first = Some(title);
                }
                content.extend_from_slice(sheet.bytes());
            }
        }
        let replacement = sheets.with_content(&content)?;
        let workbook = xml::replace_element(&self.workbook_xml, &sheets, &replacement);
        let next = if self.sheet_path == path {
            first
                .as_ref()
                .map(|n| {
                    sheet_path(&workbook, &rels, n)
                        .and_then(|p| self.load_sheet(&p).map(|b| (p, b)))
                })
                .transpose()?
        } else {
            None
        };
        self.workbook_xml = workbook.into();
        self.rels_xml = rels.into();
        self.removed_files.insert(path.clone());
        self.new_files.remove(&path);
        self.loaded_files.remove(&path);
        self.dimensions.remove(&path);
        if self.sheet_path == path {
            if let Some((path, bytes)) = next {
                self.last_row = last_row(&bytes)?;
                self.sheet_path = path.clone();
                self.sheet_xml = bytes;
                self.new_files.remove(&path);
                self.loaded_files.remove(&path);
                self.sheet_data_end = None;
                self.row_index = None;
            } else {
                self.sheet_path.clear();
                self.sheet_xml = Default::default();
                self.add_worksheet("Sheet1")?;
                self.new_files.remove("");
            }
        }
        Ok(self)
    }
    pub fn save_in_place(&mut self) -> Result<()> {
        let path = self.src_path.clone();
        self.save(path)
    }
    pub fn save<P: AsRef<Path>>(&mut self, destination: P) -> Result<()> {
        self.save_with_compression(destination, 3)
    }
    pub fn save_with_compression<P: AsRef<Path>>(
        &mut self,
        destination: P,
        compression_level: u8,
    ) -> Result<()> {
        ensure!(
            compression_level <= 9,
            "compression_level must be between 0 and 9"
        );
        let path = destination.as_ref();
        let parent = path
            .parent()
            .filter(|p| !p.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        let temp = tempfile::Builder::new()
            .prefix(".excelsior-save-")
            .suffix(".xlsx")
            .tempfile_in(parent)?;
        self.write_archive(temp.path(), compression_level)?;
        if let Ok(metadata) = std::fs::metadata(path) {
            temp.as_file().set_permissions(metadata.permissions())?;
        }
        temp.as_file().sync_all()?;
        temp.persist(path).map_err(|e| e.error)?;
        Ok(())
    }
    fn write_archive(&mut self, path: &Path, compression_level: u8) -> Result<()> {
        let mut input = ZipArchive::new(self.source.reader())?;
        let mut output = ZipWriter::new(File::create(path)?);
        let ct = content_types(&read_part(&mut input, "[Content_Types].xml")?, self)?;
        let mut written = std::collections::HashSet::new();
        let mut dimensions = self.dimensions.clone();
        for index in 0..input.len() {
            let name = input.by_index_raw(index)?.name()?.into_owned();
            if self.removed_files.contains(&name) {
                continue;
            }
            if name == self.sheet_path {
                let stats = write_sheet(
                    &mut output,
                    &name,
                    &self.sheet_xml,
                    compression_level,
                    dimensions.get(&name).copied(),
                )?;
                dimensions.insert(name.clone(), stats);
            } else if name == "xl/workbook.xml" {
                write_part(&mut output, &name, &self.workbook_xml, compression_level)?;
            } else if name == "xl/_rels/workbook.xml.rels" {
                write_part(&mut output, &name, &self.rels_xml, compression_level)?;
            } else if name == "xl/styles.xml" {
                write_part(&mut output, &name, &self.styles_xml, compression_level)?;
            } else if name == "[Content_Types].xml" {
                write_part(&mut output, &name, &ct, compression_level)?;
            } else if let Some(bytes) = self.new_files.get(&name) {
                let stats = write_sheet(
                    &mut output,
                    &name,
                    bytes,
                    compression_level,
                    dimensions.get(&name).copied(),
                )?;
                dimensions.insert(name.clone(), stats);
            } else {
                let mut part = input.by_index(index)?;
                output.start_file(&name, compression_options(compression_level))?;
                std::io::copy(&mut part, &mut output)?;
            }
            written.insert(name);
        }
        for (name, bytes) in &self.new_files {
            if !written.contains(name)
                && !self.removed_files.contains(name)
                && *name != self.sheet_path
            {
                let stats = write_sheet(
                    &mut output,
                    name,
                    bytes,
                    compression_level,
                    dimensions.get(name).copied(),
                )?;
                dimensions.insert(name.clone(), stats);
            }
        }
        if !written.contains(&self.sheet_path) {
            let stats = write_sheet(
                &mut output,
                &self.sheet_path,
                &self.sheet_xml,
                compression_level,
                dimensions.get(&self.sheet_path).copied(),
            )?;
            dimensions.insert(self.sheet_path.clone(), stats);
        }
        output.finish()?;
        self.dimensions = dimensions;
        Ok(())
    }
}
fn sheet_file_number(path: &str) -> usize {
    path.strip_prefix("xl/worksheets/sheet")
        .and_then(|s| s.strip_suffix(".xml"))
        .and_then(|n| n.parse().ok())
        .unwrap_or(0)
}
pub(crate) fn empty_sheet() -> Vec<u8> {
    format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{}"><dimension ref="A1"/><sheetData></sheetData></worksheet>"#,xml::SHEET_NS).into_bytes()
}
fn content_types(original: &[u8], editor: &XlsxEditor) -> Result<Vec<u8>> {
    let original = xml::ensure_default_namespace(original.to_vec(), xml::CT_NS)?;
    let root = xml::root(&original)?;
    let mut content = Vec::new();
    let mut present = std::collections::HashSet::new();
    for entry in root.children() {
        let entry = entry?;
        if let Some(part) = entry.attr("PartName")? {
            let name = part.trim_start_matches('/');
            if editor.removed_files.contains(name) {
                continue;
            }
            present.insert(name.to_owned());
        }
        content.extend_from_slice(entry.bytes());
    }
    for name in editor
        .new_files
        .keys()
        .chain(std::iter::once(&editor.sheet_path))
    {
        if !present.contains(name)
            && name.starts_with("xl/worksheets/")
            && !editor.removed_files.contains(name)
        {
            content.extend_from_slice(format!(r#"<Override xmlns="http://schemas.openxmlformats.org/package/2006/content-types" PartName="/{}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,xml::escape_attribute(name)).as_bytes());
            present.insert(name.clone());
        }
    }
    root.with_content(&content)
}
pub(crate) fn compression_options(level: u8) -> SimpleFileOptions {
    SimpleFileOptions::default()
        .compression_method(if level == 0 {
            CompressionMethod::Stored
        } else {
            CompressionMethod::Deflated
        })
        .compression_level(if level == 0 { None } else { Some(level as i64) })
}
fn write_part(output: &mut ZipWriter<File>, name: &str, bytes: &[u8], level: u8) -> Result<()> {
    output.start_file(name, compression_options(level))?;
    output.write_all(bytes)?;
    Ok(())
}
fn write_sheet(
    output: &mut ZipWriter<File>,
    name: &str,
    bytes: &[u8],
    level: u8,
    known: Option<(usize, u32)>,
) -> Result<(usize, u32)> {
    let context = xml::root_context(bytes)?;
    let (mut max_col, mut max_row) = (0, 0);
    if let Some((columns, rows)) = known {
        max_col = columns.saturating_sub(1);
        max_row = rows;
    } else {
        let data = xml::require(bytes, "sheetData")?;
        for row in data.children() {
            for cell in row?.children() {
                let cell = cell?;
                if cell.is("c", xml::SHEET_NS) {
                    let coord = cell.attr("r")?.context("Cell coordinate missing")?;
                    let (col, row) = data::parse_coord(&coord)?;
                    max_col = max_col.max(col);
                    max_row = max_row.max(row);
                }
            }
        }
    }
    let stats = (if max_row == 0 { 0 } else { max_col + 1 }, max_row);
    let value = if max_row == 0 {
        "A1".to_owned()
    } else {
        format!("A1:{}{}", crate::col_idx_to_letters(max_col), max_row)
    };
    let (range, replacement) = if let Some(dimension) = xml::header_child(bytes, "dimension")? {
        if dimension.attr("ref")?.as_deref() == Some(value.as_str()) {
            write_part(output, name, bytes, level)?;
            return Ok(stats);
        }
        (
            dimension.open.clone(),
            dimension.set_attr("ref", Some(&value))?,
        )
    } else {
        let position = xml::header_child(bytes, "sheetPr")?
            .map(|e| e.full.end)
            .unwrap_or(context.content_start);
        (
            position..position,
            context
                .fragment(format!(r#"<dimension ref="{value}"/>"#).as_bytes())?
                .into_owned(),
        )
    };
    output.start_file(name, compression_options(level))?;
    output.write_all(&bytes[..range.start])?;
    output.write_all(&replacement)?;
    output.write_all(&bytes[range.end..])?;
    Ok(stats)
}
