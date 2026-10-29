//! Minimal XLSX creation shared by the Rust and Python APIs.
use anyhow::Result;
use std::{fs::File, path::Path};
use zip::write::SimpleFileOptions;

pub fn normalize_sheet_name(sheet_name: &str) -> Result<String> {
    let normalized = sheet_name.trim();
    if normalized.is_empty() {
        anyhow::bail!("Sheet name cannot be empty");
    }
    if normalized.chars().count() > 31 {
        anyhow::bail!("Sheet name cannot be longer than 31 characters");
    }
    if normalized
        .chars()
        .any(|c| matches!(c, ':' | '\\' | '/' | '?' | '*' | '[' | ']'))
    {
        anyhow::bail!("Sheet name contains forbidden Excel characters (: \\ / ? * [ ])");
    }
    if normalized.starts_with('\'') || normalized.ends_with('\'') {
        anyhow::bail!("Sheet name cannot begin or end with an apostrophe");
    }
    if normalized.chars().any(|c| c.is_control()) {
        anyhow::bail!("Sheet name cannot contain control characters");
    }
    Ok(normalized.to_owned())
}

fn xml_escape_attr(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('\'', "&apos;")
}

pub fn create_excel<P: AsRef<Path>>(path: P, sheet_name: &str) -> Result<()> {
    let path = path.as_ref();
    let sheet_name = normalize_sheet_name(sheet_name)?;
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)?;
    }

    let file = File::create(path)?;
    let mut zip = zip::ZipWriter::new(file);
    let options: SimpleFileOptions =
        SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#;

    let rels_root = ROOT_RELS;

    let workbook = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
  <sheets>\
    <sheet name=\"{}\" sheetId=\"1\" r:id=\"rId1\"/>\
  </sheets>\
</workbook>",
        xml_escape_attr(&sheet_name)
    );

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

    let sheet1 = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData></sheetData>
</worksheet>"#;

    let styles = BASE_STYLES;

    zip.start_file("[Content_Types].xml", options)?;
    std::io::Write::write_all(&mut zip, content_types.as_bytes())?;

    zip.start_file("_rels/.rels", options)?;
    std::io::Write::write_all(&mut zip, rels_root.as_bytes())?;

    zip.start_file("xl/workbook.xml", options)?;
    std::io::Write::write_all(&mut zip, workbook.as_bytes())?;

    zip.start_file("xl/_rels/workbook.xml.rels", options)?;
    std::io::Write::write_all(&mut zip, workbook_rels.as_bytes())?;

    zip.start_file("xl/worksheets/sheet1.xml", options)?;
    std::io::Write::write_all(&mut zip, sheet1.as_bytes())?;

    zip.start_file("xl/styles.xml", options)?;
    std::io::Write::write_all(&mut zip, styles.as_bytes())?;

    zip.finish()?;
    Ok(())
}

pub(crate) const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

pub(crate) const BASE_STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>"#;

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{XlsxEditor, scan};

    #[test]
    fn create_and_edit_without_a_template() -> Result<()> {
        let dir = tempfile::tempdir()?;
        let path = dir.path().join("nested/new.xlsx");
        let name = "Я".repeat(31);
        let mut editor = XlsxEditor::create(&path, &name)?;
        editor.append_row(["Hello", "42"])?;
        editor.save(&path)?;
        assert_eq!(scan(&path)?, vec![name.clone()]);
        assert_eq!(XlsxEditor::open(&path, &name)?.get_last_row_index("A")?, 1);
        Ok(())
    }

    #[test]
    fn invalid_sheet_names_do_not_truncate_the_destination() -> Result<()> {
        let dir = tempfile::tempdir()?;
        let path = dir.path().join("existing.xlsx");
        std::fs::write(&path, b"preserve")?;
        for name in ["", "bad/name", "'quoted", "bad\0name"] {
            assert!(create_excel(&path, name).is_err());
            assert_eq!(std::fs::read(&path)?, b"preserve");
        }
        Ok(())
    }
}
