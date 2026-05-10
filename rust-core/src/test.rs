#[cfg(test)]
use crate::{
    files_part::normalize_workbook_target,
    XlsxEditor, scan,
    style::{AlignSpec, HorizAlignment, VertAlignment},
};
#[cfg(test)]
use anyhow::Result;
#[test]
#[cfg(test)]
fn test_insert_table_at() -> Result<()> {
    let file_name = "../test/test.xlsx"; // Шаблон53. РД Выборка.xlsx result.xlsx
    let sheet_names: Vec<String> = scan(file_name)?;
    let data = vec![
        ["Name", "Score", "Status", "Number"],
        ["Alice", "123", "OK", "1"],
        ["Bob", "456", "FAIL", "2"],
    ];

    let mut app = XlsxEditor::open(file_name, &sheet_names[0])?;
    app.append_table_at("A4", data)?;
    app.save(file_name.to_owned() + "_appended.xlsx")?;

    Ok(())
}
#[test]
fn test_insert_cells() -> Result<()> {
    let file_name = "../test/test.xlsx"; // Шаблон53. РД Выборка.xlsx result.xlsx
    let sheet_names: Vec<String> = scan(file_name)?;
    let mut app = XlsxEditor::open(file_name, &sheet_names[0])?;
    app.set_cell("A25", "Hello")?;
    app.set_cell("B25", "World")?;
    app.set_cell("C25", "!")?;
    app.save(file_name.to_owned() + "_appended.xlsx")?;
    Ok(())
}
#[test]
fn test_get_last_row_index() -> Result<()> {
    let file_name = "../test/test_last_row_index.xlsx"; // Шаблон53. РД Выборка.xlsx result.xlsx
    let sheet_names: Vec<String> = scan(file_name)?;
    let app = XlsxEditor::open(file_name, &sheet_names[0])?;
    assert_eq!(app.get_last_row_index("A")?, 4);
    assert_eq!(app.get_last_row_index("B")?, 5);
    assert_eq!(app.get_last_row_index("C")?, 8);
    assert_eq!(app.get_last_row_index("D")?, 8);
    Ok(())
}
#[test]
fn test_get_last_roww_index() -> Result<()> {
    let file_name = "../test/test_last_row_index.xlsx";
    let sheet_names: Vec<String> = scan(file_name)?;
    let app = XlsxEditor::open(file_name, &sheet_names[0])?;
    assert_eq!(app.get_last_roww_index("A:D")?, vec![4, 5, 8, 8]);
    Ok(())
}

#[test]
fn add_new_worksheet() -> Result<()> {
    let file_name = "../test/test_new_ws.xlsx"; // fixed
    let new_file_name = "../test/test_new_ws_out.xlsx";

    let mut app = XlsxEditor::open(file_name, &scan(file_name)?[0])?;
    app.append_table_at("A1", [["Name", "Score", "Status", "Number"]])?;
    app.add_worksheet("NewSheet")?.set_cell("A1", "123")?;
    app.add_worksheet("NewSheet2")?
        .append_table_at("A1", [["Name", "Score", "Status", "Number"]])?;
    app.save(new_file_name)?;
    let sheet_names: Vec<String> = scan(new_file_name)?;

    println!("Sheet names: {:#?}", sheet_names);
    assert!(sheet_names.contains(&"NewSheet".to_owned()));
    assert!(sheet_names.contains(&"NewSheet2".to_owned()));
    Ok(())
}

#[test]
fn set_number_format() -> Result<()> {
    let file_name = "../test/numeric_format_test.xlsx";
    let file_name_out = "../test/numeric_format_test_out.xlsx";
    let sheet_names: Vec<String> = scan(file_name)?;
    let mut app = XlsxEditor::open(file_name, &sheet_names[0])?;
    app.set_number_format("A9", "#,##0.00")?;
    app.set_number_format("B3:C5", "#,##0.00")?;
    app.save(file_name_out)?;
    Ok(())
}

#[test]
fn style_patch_preserves_number_format() -> Result<()> {
    let file_name = "../test/numeric_format_test.xlsx";
    let mut app = XlsxEditor::open(file_name, "Sheet1")?;

    app.set_number_format("A9", "#,##0.00")?;
    let before_sid = app.cell_style_id("A9")?;
    let before = app.read_style_parts(before_sid)?;
    assert_eq!(before.num_fmt_code.as_deref(), Some("#,##0.00"));

    app.set_fill("A9", "FFCCCC")?;
    let after_sid = app.cell_style_id("A9")?;
    let after = app.read_style_parts(after_sid)?;
    assert_eq!(after.num_fmt_code.as_deref(), Some("#,##0.00"));
    assert!(after.fill.is_some());

    Ok(())
}

#[test]
fn column_number_format_preserves_existing_fill() -> Result<()> {
    let file_name = "../test/style_test.xlsx";
    let mut app = XlsxEditor::open(file_name, "Sheet1")?;

    app.set_fill("A2", "FFCCCC")?;
    let before = app.read_style_parts(app.cell_style_id("A2")?)?;
    assert!(before.fill.is_some());

    app.set_number_format("A:", "#,##0.00")?;
    let after = app.read_style_parts(app.cell_style_id("A2")?)?;
    assert!(after.fill.is_some());
    assert_eq!(after.num_fmt_code.as_deref(), Some("#,##0.00"));
    Ok(())
}

#[test]
fn row_number_format_preserves_existing_fill() -> Result<()> {
    let file_name = "../test/style_test.xlsx";
    let mut app = XlsxEditor::open(file_name, "Sheet1")?;

    app.set_fill("B2", "CCE5FF")?;
    let before = app.read_style_parts(app.cell_style_id("B2")?)?;
    assert!(before.fill.is_some());

    app.set_number_format("2:", "yyyy-mm-dd")?;
    let after = app.read_style_parts(app.cell_style_id("B2")?)?;
    assert!(after.fill.is_some());
    assert_eq!(after.num_fmt_code.as_deref(), Some("yyyy-mm-dd"));
    Ok(())
}
#[test]
fn set_style() -> Result<()> {
    let file_name = "../test/style_test.xlsx";
    let file_name_out = "../test/style_test_out.xlsx";

    let mut xl = XlsxEditor::open(file_name, "Sheet1")?;

    xl.set_fill("B14:B18", "FFFF00")?
        .set_font("D4:D8", "Arial", 12.0, true, false)?
        .set_fill("E4:E8", "FFCCCC")?
        .set_font("A1:C3", "Calibri", 10.0, false, true)?
        .set_fill("A1:C3", "FFFF00")?
        .merge_cells("B12:D12")?;

    xl.save(file_name_out)?;
    Ok(())
}
#[test]
fn set_column_number_format() -> Result<()> {
    let file_name = "../test/numeric_format_test.xlsx";
    let file_name_out = "../test/numeric_format_column_test.xlsx";

    let mut xl = XlsxEditor::open(file_name, "Sheet1")?;

    xl.set_number_format("A:", "#,##0.00")?;
    xl.set_number_format("B:", "#,##0.00")?;
    xl.set_number_format("C:", "#,##0.00")?;
    xl.set_number_format("G:", "#,##0.00")?;

    xl.save(file_name_out)?;
    Ok(())
}
#[test]
fn set_border() -> Result<()> {
    let file_name = "../test/style_test.xlsx";
    let file_name_out = "../test/style_test_out_borders.xlsx";

    let mut xl: XlsxEditor = XlsxEditor::open(file_name, "Sheet1")?;
    xl.set_border("A2:C3", "thin")?
        .set_fill("A2:C3", "FFCCCC")?
        .set_font("A2:C3", "Arial", 12.0, true, false)?
        .set_alignment(
            "A2:C3",
            &AlignSpec {
                horiz: Some(HorizAlignment::Center),

                vert: Some(VertAlignment::Bottom),
                wrap: true,
            },
        )?
        .merge_cells("A3:C3")?;
    xl.add_worksheet("Sheet2")?
        .set_border("A2:C3", "thin")?
        .set_fill("A2:C3", "FFCCCC")?
        .set_font("A2:C3", "Arial", 12.0, true, false)?
        .merge_cells("A3:C3")?;
    xl.save(file_name_out)?;
    Ok(())
}
#[test]
fn set_border_font_color() -> Result<()> {
    let file_name = "../test/style_test.xlsx";
    let file_name_out = "../test/style_test_out_borders_font_color.xlsx";

    let mut xl: XlsxEditor = XlsxEditor::open(file_name, "Sheet1")?;
    xl.append_table_at("A1", [["1", "2", "3"], ["1", "2", "3"], ["1", "2", "3"]])?;

    xl.set_font("D4:D8", "Arial", 12.0, true, false)?
        .set_border("A1:C3", "thin")?
        .set_font_with_alignment(
            "A1:C3",
            "Calibri",
            10.0,
            false,
            true,
            &AlignSpec {
                horiz: Some(HorizAlignment::Center),
                vert: None,
                wrap: true,
            },
        )?;
    xl.add_worksheet("Sheet2")?;
    xl.append_table_at("A1", [["1", "2", "3"], ["1", "2", "3"], ["1", "2", "3"]])?;
    xl.set_font_with_alignment(
        "A1:C3",
        "Calibri",
        10.0,
        false,
        true,
        &AlignSpec {
            horiz: Some(HorizAlignment::Center),
            vert: None,
            wrap: true,
        },
    )?
    .set_border("A1:C3", "thin")?;
    xl.save(file_name_out)?;
    Ok(())
}

#[test]
fn add_worksheet() -> Result<()> {
    let file_name = "../test/test.xlsx";
    let file_name_out = "../test/add_worksheets_test.xlsx";

    let mut xl: XlsxEditor = XlsxEditor::open(file_name, "Sheet1")?;
    xl.add_worksheet("Sheet2")?;
    xl.add_worksheet_at("TitleWS", 0)?;
    xl.add_worksheet("Sheet3")?;
    xl.with_worksheet("Sheet1")?
        .append_table_at("A1", [["1", "2", "3"], ["1", "2", "3"], ["1", "2", "3"]])?;

    xl.with_worksheet("Sheet2")?
        .append_table_at("A1", [["4", "5", "6"], ["7", "8", "9"]])?;

    xl.save(file_name_out)?;
    Ok(())
}

#[test]
fn inserted_sheet_keeps_name_based_open_and_switching() -> Result<()> {
    let file_name = "../test/edge_cases/generated/inn_filter_cases.xlsx";
    let tmp = tempfile::NamedTempFile::new()?;
    let out = tmp.path().with_extension("xlsx");
    let out2 = tmp.path().with_extension("roundtrip.xlsx");

    let mut xl = XlsxEditor::open(file_name, "дебет")?;
    xl.add_worksheet_at("Титульник", 0)?
        .with_worksheet("дебет")?
        .set_cell("A1", "patched-debit")?;
    xl.save(&out)?;

    let mut reopened = XlsxEditor::open(&out, "дебет")?;
    reopened.set_cell("B1", "opened-by-name")?;
    reopened.save(&out2)?;

    let mut zip = zip::ZipArchive::new(std::fs::File::open(&out2)?)?;
    let mut sheet1 = String::new();
    let mut sheet2 = String::new();
    {
        use std::io::Read;
        zip.by_name("xl/worksheets/sheet1.xml")?
            .read_to_string(&mut sheet1)?;
        zip.by_name("xl/worksheets/sheet2.xml")?
            .read_to_string(&mut sheet2)?;
    }

    assert!(sheet1.contains("patched-debit"));
    assert!(sheet1.contains("opened-by-name"));
    assert!(!sheet2.contains("patched-debit"));
    assert!(!sheet2.contains("opened-by-name"));
    Ok(())
}

#[test]
fn absolute_workbook_relationship_targets_are_supported() -> Result<()> {
    assert_eq!(
        normalize_workbook_target("/xl/worksheets/sheet1.xml"),
        "xl/worksheets/sheet1.xml"
    );
    assert_eq!(
        normalize_workbook_target("xl/worksheets/sheet1.xml"),
        "xl/worksheets/sheet1.xml"
    );
    assert_eq!(
        normalize_workbook_target("worksheets/sheet1.xml"),
        "xl/worksheets/sheet1.xml"
    );
    Ok(())
}

#[test]
fn saving_over_source_keeps_valid_zip() -> Result<()> {
    let file_name = "../test/edge_cases/generated/inn_filter_cases.xlsx";
    let tmp = tempfile::NamedTempFile::new()?;
    let path = tmp.path().with_extension("xlsx");
    std::fs::copy(file_name, &path)?;

    let mut xl = XlsxEditor::open(&path, "дебет")?;
    xl.set_cell("A1", "same-path-save-ok")?;
    xl.save(&path)?;

    let mut zip = zip::ZipArchive::new(std::fs::File::open(&path)?)?;
    let mut sheet1 = String::new();
    {
        use std::io::Read;
        zip.by_name("xl/worksheets/sheet1.xml")?
            .read_to_string(&mut sheet1)?;
    }

    assert!(sheet1.contains("same-path-save-ok"));
    Ok(())
}

#[test]
fn style_empty_inserted_sheet_before_cells_keeps_single_xml_root() -> Result<()> {
    let file_name = "../test/edge_cases/generated/inn_filter_cases.xlsx";
    let tmp = tempfile::NamedTempFile::new()?;
    let out = tmp.path().with_extension("xlsx");

    let mut xl = XlsxEditor::open(file_name, "дебет")?;
    xl.add_worksheet_at("Титульник", 0)?
        .set_border("B4:D12", "thin")?
        .set_column_width("A", 5.0)?
        .set_column_width("B", 15.0)?;
    xl.set_cell("B4", "РАБОЧИЕ ДОКУМЕНТЫ АУДИТОРА")?;
    xl.merge_cells("B4:D4")?;
    xl.save(&out)?;

    let mut zip = zip::ZipArchive::new(std::fs::File::open(&out)?)?;
    let mut title = String::new();
    {
        use std::io::Read;
        zip.by_name("xl/worksheets/sheet3.xml")?
            .read_to_string(&mut title)?;
    }

    assert!(title.trim_start().starts_with("<?xml"));
    assert_eq!(title.matches("<worksheet").count(), 1);
    assert!(title.contains("<sheetData>"));
    assert!(title.contains("<row r=\"4\">"));
    assert!(title.contains("РАБОЧИЕ ДОКУМЕНТЫ АУДИТОРА"));
    Ok(())
}

#[test]
fn column_text_format_then_title_styles_preserve_widths_and_styles() -> Result<()> {
    let file_name = "../test/edge_cases/generated/inn_filter_cases.xlsx";
    let tmp = tempfile::NamedTempFile::new()?;
    let out = tmp.path().with_extension("xlsx");

    let mut xl = XlsxEditor::open(file_name, "дебет")?;
    xl.add_worksheet("ИНН_аггдебет")?
        .append_table_at("A1", [["ИНН", "количество"], ["1233555", "1"], ["0000000000", "1"]])?;
    xl.set_number_format("A:", "@")?;
    xl.set_number_format("C:", "#,##0.00")?;
    xl.set_number_format("D:", "#,##0.00")?;
    xl.add_worksheet_at("Титульник", 0)?
        .set_border("B4:D12", "thin")?
        .set_column_width("A", 5.0)?
        .set_column_width("B", 15.0)?
        .set_column_width("C", 40.0)?
        .set_column_width("D", 15.0)?
        .set_column_width("E", 5.0)?;
    xl.set_cell("B4", "РАБОЧИЕ ДОКУМЕНТЫ АУДИТОРА")?;
    xl.set_font_with_alignment(
        "B4",
        "Times New Roman",
        16.0,
        true,
        false,
        &AlignSpec {
            horiz: Some(HorizAlignment::Center),
            vert: Some(VertAlignment::Center),
            wrap: false,
        },
    )?;
    xl.save(&out)?;

    let mut zip = zip::ZipArchive::new(std::fs::File::open(&out)?)?;
    let mut styles = String::new();
    let mut agg = String::new();
    let mut title = String::new();
    {
        use std::io::Read;
        zip.by_name("xl/styles.xml")?.read_to_string(&mut styles)?;
        zip.by_name("xl/worksheets/sheet3.xml")?
            .read_to_string(&mut agg)?;
        zip.by_name("xl/worksheets/sheet4.xml")?
            .read_to_string(&mut title)?;
    }

    assert!(agg.contains(r#"<col min="1" max="1" width="8.43" customWidth="1" style=""#));
    assert!(agg.contains(r#"<col min="3" max="4" width="8.43" customWidth="1" style=""#));
    assert!(agg.contains(r#"<c r="A2" s="1" t="inlineStr"><is><t>1233555</t></is></c>"#));
    assert!(agg.contains(r#"<c r="A3" s="1" t="inlineStr"><is><t>0000000000</t></is></c>"#));
    assert!(!agg.contains(r#"<c r="A2" s="1"><v>1233555</v></c>"#));
    assert!(styles.contains(r#"numFmtId="49""#));
    assert!(styles.contains(r#"<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#));
    assert!(!styles.contains(r#"formatCode="@""#));
    assert!(title.contains(r#"<col min="1" max="1" width="5" customWidth="1""#));
    assert!(title.contains(r#"<col min="3" max="3" width="40" customWidth="1""#));
    assert!(!title.contains(r#"width="0""#));
    assert!(title.contains(r#"<c r="B4" s=""#));
    Ok(())
}

#[test]
fn rename_del_worksheet() -> Result<()> {
    let file_name = "../test/test_rename_worksheets.xlsx";
    let file_name_out = "../test/rename_worksheets_out.xlsx";

    let mut xl: XlsxEditor = XlsxEditor::open(file_name, "Sheet1")?;
    xl.rename_worksheet("for_rename", "renamed")?
        .delete_worksheet("for_delete")?;

    xl.save(file_name_out)?;
    let new_sheets: Vec<String> = scan(file_name_out)?;
    println!("New sheets: {:?}", new_sheets);
    assert!(new_sheets.contains(&"renamed".to_owned()));
    assert!(!new_sheets.contains(&"for_delete".to_owned()));

    Ok(())
}
