use std::collections::BTreeMap;

use anyhow::{Context, Result};

use crate::XlsxEditor;

const DEFAULT_COLUMN_WIDTH: f64 = 8.43;

#[derive(Clone, Debug, Default, PartialEq)]
struct ColProp {
    width: Option<f64>,
    style: Option<u32>,
    best_fit: bool,
    custom_width: bool,
    hidden: bool,
    extra: BTreeMap<String, String>,
}

fn equal_props(a: &ColProp, b: &ColProp) -> bool {
    a.width == b.width
        && a.style == b.style
        && a.best_fit == b.best_fit
        && a.custom_width == b.custom_width
        && a.hidden == b.hidden
        && a.extra == b.extra
}

impl XlsxEditor {
    /// Главный публичный метод для столбца: точечное изменение + нормализация.
    pub(crate) fn set_column_properties(
        &mut self,
        col0: u32, // 0-based
        width: Option<f64>,
        style_id: Option<u32>,
    ) -> Result<()> {
        let (cols_start, cols_end) = self.ensure_cols_block()?;

        let mut cols_map = self.read_cols_map(cols_start, cols_end)?;
        let idx = col0 + 1; // храним в map 1-based для удобства
        let prop = cols_map.entry(idx).or_default();

        if let Some(w) = width {
            prop.width = Some(w);
            prop.custom_width = true;
        }
        if let Some(s) = style_id {
            prop.style = Some(s);
            if prop.width.is_none() {
                prop.width = Some(DEFAULT_COLUMN_WIDTH);
                prop.custom_width = true;
            }
        }

        self.write_cols_map(cols_start, cols_end, &cols_map)
    }

    pub(crate) fn set_column_style_ids(&mut self, styles: &[(u32, u32)]) -> Result<()> {
        if styles.is_empty() {
            return Ok(());
        }
        let (start, end) = self.ensure_cols_block()?;
        let mut map = self.read_cols_map(start, end)?;
        for &(column, style) in styles {
            let prop = map.entry(column + 1).or_default();
            prop.style = Some(style);
            if prop.width.is_none() {
                prop.width = Some(DEFAULT_COLUMN_WIDTH);
                prop.custom_width = true;
            }
        }
        self.write_cols_map(start, end, &map)
    }

    pub(crate) fn column_style_map(&self) -> Result<BTreeMap<usize, u32>> {
        let Some(cols) = crate::xml::header_child(&self.sheet_xml, "cols")? else {
            return Ok(BTreeMap::new());
        };
        Ok(self
            .read_cols_map(cols.full.start, cols.full.end)?
            .into_iter()
            .filter_map(|(col, prop)| prop.style.map(|style| ((col - 1) as usize, style)))
            .collect())
    }

    fn ensure_cols_block(&mut self) -> Result<(usize, usize)> {
        self.sheet_data_end = None;
        self.row_index = None;
        let context = crate::xml::root_context(&self.sheet_xml)?;
        if let Some(cols) = crate::xml::header_child(&self.sheet_xml, "cols")? {
            if cols.empty {
                let replacement = cols.with_content(&[])?;
                let begin = cols.full.start;
                let len = replacement.len();
                self.sheet_xml =
                    crate::xml::replace_element(&self.sheet_xml, &cols, &replacement).into();
                return Ok((begin, begin + len));
            }
            return Ok((cols.full.start, cols.full.end));
        }
        let data = crate::xml::require(&self.sheet_xml, "sheetData")?;
        let before = data.full.start;
        let block = context.fragment(b"<cols></cols>")?.into_owned();
        let length = block.len();
        self.sheet_xml.splice(before..before, block);
        Ok((before, before + length))
    }

    fn read_cols_map(&self, start: usize, end: usize) -> Result<BTreeMap<u32, ColProp>> {
        let root = crate::xml::root_in_root_context(
            &self.sheet_xml[start..end],
            &crate::xml::root_context(&self.sheet_xml)?,
        )?;
        let mut map = BTreeMap::new();
        for column in root.children() {
            let column = column?;
            if !column.is("col", crate::xml::SHEET_NS) {
                continue;
            }
            let min: u32 = column.attr("min")?.context("Column min missing")?.parse()?;
            let max: u32 = column
                .attr("max")?
                .map(|v| v.parse())
                .transpose()?
                .unwrap_or(min);
            anyhow::ensure!(
                min >= 1 && min <= max && max <= crate::data::MAX_COLUMNS as u32,
                "Invalid XLSX column range"
            );
            let prop = ColProp {
                width: column.attr("width")?.map(|v| v.parse()).transpose()?,
                style: column.attr("style")?.map(|v| v.parse()).transpose()?,
                best_fit: matches!(column.attr("bestFit")?.as_deref(), Some("1" | "true")),
                custom_width: matches!(column.attr("customWidth")?.as_deref(), Some("1" | "true")),
                hidden: matches!(column.attr("hidden")?.as_deref(), Some("1" | "true")),
                extra: column
                    .attributes()?
                    .into_iter()
                    .filter(|(name, _)| {
                        !matches!(
                            name.as_str(),
                            "min"
                                | "max"
                                | "width"
                                | "style"
                                | "bestFit"
                                | "customWidth"
                                | "hidden"
                        )
                    })
                    .collect(),
            };
            for index in min..=max {
                map.insert(index, prop.clone());
            }
        }
        Ok(map)
    }

    fn write_cols_map(
        &mut self,
        cols_start: usize,
        cols_end: usize,
        map: &BTreeMap<u32, ColProp>,
    ) -> Result<()> {
        // Сжимаем одинаковые проперти в диапазоны
        let mut out = String::with_capacity(256);
        out.push_str("<cols>");

        let mut it = map.iter().peekable();
        while let Some((&i, prop)) = it.next() {
            let mut j = i;
            while let Some(&(&k, prop2)) = it.peek() {
                if k == j + 1 && equal_props(prop, prop2) {
                    j = k;
                    it.next();
                } else {
                    break;
                }
            }
            out.push_str(&build_one_col_tag(i, j, prop));
        }

        out.push_str("</cols>");

        // подменяем всё содержимое старого блока
        let context = crate::xml::root_context(&self.sheet_xml)?;
        let out = context.fragment(out.as_bytes())?.into_owned();
        self.sheet_xml.splice(cols_start..cols_end, out);
        Ok(())
    }
}

fn build_one_col_tag(min: u32, max: u32, p: &ColProp) -> String {
    let mut s = format!(r#"<col min="{min}" max="{max}""#);
    let width = p.width.or_else(|| p.style.map(|_| DEFAULT_COLUMN_WIDTH));
    if let Some(w) = width {
        s.push_str(&format!(r#" width="{w}""#));
        if p.custom_width || p.style.is_some() {
            s.push_str(r#" customWidth="1""#);
        }
    }
    if let Some(st) = p.style {
        s.push_str(&format!(r#" style="{st}""#));
    }
    if p.best_fit {
        s.push_str(r#" bestFit="1""#);
    }
    if p.hidden {
        s.push_str(r#" hidden="1""#);
    }
    for (name, value) in &p.extra {
        s.push_str(&format!(
            r#" {name}="{}""#,
            crate::xml::escape_attribute(value)
        ));
    }
    s.push_str("/>");
    s
}
