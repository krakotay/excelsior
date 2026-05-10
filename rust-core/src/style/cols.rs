use std::collections::BTreeMap;

use anyhow::{Context, Result};
use quick_xml::{events::Event, Reader};

use crate::XlsxEditor;
use crate::style::util::find_bytes_from;

const DEFAULT_COLUMN_WIDTH: f64 = 8.43;

#[derive(Clone, Debug, Default, PartialEq)]
struct ColProp {
    width: Option<f64>,
    style: Option<u32>,
    best_fit: bool,
    custom_width: bool,
    hidden: bool,
}

fn equal_props(a: &ColProp, b: &ColProp) -> bool {
    a.width == b.width
        && a.style == b.style
        && a.best_fit == b.best_fit
        && a.custom_width == b.custom_width
        && a.hidden == b.hidden
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

    fn ensure_cols_block(&mut self) -> Result<(usize, usize)> {
        use memchr::memmem;

        // уже есть?
        if let Some(start) = memmem::find(&self.sheet_xml, b"<cols") {
            let end = find_bytes_from(&self.sheet_xml, b"</cols>", start)
                .context("</cols> not found")?;
            return Ok((start, end + "</cols>".len()));
        }

        // найдём опорные точки
        let sd_pos = memmem::find(&self.sheet_xml, b"<sheetData")
            .context("<sheetData> not found on the current sheet")?;

        // гарантируем наличие <dimension .../> ПЕРЕД cols
        let (insert_after_dim, _inserted_len): (usize, usize) =
            if let Some(dim_start) = memmem::find(&self.sheet_xml, b"<dimension") {
                // конец тега <dimension ...>
                let dim_end =
                    find_bytes_from(&self.sheet_xml, b">", dim_start).context("<dimension> not closed")?
                        + 1;
                (dim_end, 0)
            } else {
                // создаём минимальный dimension прямо перед sheetData
                let dim_tag = r#"<dimension ref="A1"/>"#.as_bytes();
                self.sheet_xml
                    .splice(sd_pos..sd_pos, dim_tag.iter().copied());
                (sd_pos + dim_tag.len(), dim_tag.len())
            };

        // вставляем пустой блок cols СРАЗУ ПОСЛЕ <dimension ...>
        let block = b"<cols></cols>";
        self.sheet_xml
            .splice(insert_after_dim..insert_after_dim, block.iter().copied());

        // вернуть границы нового блока
        let start = insert_after_dim;
        let end = start + block.len();

        Ok((start, end))
    }

    fn read_cols_map(&self, cols_start: usize, cols_end: usize) -> Result<BTreeMap<u32, ColProp>> {
        let mut map = BTreeMap::new();
        let slice = &self.sheet_xml[cols_start..cols_end];
        let mut rdr = Reader::from_reader(slice);
        rdr.config_mut().trim_text(true);

        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"col" => {
                    let mut min = None;
                    let mut max = None;
                    let mut style = None;
                    let mut width = None;
                    let mut best_fit = false;
                    let mut custom_width = false;
                    let mut hidden = false;

                    for a in e.attributes().with_checks(false).flatten() {
                        let v = String::from_utf8_lossy(&a.value);
                        match a.key.as_ref() {
                            b"min" => min = Some(v.parse()?),
                            b"max" => max = Some(v.parse()?),
                            b"style" => style = v.parse().ok(),
                            b"width" => width = v.parse().ok(),
                            b"bestFit" => best_fit = v == "1" || v == "true",
                            b"customWidth" => custom_width = v == "1" || v == "true",
                            b"hidden" => hidden = v == "1" || v == "true",
                            _ => {}
                        }
                    }
                    let min = min.unwrap_or(1);
                    let max = max.unwrap_or(min);
                    let p = ColProp {
                        width,
                        style,
                        best_fit,
                        custom_width,
                        hidden,
                    };
                    for i in min..=max {
                        map.insert(i, p.clone());
                    }
                }
                Event::Eof => break,
                _ => {}
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
        self.sheet_xml.splice(cols_start..cols_end, out.bytes());
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
    s.push_str("/>");
    s
}
