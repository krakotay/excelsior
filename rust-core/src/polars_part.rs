use crate::{XlsxEditor, find_bytes_from};
use anyhow::{Context, Result};
use polars_core::prelude::*;

impl XlsxEditor {
    pub fn with_polars(&mut self, df: &DataFrame, start_cell: Option<&str>) -> Result<()> {
        use crate::style::{col_letter, split_coord};
        use memchr::memmem;
        use quick_xml::{Writer, events::BytesText};

        // ---------- 0. Координаты ----------
        let start_coord = start_cell.unwrap_or("A1");
        let (base_col, first_row) = split_coord(start_coord); // (col0, row0)
        let last_row = first_row + df.height() as u32; // хедер + N строк

        // ---------- 1. Метаданные столбцов (как у тебя) ----------
        struct ColMeta {
            is_number: bool,
            style_id: Option<u32>,
            conv: Box<dyn Fn(AnyValue) -> String>,
        }
        let mut cols_meta = Vec::<ColMeta>::with_capacity(df.width());
        for s in df.get_columns() {
            match s.dtype() {
                DataType::String => cols_meta.push(ColMeta {
                    is_number: false,
                    style_id: None,
                    conv: Box::new(|v| match v {
                        AnyValue::String(s) => s.to_string(),
                        _ => {
                            let mut t = v.to_string();
                            if t.starts_with('"') && t.ends_with('"') {
                                t.truncate(t.len() - 1);
                                t.remove(0);
                            }
                            t
                        }
                    }),
                }),
                DataType::Int8
                | DataType::Int16
                | DataType::Int32
                | DataType::Int64
                | DataType::UInt8
                | DataType::UInt16
                | DataType::UInt32
                | DataType::UInt64
                | DataType::Float32
                | DataType::Float64 => cols_meta.push(ColMeta {
                    is_number: true,
                    style_id: None,
                    conv: Box::new(|v| v.to_string()),
                }),
                _ => cols_meta.push(ColMeta {
                    is_number: false,
                    style_id: None,
                    conv: Box::new(|v| v.to_string()),
                }),
            }
        }

        // заранее кэш буквенных колонок (чтобы не дергать format! на каждую ячейку)
        let mut col_letters: Vec<String> = Vec::with_capacity(df.width());
        for i in 0..df.width() {
            col_letters.push(col_letter(base_col + i as u32));
        }

        // ---------- 2. Генерим новые <row> одним буфером ----------
        let mut bulk_rows_xml = Vec::<u8>::with_capacity(64 * 1024);

        // 2.1 заголовок
        let mut cur_row = first_row;
        {
            let mut w = Writer::new(Vec::new());
            w.create_element("row")
                .with_attribute(("r", cur_row.to_string().as_str()))
                .write_inner_content(|wr| {
                    for (col_idx, s) in df.get_columns().iter().enumerate() {
                        let mut coord = String::with_capacity(8);
                        coord.push_str(&col_letters[col_idx]);
                        coord.push_str(&cur_row.to_string());

                        let c = wr
                            .create_element("c")
                            .with_attribute(("r", coord.as_str()))
                            .with_attribute(("t", "inlineStr"));

                        c.write_inner_content(|w2| {
                            w2.create_element("is").write_inner_content(|w3| {
                                w3.create_element("t")
                                    .write_text_content(BytesText::new(s.name()))?;
                                Ok(())
                            })?;
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;
            bulk_rows_xml.extend_from_slice(&w.into_inner());
            cur_row += 1;
        }

        // 2.2 данные
        for ridx in 0..df.height() {
            let mut w = Writer::new(Vec::new());
            w.create_element("row")
                .with_attribute(("r", cur_row.to_string().as_str()))
                .write_inner_content(|wr| {
                    for (col_idx, s) in df.get_columns().iter().enumerate() {
                        let val = s.get(ridx).unwrap_or(AnyValue::Null);
                        let meta = &cols_meta[col_idx];

                        enum Kind {
                            Blank,
                            Num(String),
                            Str(String),
                        }
                        let kind = match val {
                            AnyValue::Null => Kind::Blank,
                            AnyValue::Float64(x) if x.is_finite() => Kind::Num(x.to_string()),
                            AnyValue::Float32(x) if x.is_finite() => Kind::Num(x.to_string()),
                            _ => {
                                if meta.is_number {
                                    Kind::Num((meta.conv)(val))
                                } else {
                                    Kind::Str((meta.conv)(val))
                                }
                            }
                        };

                        let mut coord = String::with_capacity(8);
                        coord.push_str(&col_letters[col_idx]);
                        coord.push_str(&cur_row.to_string());

                        let mut c = wr.create_element("c").with_attribute(("r", coord.as_str()));
                        if let Some(sid) = meta.style_id {
                            c = c.with_attribute(("s", sid.to_string().as_str()));
                        }
                        if matches!(kind, Kind::Str(_)) {
                            c = c.with_attribute(("t", "inlineStr"));
                        }

                        c.write_inner_content(|w2| {
                            match kind {
                                Kind::Blank => {}
                                Kind::Num(txt) => {
                                    w2.create_element("v")
                                        .write_text_content(BytesText::new(&txt))?;
                                }
                                Kind::Str(txt) => {
                                    w2.create_element("is").write_inner_content(|w3| {
                                        w3.create_element("t")
                                            .write_text_content(BytesText::new(&txt))?;
                                        Ok(())
                                    })?;
                                }
                            }
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;
            bulk_rows_xml.extend_from_slice(&w.into_inner());
            cur_row += 1;
        }

        // ---------- 3. Однопроходная замена блоков строк в <sheetData> ----------
        let src = std::mem::take(&mut self.sheet_xml);
        let sd_open = memmem::find(&src, b"<sheetData>").context("<sheetData> tag not found")?
            + "<sheetData>".len();
        let sd_close =
            memmem::rfind(&src, b"</sheetData>").context("</sheetData> tag not found")?;

        // заранее: место до <sheetData>
        let mut dst = Vec::with_capacity(
            src.len()
                + bulk_rows_xml
                    .len()
                    .saturating_sub((last_row - first_row + 1) as usize * 8),
        );
        dst.extend_from_slice(&src[..sd_open]);

        // сканер по строкам внутри sheetData
        let mut i = sd_open;
        let row_finder = memmem::Finder::new(b"<row");
        let mut inserted = false;

        while let Some(off) = row_finder.find(&src[i..]) {
            let row_open = i + off;
            if row_open >= sd_close {
                break;
            }

            // валиден ли <row ...> (а не <rower> и т.п.)
            let next = *src.get(row_open + 4).unwrap_or(&b'>');
            if next != b' ' && next != b'>' {
                dst.extend_from_slice(&src[i..row_open + 4]);
                i = row_open + 4;
                continue;
            }

            // границы тега и блока
            let open_end = find_bytes_from(&src, b">", row_open).context("malformed <row>")? + 1;
            let row_end = find_bytes_from(&src, b"</row>", open_end).context("</row> missing")?
                + "</row>".len();

            // быстрый парс r=".."
            let mut row_num: Option<u32> = None;
            if let Some(r_attr) = find_bytes_from(&src, b" r=\"", row_open) {
                if r_attr < open_end {
                    let v0 = r_attr + 4;
                    if let Some(v1) = find_bytes_from(&src, b"\"", v0) {
                        row_num = std::str::from_utf8(&src[v0..v1])
                            .ok()
                            .and_then(|s| s.parse().ok());
                    }
                }
            }
            // fallback: по первой ячейке
            if row_num.is_none() {
                if let Some(rpos) = find_bytes_from(&src, b" r=\"", open_end) {
                    if rpos < row_end {
                        let v0 = rpos + 4;
                        if let Some(v1) = find_bytes_from(&src, b"\"", v0) {
                            let s = &src[v0..v1];
                            // отделяем хвост цифр
                            let digits_start = s
                                .iter()
                                .rposition(|&b| !(b as char).is_ascii_digit())
                                .map(|p| p + 1)
                                .unwrap_or(0);
                            row_num = std::str::from_utf8(&s[digits_start..])
                                .ok()
                                .and_then(|x| x.parse().ok());
                        }
                    }
                }
            }

            if let Some(n) = row_num {
                if !inserted && n >= first_row {
                    dst.extend_from_slice(&bulk_rows_xml);
                    inserted = true;
                }
                if n < first_row || n > last_row {
                    dst.extend_from_slice(&src[row_open..row_end]);
                }
            } else {
                // не удалось распарсить — на всякий случай копируем как есть
                dst.extend_from_slice(&src[row_open..row_end]);
            }

            i = row_end;
        }

        if !inserted {
            // не нашли ряд >= first_row — вставляем перед </sheetData>
            dst.extend_from_slice(&bulk_rows_xml);
        }

        // хвост внутри sheetData и дальше документ
        dst.extend_from_slice(&src[i..sd_close]);
        dst.extend_from_slice(&src[sd_close..]);

        self.sheet_xml = dst;

        // ---------- 4. Обновляем last_row и dimension ----------
        self.last_row = last_row; // NOTE: если на листе есть хвостовые строки с большим r, это занижает last_row; при необходимости можно пересчитать отдельно.
        if let Some(dim_beg) = memmem::find(&self.sheet_xml, b"<dimension ref=\"") {
            let start = dim_beg + "<dimension ref=\"".len();
            if let Some(end) = find_bytes_from(&self.sheet_xml, b"\"", start) {
                let last_col_letter = col_letter(base_col + (df.width().saturating_sub(1) as u32));
                let dim = format!(
                    "{}{}:{}{}",
                    col_letter(base_col),
                    first_row,
                    last_col_letter,
                    last_row
                );
                self.sheet_xml.splice(start..end, dim.into_bytes());
            }
        }

        Ok(())
    }
}
