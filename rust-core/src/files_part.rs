/// files_part.rs
use crate::{XlsxEditor, find_bytes_from, scan};
use ::zip as zip_crate;
use anyhow::{Context, Result, bail};
use memchr::memmem;
use quick_xml::{Reader, events::Event};
use std::{
    collections::{HashMap, HashSet},
    fs::File,
    io::{Read, Write},
    path::Path,
}; // ← понадобится для dimension

pub(crate) fn needs_xml_space_preserve(s: &str) -> bool {
    let b = s.as_bytes();
    if b.is_empty() {
        return false;
    }
    b[0].is_ascii_whitespace()
        || b.last().unwrap().is_ascii_whitespace()
        || b.iter().any(|&c| matches!(c, b'\n' | b'\r' | b'\t'))
}

// A1 … A1:Z99
fn update_dimension_bytes(xml: &mut Vec<u8>) {
    // 1) вычисляем максимум по <c r="..">
    let (mut max_c, mut max_r) = (0u32, 0u32);
    let f = memmem::Finder::new(b"<c r=\"");
    let mut i = 0usize;
    while let Some(off) = f.find(&xml[i..]) {
        let start = i + off;
        let v0 = start + 6;
        let Some(v1) = super::find_bytes_from(xml, b"\"", v0) else {
            break;
        };
        let coord = &xml[v0..v1];
        if let Some(p) = coord.iter().position(|b| b.is_ascii_digit()) {
            let mut cidx: u32 = 0;
            for &b in &coord[..p] {
                let u = (b as char).to_ascii_uppercase() as u8;
                cidx = cidx * 26 + ((u - b'A') as u32 + 1);
            }
            if cidx > 0 {
                max_c = max_c.max(cidx - 1);
            }
            if let Ok(rs) = std::str::from_utf8(&coord[p..]) {
                if let Ok(rn) = rs.parse::<u32>() {
                    max_r = max_r.max(rn);
                }
            }
        }
        i = v1 + 1;
    }

    // 2) собрать новый <dimension .../>
    let dim_tag = if max_c == 0 && max_r <= 1 {
        r#"<dimension ref="A1"/>"#.to_string()
    } else {
        let last = crate::style::col_letter(max_c);
        format!(r#"<dimension ref="A1:{}{}"/>"#, last, max_r.max(1))
    };

    // 3) заменить/вставить с правильным местом
    if let Some(start) = memmem::find(xml, b"<dimension") {
        let end = super::find_bytes_from(xml, b">", start)
            .map(|p| p + 1)
            .unwrap_or(start);
        xml.splice(start..end, dim_tag.bytes());
    } else {
        // если есть <cols>, вставляем ПЕРЕД ним; иначе — перед <sheetData>
        if let Some(cols_pos) = memmem::find(xml, b"<cols") {
            xml.splice(cols_pos..cols_pos, dim_tag.bytes());
        } else if let Some(sd) = memmem::find(xml, b"<sheetData") {
            xml.splice(sd..sd, dim_tag.bytes());
        }
    }
}
// добавляет Override в [Content_Types].xml, если нет
fn ensure_ct_override_for_sheets(ct_xml: &mut Vec<u8>, sheet_paths: &[String]) {
    // ищем </Types>
    let Some(types_end) = memmem::rfind(ct_xml, b"</Types>") else {
        return;
    };

    for p in sheet_paths {
        if !p.starts_with("xl/worksheets/") || !p.ends_with(".xml") {
            continue;
        }
        let part = format!("/{}", p);
        let needle = format!(r#"PartName="{part}""#);
        if memmem::find(&ct_xml, needle.as_bytes()).is_some() {
            continue; // уже есть
        }
        let override_tag = format!(
            r#"<Override PartName="{part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        );
        ct_xml.splice(types_end..types_end, override_tag.bytes());
    }
}

/// Work with files
impl XlsxEditor {
    /// Открывает книгу и подготавливает лист `sheet_id` (1‑based).
    pub fn open_sheet<P: AsRef<Path>>(src: P, sheet_id: usize) -> Result<Self> {
        let src_path = src.as_ref().to_path_buf();
        let mut zip = zip_crate::ZipArchive::new(File::open(&src_path)?)?;

        // ── sheet#.xml ───────────────────────────────────────────────
        let sheet_path = format!("xl/worksheets/sheet{sheet_id}.xml");

        // читаем XML листа в отдельном блоке, чтобы `sheet` дропнулся,
        // и эксклюзивный займ `zip` освободился
        let sheet_xml: Vec<u8> = {
            let mut sheet = zip
                .by_name(&sheet_path)
                .with_context(|| format!("{sheet_path} not found"))?;
            let mut buf = Vec::with_capacity(sheet.size() as usize);
            sheet.read_to_end(&mut buf)?;
            buf
        };

        // ── styles.xml ───────────────────────────────────────────────
        let styles_xml: Vec<u8> = {
            let mut styles = zip
                .by_name("xl/styles.xml")
                .context("styles.xml not found")?;
            let mut buf = Vec::with_capacity(styles.size() as usize);
            styles.read_to_end(&mut buf)?;
            buf
        };

        // ── workbook.xml ───────────────────────────────────────────────
        let workbook_xml: Vec<u8> = {
            let mut wb = zip
                .by_name("xl/workbook.xml")
                .context("xl/workbook.xml not found")?;
            let mut buf = Vec::with_capacity(wb.size() as usize);
            wb.read_to_end(&mut buf)?;
            buf
        };

        // ── workbook.xml.rels ──────────────────────────────────────────
        let rels_xml: Vec<u8> = {
            let mut rels = zip
                .by_name("xl/_rels/workbook.xml.rels")
                .context("xl/_rels/workbook.xml.rels not found")?;
            let mut buf = Vec::with_capacity(rels.size() as usize);
            rels.read_to_end(&mut buf)?;
            buf
        };

        // ── вычисляем last_row ───────────────────────────────────────
        let mut reader = Reader::from_reader(sheet_xml.as_slice());
        // check_utf8(&mut reader)?;
        reader.config_mut().trim_text(true);

        let mut last_row = 0;
        while let Ok(ev) = reader.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"row" => {
                    if let Some(r) = e.attributes().with_checks(false).flatten().find_map(|a| {
                        (a.key.as_ref() == b"r")
                            .then(|| String::from_utf8_lossy(&a.value).into_owned())
                    }) {
                        last_row = r.parse::<u32>().unwrap_or(last_row);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        Ok(Self {
            src_path,
            sheet_path,
            sheet_xml,
            last_row,
            styles_xml,
            workbook_xml,
            rels_xml,
            new_files: HashMap::new(),
            styles_index: None,
            loaded_files: std::collections::HashMap::new(), // ← добавлено
            removed_files: HashSet::new(),                  // ← НОВОЕ
        })
    }

    fn flush_current_sheet(&mut self) {
        self.new_files
            .insert(self.sheet_path.clone(), self.sheet_xml.clone());

        // let cur_path = self.sheet_path.clone();
        // let cur_xml = self.sheet_xml.clone();
        // if let Some((_, c)) = self.new_files.iter_mut().find(|(p, _)| p == &cur_path) {
        //     *c = cur_xml;
        // } else {
        //     self.new_files.push((cur_path, cur_xml));
        // }
    }

    pub fn save<P: AsRef<Path>>(&mut self, dst: P) -> Result<()> {
        self.flush_current_sheet();

        let mut zin = zip_crate::ZipArchive::new(File::open(&self.src_path)?)?;
        let mut zout = zip_crate::ZipWriter::new(File::create(dst)?);

        let deflated: zip_crate::write::FileOptions<'_, ()> =
            zip_crate::write::FileOptions::default()
                .compression_method(zip_crate::CompressionMethod::Deflated)
                .compression_level(Some(1));

        let stored: zip_crate::write::FileOptions<'_, ()> =
            zip_crate::write::FileOptions::default()
                .compression_method(zip_crate::CompressionMethod::Stored);

        use std::collections::HashSet;
        let mut written: HashSet<String> = HashSet::new();

        // 0) прочитаем [Content_Types].xml в буфер и допишем Override для новых листов
        let mut ct_xml_opt: Option<Vec<u8>> = {
            let mut z = zip_crate::ZipArchive::new(File::open(&self.src_path)?)?;
            if let Ok(mut f) = z.by_name("[Content_Types].xml") {
                let mut buf = Vec::with_capacity(f.size() as usize);
                use std::io::Read;
                f.read_to_end(&mut buf).ok();
                Some(buf)
            } else {
                None
            }
        };
        if let Some(ct) = ct_xml_opt.as_mut() {
            // собрать список новых листов
            let mut new_sheet_paths: Vec<String> = self
                .new_files
                .keys()
                .filter(|p| p.starts_with("xl/worksheets/") && p.ends_with(".xml"))
                .cloned()
                .collect();
            // (опционально) если текущий лист вообще новый — он уже в new_files
            new_sheet_paths.sort();
            new_sheet_paths.dedup();
            ensure_ct_override_for_sheets(ct, &new_sheet_paths);
            // НОВОЕ: убрать overrides для удалённых листов
            for p in &self.removed_files {
                if p.starts_with("xl/worksheets/") && p.ends_with(".xml") {
                    remove_ct_override_for_path(ct, p);
                }
            }
        }

        for i in 0..zin.len() {
            let name = { zin.by_index_raw(i)?.name().to_string() };
            // НОВОЕ: пропускаем удалённые файлы
            if self.removed_files.contains(&name) {
                continue;
            }

            // Новая версия файла?
            if let Some(content) = self.new_files.get(&name) {
                let mut out = content.clone();

                // фиксим dimension для листов
                if name.starts_with("xl/worksheets/") && name.ends_with(".xml") {
                    update_dimension_bytes(&mut out);
                }

                let opt = if should_store_uncompressed(&name, out.len()) {
                    stored
                } else {
                    deflated
                };
                zout.start_file(&name, opt)?;
                zout.write_all(&out)?;
                written.insert(name);
                continue;
            }

            match name.as_str() {
                "xl/workbook.xml" => {
                    let content = &self.workbook_xml;
                    let opt = if should_store_uncompressed(&name, content.len()) {
                        stored
                    } else {
                        deflated
                    };
                    zout.start_file(&name, opt)?;
                    zout.write_all(content)?;
                }
                "xl/_rels/workbook.xml.rels" => {
                    let content = &self.rels_xml;
                    let opt = if should_store_uncompressed(&name, content.len()) {
                        stored
                    } else {
                        deflated
                    };
                    zout.start_file(&name, opt)?;
                    zout.write_all(content)?;
                }
                "xl/styles.xml" => {
                    // пишем актуальные стили, а не raw_copy
                    let mut content = self.styles_xml.clone();
                    normalize_styles_root(&mut content); // как и раньше
                    let opt = if should_store_uncompressed(&name, content.len()) {
                        stored
                    } else {
                        deflated
                    };
                    zout.start_file(&name, opt)?;
                    zout.write_all(&content)?;
                }

                "[Content_Types].xml" => {
                    // если удалось прочитать/поправить — пишем её, иначе raw_copy
                    if let Some(ct) = ct_xml_opt.take() {
                        let opt = if should_store_uncompressed(&name, ct.len()) {
                            stored
                        } else {
                            deflated
                        };
                        zout.start_file(&name, opt)?;
                        zout.write_all(&ct)?;
                    } else {
                        let f = zin.by_index_raw(i)?;
                        zout.raw_copy_file(f)?;
                    }
                }
                p if p == self.sheet_path => {
                    // текущий лист (в буфере self.sheet_xml) — обновим dimension перед записью
                    let mut content = self.sheet_xml.clone();
                    update_dimension_bytes(&mut content);

                    let opt = if should_store_uncompressed(&name, content.len()) {
                        stored
                    } else {
                        deflated
                    };
                    zout.start_file(&name, opt)?;
                    zout.write_all(&content)?;
                }
                _ => {
                    let f = zin.by_index_raw(i)?;
                    zout.raw_copy_file(f)?
                }
            }
        }

        for (path, content) in &self.new_files {
            if self.removed_files.contains(path) {
                continue; // НОВОЕ: не писать удалённые
            }
            if !written.contains(path) {
                let mut out = content.clone();
                if path.starts_with("xl/worksheets/") && path.ends_with(".xml") {
                    update_dimension_bytes(&mut out);
                }
                let opt = if should_store_uncompressed(path, out.len()) {
                    stored
                } else {
                    deflated
                };
                zout.start_file(path, opt)?;
                zout.write_all(&out)?;
            }
        }

        zout.finish()?;
        Ok(())
    }
}

impl XlsxEditor {
    /// Считает количество листов по текущему состоянию `workbook_xml`
    fn sheet_count(&self) -> usize {
        let mut rdr = Reader::from_reader(self.workbook_xml.as_slice());
        rdr.config_mut().trim_text(true);
        let mut n = 0usize;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"sheet" => {
                    n += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        n
    }

    /// Возвращает (позиция_начала_контента, позиция_конца_контента) для содержимого между
    /// `<sheets ...>` и `</sheets>` в `workbook_xml`.
    fn find_sheets_section(workbook_xml: &[u8]) -> Result<(usize, usize)> {
        let xml = workbook_xml;
        let open_tag =
            memchr::memmem::find(xml, b"<sheets").context("<sheets> not found in workbook.xml")?;
        let mut pos = open_tag;
        // ищем '>' у открывающего <sheets ...>
        while pos < xml.len() && xml[pos] != b'>' {
            pos += 1;
        }
        if pos >= xml.len() {
            bail!("Malformed workbook.xml: <sheets ...> not closed with '>'");
        }
        let content_start = pos + 1;

        let close_tag = memchr::memmem::rfind(xml, b"</sheets>")
            .context("</sheets> not found in workbook.xml")?;
        Ok((content_start, close_tag))
    }

    /// Добавляет новый пустой лист c именем `sheet_name` **на позицию `index` (0‑based)**,
    /// пересобирая порядок `<sheet/>` в workbook.xml.
    pub fn add_worksheet_at(&mut self, sheet_name: &str, mut index: usize) -> Result<&mut Self> {
        // -------- 0) валидации / подготовка ----------
        // 0.1) имя уже существует?
        let sheet_names = scan(&self.src_path)?;
        if sheet_names.contains(&sheet_name.to_owned()) {
            bail!("Sheet {} already exists", sheet_name);
        }

        // 0.2) текущее количество листов
        let cur_cnt = self.sheet_count();
        if index > cur_cnt {
            index = cur_cnt; // кладём в конец
        }

        // 0.3) читаем исходный архив (для поиска свободного sheet#.xml)
        let mut zin = zip::ZipArchive::new(File::open(&self.src_path)?)?;

        // 0.4) локальные (редактируемые) копии XML
        let mut wb_xml = self.workbook_xml.clone();
        let mut rels_xml = self.rels_xml.clone();

        // -------- 1) найдём max sheetId и max rId ----------
        let mut max_sheet_id = 0u32;
        let mut rdr = Reader::from_reader(wb_xml.as_slice());
        rdr.config_mut().trim_text(true);
        while let Ok(ev) = rdr.read_event() {
            if let Event::Empty(ref e) | Event::Start(ref e) = ev {
                if e.name().as_ref() == b"sheet" {
                    if let Some(id) = e.attributes().with_checks(false).flatten().find_map(|a| {
                        (a.key.as_ref() == b"sheetId")
                            .then(|| String::from_utf8_lossy(&a.value).into_owned())
                    }) {
                        max_sheet_id = max_sheet_id.max(id.parse::<u32>().unwrap_or(0));
                    }
                }
            }
            if matches!(ev, Event::Eof) {
                break;
            }
        }
        // Новый sheetId нам не особо важен (мы потом все перенумеруем), но пусть будет > max_sheet_id
        let _new_sheet_id = max_sheet_id + 1;

        let mut max_rid = 0u32;
        let mut rdr = Reader::from_reader(rels_xml.as_slice());
        rdr.config_mut().trim_text(true);
        while let Ok(ev) = rdr.read_event() {
            if let Event::Empty(ref e) | Event::Start(ref e) = ev {
                if e.name().as_ref() == b"Relationship" {
                    if let Some(id) = e.attributes().with_checks(false).flatten().find_map(|a| {
                        (a.key.as_ref() == b"Id")
                            .then(|| String::from_utf8_lossy(&a.value).into_owned())
                    }) {
                        if let Some(num) = id.strip_prefix("rId") {
                            max_rid = max_rid.max(num.parse::<u32>().unwrap_or(0));
                        }
                    }
                }
            }
            if matches!(ev, Event::Eof) {
                break;
            }
        }
        let new_rid = max_rid + 1;

        // -------- 2) найти свободный sheet#.xml ----------
        let mut max_sheet_file = 0usize;
        for i in 0..zin.len() {
            let name = zin.by_index(i)?.name().to_owned();
            if let Some(n) = name
                .strip_prefix("xl/worksheets/sheet")
                .and_then(|s| s.strip_suffix(".xml"))
                .and_then(|s| s.parse::<usize>().ok())
            {
                max_sheet_file = max_sheet_file.max(n);
            }
        }
        for path in self.new_files.keys() {
            if let Some(n) = path
                .strip_prefix("xl/worksheets/sheet")
                .and_then(|s| s.strip_suffix(".xml"))
                .and_then(|s| s.parse::<usize>().ok())
            {
                max_sheet_file = max_sheet_file.max(n);
            }
        }
        let new_sheet_file = max_sheet_file + 1;
        let new_sheet_path = format!("xl/worksheets/sheet{new}.xml", new = new_sheet_file);
        let new_sheet_target = format!("worksheets/sheet{new}.xml", new = new_sheet_file);

        // -------- 3) распарсим текущие <sheet .../> из workbook.xml ----------
        #[allow(dead_code)]
        #[derive(Debug, Clone)]
        struct SheetTag {
            name: String,
            rid: String,  // "rIdNN"
            path: String, // worksheets/sheet#.xml (нам нужно только для инфы; можно не хранить)
        }
        let (sheets_content_start, sheets_content_end) = Self::find_sheets_section(&wb_xml)?;
        let sheets_slice = &wb_xml[sheets_content_start..sheets_content_end];

        let mut rdr = Reader::from_reader(sheets_slice);
        rdr.config_mut().trim_text(true);
        let mut sheets: Vec<SheetTag> = Vec::new();

        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"sheet" => {
                    let mut name = None;
                    let mut rid = None;
                    // Target пути тут нет — он в rels, так что просто пустим.
                    for a in e.attributes().with_checks(false).flatten() {
                        let k = a.key.as_ref();
                        let v = String::from_utf8_lossy(&a.value).into_owned();
                        if k == b"name" {
                            name = Some(v.clone());
                        }
                        if k == b"r:id" {
                            rid = Some(v);
                        }
                    }
                    sheets.push(SheetTag {
                        name: name.unwrap_or_default(),
                        rid: rid.unwrap_or_default(),
                        path: String::new(),
                    });
                }
                Event::Eof => break,
                _ => {}
            }
        }

        // -------- 4) формируем новый tag для нового листа ----------
        let new_sheet = SheetTag {
            name: sheet_name.to_string(),
            rid: format!("rId{}", new_rid),
            path: new_sheet_target.clone(),
        };

        // вставляем по индексу
        if index >= sheets.len() {
            sheets.push(new_sheet);
        } else {
            sheets.insert(index, new_sheet);
        }

        // -------- 5) перегенерируем <sheets>...</sheets> с новой нумерацией sheetId ----------
        let mut new_inner = Vec::new();
        // Сохраним форматирование: перенос строки + два пробела
        for (i, sh) in sheets.iter().enumerate() {
            let sheet_id = (i as u32) + 1; // «естественная» нумерация
            let line = format!(
                "\n  <sheet name=\"{}\" sheetId=\"{}\" r:id=\"{}\"/>",
                xml_escape(&sh.name),
                sheet_id,
                sh.rid
            );
            new_inner.extend_from_slice(line.as_bytes());
        }

        // подменяем содержимое между <sheets ...> и </sheets>
        wb_xml.splice(
            sheets_content_start..sheets_content_end,
            new_inner.into_iter(),
        );

        // -------- 6) вставляем Relationship под конец </Relationships> ----------
        let rel_tag = format!(
            r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="{}"/>"#,
            new_rid, new_sheet_target
        );
        if let Some(pos) = memmem::rfind(&rels_xml, b"</Relationships") {
            rels_xml.splice(pos..pos, rel_tag.bytes());
        } else {
            bail!("</Relationships> not found in workbook.xml.rels");
        }

        // -------- 7) минимальный XML нового листа ----------
        const EMPTY_SHEET: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData> </sheetData>
        </worksheet>"#;

        // Обновляем внутреннее состояние
        self.workbook_xml = wb_xml;
        self.rels_xml = rels_xml;

        // забираем буферы без clone
        {
            let cur_path = std::mem::take(&mut self.sheet_path);
            let cur_xml = std::mem::take(&mut self.sheet_xml);
            if !cur_path.is_empty() {
                self.new_files.insert(cur_path, cur_xml);
            }
        }

        // сразу кладём заготовку нового листа
        self.new_files
            .insert(new_sheet_path.clone(), EMPTY_SHEET.as_bytes().to_vec());

        // переключаемся
        self.sheet_path = new_sheet_path;
        self.sheet_xml = EMPTY_SHEET.as_bytes().to_vec();
        self.last_row = 0;

        Ok(self)
    }

    /// Старый API: просто добавляет в конец.
    pub fn add_worksheet(&mut self, sheet_name: &str) -> Result<&mut Self> {
        let last_idx = self.sheet_count(); // вставка в конец
        self.add_worksheet_at(sheet_name, last_idx)
    }
}

impl XlsxEditor {
    pub fn with_worksheet(&mut self, sheet_name: &str) -> Result<&mut Self> {
        // 0) Сохраняем текущий лист в new_files без лишних clone
        {
            let cur_path = std::mem::take(&mut self.sheet_path);
            let cur_xml = std::mem::take(&mut self.sheet_xml);
            if !cur_path.is_empty() {
                self.new_files.insert(cur_path, cur_xml);
            }
        }

        // 1) Найти r:id по имени листа в workbook.xml
        let mut rdr = Reader::from_reader(self.workbook_xml.as_slice());
        rdr.config_mut().trim_text(true);

        let mut target_rid: Option<String> = None;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"sheet" => {
                    let mut name: Option<String> = None;
                    let mut rid: Option<String> = None;

                    for a in e.attributes().with_checks(false).flatten() {
                        let k = a.key.as_ref();
                        let v = String::from_utf8_lossy(&a.value).into_owned();
                        if k == b"name" {
                            name = Some(v.clone());
                        }
                        if k == b"r:id" {
                            rid = Some(v);
                        }
                    }

                    if let (Some(n), Some(r)) = (name, rid) {
                        if n == sheet_name {
                            target_rid = Some(r);
                            break;
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        let target_rid = target_rid
            .with_context(|| format!("Sheet `{}` not found in workbook.xml", sheet_name))?;

        // 2) По r:id найти Target в workbook.xml.rels
        let mut rdr = Reader::from_reader(self.rels_xml.as_slice());
        rdr.config_mut().trim_text(true);

        let mut target_rel: Option<String> = None;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Empty(ref e) | Event::Start(ref e)
                    if e.name().as_ref() == b"Relationship" =>
                {
                    let mut id: Option<String> = None;
                    let mut target: Option<String> = None;

                    for a in e.attributes().with_checks(false).flatten() {
                        let k = a.key.as_ref();
                        let v = String::from_utf8_lossy(&a.value).into_owned();
                        if k == b"Id" {
                            id = Some(v.clone());
                        }
                        if k == b"Target" {
                            target = Some(v);
                        }
                    }

                    if let (Some(idv), Some(t)) = (id, target) {
                        if idv == target_rid {
                            target_rel = Some(t);
                            break;
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        let target_rel = target_rel.with_context(|| {
            format!(
                "Relationship for `{}` not found in workbook.xml.rels",
                sheet_name
            )
        })?;

        // 3) Абсолютный путь внутри архива
        let new_sheet_path = if target_rel.starts_with("xl/") {
            target_rel.clone()
        } else {
            format!("xl/{}", target_rel)
        };

        // 4) Берём XML листа: new_files → loaded_files → zip
        let sheet_xml: Vec<u8> = if let Some(content) = self.new_files.get(&new_sheet_path) {
            content.clone()
        } else if let Some(buf) = self.loaded_files.get(&new_sheet_path) {
            buf.clone()
        } else {
            let mut zin = zip_crate::ZipArchive::new(File::open(&self.src_path)?)?;
            let mut f = zin
                .by_name(&new_sheet_path)
                .with_context(|| format!("{} not found in zip", new_sheet_path))?;
            let mut buf = Vec::with_capacity(f.size() as usize);
            f.read_to_end(&mut buf)?;
            self.loaded_files
                .insert(new_sheet_path.clone(), buf.clone()); // кэшируем
            buf
        };

        // 5) Пересчитываем last_row и переключаемся
        let last_row = calc_last_row(&sheet_xml);

        self.sheet_path = new_sheet_path;
        self.sheet_xml = sheet_xml;
        self.last_row = last_row;

        Ok(self)
    }
}

impl XlsxEditor {
    pub fn rename_worksheet(&mut self, old_name: &str, new_name: &str) -> Result<&mut Self> {
        if old_name == new_name {
            return Ok(self);
        }
        // запретим коллизию имён
        let names = scan(&self.src_path)?;
        if names.iter().any(|n| n == new_name) {
            anyhow::bail!("Sheet `{}` already exists", new_name);
        }

        // парсим <sheets> и меняем имя
        let (s_start, s_end, mut tags) = parse_sheets_inner(&self.workbook_xml)?;
        let mut found = false;
        for t in &mut tags {
            if t.name == old_name {
                t.name = new_name.to_string();
                found = true;
                break;
            }
        }
        if !found {
            anyhow::bail!("Sheet `{}` not found", old_name);
        }

        // собираем обратно (sheetId не трогаем при rename)
        let new_inner = build_sheets_inner(&tags, /*renumber=*/ false);
        self.workbook_xml.splice(s_start..s_end, new_inner);

        Ok(self)
    }
}

impl XlsxEditor {
    pub fn delete_worksheet(&mut self, name: &str) -> Result<&mut Self> {
        // 0) r:id по имени
        let Some(rid) = get_sheet_rid_by_name(&self.workbook_xml, name) else {
            anyhow::bail!("Sheet `{}` not found in workbook.xml", name);
        };

        // 1) удалить из списка <sheets> (и тут же пере-нумеровать sheetId подряд)
        let (s_start, s_end, mut tags) = parse_sheets_inner(&self.workbook_xml)?;
        let orig_len = tags.len();
        tags.retain(|t| t.name != name);
        if tags.len() == orig_len {
            anyhow::bail!("Sheet `{}` tag not found in <sheets>", name);
        }
        let new_inner = build_sheets_inner(&tags, /*renumber=*/ true);
        self.workbook_xml.splice(s_start..s_end, new_inner);

        // 2) убрать Relationship и получить Target
        let Some((rel_start, rel_end, target)) = find_relationship_by_id(&self.rels_xml, &rid)
        else {
            anyhow::bail!("Relationship `{}` not found in workbook.xml.rels", rid);
        };
        self.rels_xml.splice(rel_start..rel_end, std::iter::empty());

        // 3) отметить файл на удаление + вычистить кэши
        let abs_path = if target.starts_with("xl/") {
            target.clone()
        } else {
            format!("xl/{}", target)
        };
        self.removed_files.insert(abs_path.clone());
        self.new_files.remove(&abs_path);
        self.loaded_files.remove(&abs_path);

        // 4) если удалили активный лист — переключиться на первый оставшийся (или создать новый)
        if self.sheet_path == abs_path {
            // найдём первый r:id оставшегося листа
            let mut rdr = quick_xml::Reader::from_reader(self.workbook_xml.as_slice());
            rdr.config_mut().trim_text(true);
            let mut next_rid: Option<String> = None;
            while let Ok(ev) = rdr.read_event() {
                if let quick_xml::events::Event::Empty(ref e)
                | quick_xml::events::Event::Start(ref e) = ev
                {
                    if e.name().as_ref() == b"sheet" {
                        for a in e.attributes().with_checks(false).flatten() {
                            if a.key.as_ref() == b"r:id" {
                                next_rid = Some(String::from_utf8_lossy(&a.value).into_owned());
                                break;
                            }
                        }
                        if next_rid.is_some() {
                            break;
                        }
                    }
                }
                if matches!(ev, quick_xml::events::Event::Eof) {
                    break;
                }
            }

            if let Some(nrid) = next_rid {
                if let Some((_, _, t)) = find_relationship_by_id(&self.rels_xml, &nrid) {
                    let p = if t.starts_with("xl/") {
                        t
                    } else {
                        format!("xl/{}", t)
                    };
                    let sheet_xml: Vec<u8> = if let Some(c) = self.new_files.get(&p) {
                        c.clone()
                    } else if let Some(buf) = self.loaded_files.get(&p) {
                        buf.clone()
                    } else {
                        let mut zin = zip::ZipArchive::new(std::fs::File::open(&self.src_path)?)?;
                        let mut f = zin
                            .by_name(&p)
                            .with_context(|| format!("{} not found in zip", p))?;
                        let mut buf = Vec::with_capacity(f.size() as usize);
                        use std::io::Read;
                        f.read_to_end(&mut buf)?;
                        self.loaded_files.insert(p.clone(), buf.clone());
                        buf
                    };
                    self.sheet_path = p;
                    self.sheet_xml = sheet_xml;
                    self.last_row = calc_last_row(&self.sheet_xml);
                } else {
                    self.add_worksheet("Sheet1")?;
                }
            } else {
                self.add_worksheet("Sheet1")?;
            }
        }

        // 5) чистка Content_Types перенесена в save(), как раньше
        Ok(self)
    }
}

// маленький хелпер
fn calc_last_row(sheet_xml: &[u8]) -> u32 {
    let mut rdr = Reader::from_reader(sheet_xml);
    rdr.config_mut().trim_text(true);

    let mut last_row = 0u32;
    while let Ok(ev) = rdr.read_event() {
        match ev {
            Event::Empty(ref e) | Event::Start(ref e) if e.name().as_ref() == b"row" => {
                if let Some(r) = e.attributes().with_checks(false).flatten().find_map(|a| {
                    (a.key.as_ref() == b"r").then(|| String::from_utf8_lossy(&a.value).into_owned())
                }) {
                    last_row = r.parse::<u32>().unwrap_or(last_row);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    last_row
}

// Простейший экранировщик для XML-атрибутов.
fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}
fn should_store_uncompressed(name: &str, content_len: usize) -> bool {
    // Можно подобрать порог — эмпирически 64–128 КБ дают профит
    name.ends_with(".xml") && content_len <= 128 * 1024
}
fn normalize_styles_root(xml: &mut Vec<u8>) {
    if let Some(end_root) = memmem::rfind(xml, b"</styleSheet>") {
        let tail_start = end_root + "</styleSheet>".len();
        if tail_start < xml.len() {
            if let Some(nf) = memmem::find(&xml[tail_start..], b"<numFmts") {
                // вырезаем хвост <numFmts>...</numFmts>
                let nf_abs = tail_start + nf;
                let close = memmem::find(&xml[nf_abs..], b"</numFmts>")
                    .map(|p| nf_abs + p + "</numFmts>".len());
                if let Some(nf_end) = close {
                    let chunk: Vec<u8> = xml[nf_abs..nf_end].to_vec();
                    xml.splice(nf_abs..nf_end, std::iter::empty());
                    // вставляем внутрь корня (см. логику из пункта 1/2)
                    let root = memmem::find(xml, b"<styleSheet").unwrap();
                    let insert = find_bytes_from(xml, b">", root).unwrap() + 1;
                    xml.splice(insert..insert, chunk.into_iter());
                }
            }
        }
    }
}


// вытащить r:id листа по его имени
fn get_sheet_rid_by_name(wb_xml: &[u8], target_name: &str) -> Option<String> {
    let (s_start, s_end) = XlsxEditor::find_sheets_section(wb_xml).ok()?;
    let slice = &wb_xml[s_start..s_end];
    let mut rdr = Reader::from_reader(slice);
    rdr.config_mut().trim_text(true);
    while let Ok(ev) = rdr.read_event() {
        if let Event::Empty(ref e) | Event::Start(ref e) = ev {
            if e.name().as_ref() == b"sheet" {
                let mut name = None;
                let mut rid = None;
                for a in e.attributes().with_checks(false).flatten() {
                    let k = a.key.as_ref();
                    let v = String::from_utf8_lossy(&a.value).into_owned();
                    if k == b"name" {
                        name = Some(v.clone());
                    }
                    if k == b"r:id" {
                        rid = Some(v);
                    }
                }
                if let (Some(n), Some(r)) = (name, rid) {
                    if n == target_name {
                        return Some(r);
                    }
                }
            }
        }
        if matches!(ev, Event::Eof) {
            break;
        }
    }
    None
}

// найти Relationship по rId и вернуть (абсол.начало, абсол.конец, target)
fn find_relationship_by_id(rels_xml: &[u8], rid: &str) -> Option<(usize, usize, String)> {
    let mut rdr = Reader::from_reader(rels_xml);
    rdr.config_mut().trim_text(true);

    // грубо ищем подпоследовательность Id="rid"
    let needle = format!(r#"Id="{}""#, rid);
    let pos = memmem::find(rels_xml, needle.as_bytes())?;
    // откатимся назад до начала тега <Relationship
    let rel_start = memmem::rfind(&rels_xml[..pos], b"<Relationship")?;
    let rel_end = find_bytes_from(rels_xml, b">", rel_start).map(|p| p + 1)?;
    // достанем Target
    let tag = &rels_xml[rel_start..rel_end];
    let tkey = b" Target=\"";
    if let Some(t0) = memmem::find(tag, tkey) {
        let v0 = rel_start + t0 + tkey.len();
        let v1 = find_bytes_from(rels_xml, b"\"", v0)?;
        let target = String::from_utf8_lossy(&rels_xml[v0..v1]).into_owned();
        return Some((rel_start, rel_end, target));
    }
    None
}

fn remove_ct_override_for_path(ct_xml: &mut Vec<u8>, part_abs_path: &str) {
    // part_abs_path должен начинаться с "xl/..."
    let part = format!("/{}", part_abs_path);
    let needle = format!(r#"PartName="{}""#, part);
    while let Some(attr_pos) = memmem::find(&ct_xml, needle.as_bytes()) {
        // откатимся к началу тега "<Override"
        let tag_start = memmem::rfind(&ct_xml[..attr_pos], b"<Override").unwrap_or(attr_pos);
        // и найдём конец '>'
        let tag_end = find_bytes_from(&ct_xml, b">", attr_pos)
            .map(|p| p + 1)
            .unwrap_or(attr_pos);
        ct_xml.splice(tag_start..tag_end, std::iter::empty());
    }
}
#[derive(Debug, Clone)]
struct SheetTagMini {
    name: String,
    rid: String,
    sheet_id: String, // сохраняем как строку: Excel не обязан идти по порядку; при delete можно перенумеровать
}

// распарсить содержимое между <sheets>...</sheets> в вектор SheetTagMini
fn parse_sheets_inner(wb_xml: &[u8]) -> anyhow::Result<(usize, usize, Vec<SheetTagMini>)> {
    let (s_start, s_end) = XlsxEditor::find_sheets_section(wb_xml)?;
    let slice = &wb_xml[s_start..s_end];

    let mut rdr = quick_xml::Reader::from_reader(slice);
    rdr.config_mut().trim_text(true);

    let mut out = Vec::<SheetTagMini>::new();
    while let Ok(ev) = rdr.read_event() {
        match ev {
            quick_xml::events::Event::Empty(ref e) | quick_xml::events::Event::Start(ref e)
                if e.name().as_ref() == b"sheet" =>
            {
                let mut name = None;
                let mut rid = None;
                let mut sid = None;
                for a in e.attributes().with_checks(false).flatten() {
                    let k = a.key.as_ref();
                    let v = String::from_utf8_lossy(&a.value).into_owned();
                    if k == b"name" {
                        name = Some(v.clone());
                    } else if k == b"r:id" {
                        rid = Some(v.clone());
                    } else if k == b"sheetId" {
                        sid = Some(v.clone());
                    }
                }
                out.push(SheetTagMini {
                    name: name.unwrap_or_default(),
                    rid: rid.unwrap_or_default(),
                    sheet_id: sid.unwrap_or_else(|| "0".to_string()),
                });
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
    }
    Ok((s_start, s_end, out))
}

// собрать новый inner <sheets> из массива SheetTagMini (с контролем sheetId)
fn build_sheets_inner(tags: &[SheetTagMini], renumber: bool) -> Vec<u8> {
    let mut buf = Vec::new();
    for (i, t) in tags.iter().enumerate() {
        let sid = if renumber {
            (i as u32 + 1).to_string()
        } else {
            t.sheet_id.clone()
        };
        let line = format!(
            "\n  <sheet name=\"{}\" sheetId=\"{}\" r:id=\"{}\"/>",
            xml_escape(&t.name),
            sid,
            t.rid
        );
        buf.extend_from_slice(line.as_bytes());
    }
    buf
}
