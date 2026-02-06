//! style.rs – универсальный слой стилей + нормализация <cols>

use anyhow::{Context, Result, bail};
use memchr::{memchr, memmem, memrchr};
use quick_xml::{Reader, events::Event};
use std::collections::HashMap;
use std::{fmt, str::FromStr};

use crate::style::util::{bump_count, col_index, find_bytes_from};
use crate::{FontKey, StyleIndex, StyleKey, XfParts, XlsxEditor};

mod cols;
pub mod util;

pub use util::{col_letter, split_coord};

/* ========================== ALIGNMENT API ================================= */

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum HorizAlignment {
    Left,
    Center,
    Right,
    Fill,
    Justify,
}
impl fmt::Display for HorizAlignment {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            HorizAlignment::Left => "left",
            HorizAlignment::Center => "center",
            HorizAlignment::Right => "right",
            HorizAlignment::Fill => "fill",
            HorizAlignment::Justify => "justify",
        })
    }
}
impl FromStr for HorizAlignment {
    type Err = anyhow::Error;
    fn from_str(s: &str) -> Result<Self> {
        Ok(match s {
            "left" => HorizAlignment::Left,
            "center" => HorizAlignment::Center,
            "right" => HorizAlignment::Right,
            "fill" => HorizAlignment::Fill,
            "justify" => HorizAlignment::Justify,
            _ => bail!("Unknown horizontal alignment: {s}"),
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum VertAlignment {
    Top,
    Center,
    Bottom,
    Justify,
}
impl fmt::Display for VertAlignment {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            VertAlignment::Top => "top",
            VertAlignment::Center => "center",
            VertAlignment::Bottom => "bottom",
            VertAlignment::Justify => "justify",
        })
    }
}
impl FromStr for VertAlignment {
    type Err = anyhow::Error;
    fn from_str(s: &str) -> Result<Self> {
        Ok(match s {
            "top" => VertAlignment::Top,
            "center" => VertAlignment::Center,
            "bottom" => VertAlignment::Bottom,
            "justify" => VertAlignment::Justify,
            _ => bail!("Unknown vertical alignment: {s}"),
        })
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct AlignSpec {
    pub horiz: Option<HorizAlignment>,
    pub vert: Option<VertAlignment>,
    pub wrap: bool,
}

/* ========================== CORE STYLE STRUCT ============================= */

#[derive(Debug, Clone, Default)]
struct StyleParts {
    pub num_fmt_code: Option<String>,
    pub font: Option<u32>,
    pub fill: Option<u32>,
    pub border: Option<u32>,
    pub align: Option<AlignSpec>,
}

/* ========================== TARGET PARSER ================================= */

#[derive(Debug)]
enum Target {
    Cell(String),
    Rect { c0: u32, r0: u32, c1: u32, r1: u32 },
    Col(u32), // 0-based
    Row(u32),
}

fn parse_target(s: &str) -> Result<Target> {
    // столбец "A:" ?
    if s.ends_with(':') && s[..s.len() - 1].bytes().all(|b| b.is_ascii_alphabetic()) {
        return Ok(Target::Col(col_index(&s[..s.len() - 1]) as u32));
    }
    // строка "12:" ?
    if s.ends_with(':') && s[..s.len() - 1].bytes().all(|b| b.is_ascii_digit()) {
        return Ok(Target::Row(s[..s.len() - 1].parse()?));
    }
    // диапазон?
    if let Some(colon) = s.find(':') {
        let (a, b) = (&s[..colon], &s[colon + 1..]);
        if a.ends_with(':') || b.is_empty() {
            bail!("invalid range: {s}");
        }
        let (c0, r0) = split_coord(a);
        let (c1, r1) = split_coord(b);
        return Ok(Target::Rect { c0, r0, c1, r1 });
    }

    // ячейка "A12"
    let p = s
        .find(|c: char| c.is_ascii_digit())
        .ok_or_else(|| anyhow::anyhow!("invalid"))?;
    if s[..p].bytes().all(|b| b.is_ascii_alphabetic()) && s[p..].bytes().all(|b| b.is_ascii_digit())
    {
        return Ok(Target::Cell(s.to_owned()));
    }
    bail!("invalid range syntax: {s}");
}

fn parse_open_column_selector(s: &str) -> Option<(u32, u32)> {
    if !s.ends_with(':') {
        return None;
    }
    let head = &s[..s.len() - 1];
    if head.is_empty() {
        return None;
    }
    if head.bytes().all(|b| b.is_ascii_alphabetic()) {
        return Some((col_index(head) as u32, 1));
    }

    let p = head.find(|c: char| c.is_ascii_digit())?;
    if p == 0 {
        return None;
    }
    if !head[..p].bytes().all(|b| b.is_ascii_alphabetic()) {
        return None;
    }
    if !head[p..].bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }

    let row_start = head[p..].parse::<u32>().ok()?;
    if row_start == 0 {
        return None;
    }
    Some((col_index(&head[..p]) as u32, row_start))
}

impl StyleIndex {
    fn build(styles: &[u8]) -> Result<Self> {
        let mut ix = StyleIndex {
            xfs: Vec::new(),
            numfmt_by_code: HashMap::new(),
            next_custom_numfmt: 164,

            font_by_key: HashMap::new(),
            fill_by_rgb: HashMap::new(),
            border_by_key: HashMap::new(),
            xf_by_key: HashMap::new(),

            fonts_count: 0,
            fills_count: 0,
            borders_count: 0,
        };

        let mut rdr = Reader::from_reader(styles);
        rdr.config_mut().trim_text(true);

        // --- numFmts ---
        // Сканируем блок <numFmts> и заполняем карту code->id, заодно поднимаем next_custom_numfmt
        let mut max_custom = 163u32;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) | Event::Empty(ref e) if e.name().as_ref() == b"numFmt" => {
                    let mut id = None::<u32>;
                    let mut code = None::<String>;
                    for a in e.attributes().with_checks(false).flatten() {
                        match a.key.as_ref() {
                            b"numFmtId" => id = Some(lexical_core::parse(&a.value)?),
                            b"formatCode" => code = Some(String::from_utf8_lossy(&a.value).into()),
                            _ => {}
                        }
                    }
                    if let (Some(i), Some(c)) = (id, code) {
                        ix.numfmt_by_code.insert(c, i);
                        if i > max_custom {
                            max_custom = i;
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
        ix.next_custom_numfmt = max_custom.max(163) + 1;

        // --- fonts ---
        // Пройдём ещё раз для удобства (можно и за один проход, но так проще)
        let mut rdr = Reader::from_reader(styles);
        rdr.config_mut().trim_text(true);
        let mut in_fonts = false;
        let mut font_id = 0u32;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"fonts" => in_fonts = true,
                Event::End(ref e) if e.name().as_ref() == b"fonts" => {
                    // in_fonts = false;
                    break;
                }
                Event::Start(ref e) if in_fonts && e.name().as_ref() == b"font" => {
                    // читаем <font> целиком
                    let mut depth = 1;
                    let mut bold = false;
                    let mut italic = false;
                    let mut size: f32 = 11.0;
                    let mut name: String = "Calibri".into();

                    while depth > 0 {
                        match rdr.read_event()? {
                            Event::Start(ref fe) => {
                                depth += 1;
                                match fe.name().as_ref() {
                                    b"b" => bold = true,
                                    b"i" => italic = true,
                                    b"sz" => {
                                        for a in fe.attributes().with_checks(false).flatten() {
                                            if a.key.as_ref() == b"val" {
                                                size = String::from_utf8_lossy(&a.value)
                                                    .parse()
                                                    .unwrap_or(11.0);
                                            }
                                        }
                                    }
                                    b"name" => {
                                        for a in fe.attributes().with_checks(false).flatten() {
                                            if a.key.as_ref() == b"val" {
                                                name =
                                                    String::from_utf8_lossy(&a.value).into_owned();
                                            }
                                        }
                                    }
                                    _ => {}
                                }
                            }
                            Event::Empty(ref fe) => match fe.name().as_ref() {
                                b"b" => bold = true,
                                b"i" => italic = true,
                                b"sz" => {
                                    for a in fe.attributes().with_checks(false).flatten() {
                                        if a.key.as_ref() == b"val" {
                                            size = String::from_utf8_lossy(&a.value)
                                                .parse()
                                                .unwrap_or(11.0);
                                        }
                                    }
                                }
                                b"name" => {
                                    for a in fe.attributes().with_checks(false).flatten() {
                                        if a.key.as_ref() == b"val" {
                                            name = String::from_utf8_lossy(&a.value).into_owned();
                                        }
                                    }
                                }
                                _ => {}
                            },
                            Event::End(_) => depth -= 1,
                            Event::Eof => break,
                            _ => {}
                        }
                    }

                    let key = FontKey {
                        name,
                        size_100: (size * 100.0).round() as u32,
                        bold,
                        italic,
                    };
                    ix.font_by_key.entry(key).or_insert(font_id);
                    font_id += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        ix.fonts_count = font_id;

        // --- fills ---
        let mut rdr = Reader::from_reader(styles);
        rdr.config_mut().trim_text(true);
        let mut in_fills = false;
        let mut fill_id = 0u32;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"fills" => in_fills = true,
                Event::End(ref e) if e.name().as_ref() == b"fills" => {
                    // in_fills = false;
                    break;
                }
                Event::Start(ref e) if in_fills && e.name().as_ref() == b"fill" => {
                    // считываем <fill> целиком; учитываем только solid fgColor rgb
                    let mut depth = 1;
                    let mut rgb: Option<String> = None;

                    while depth > 0 {
                        match rdr.read_event()? {
                            Event::Start(ref fe) => {
                                if fe.name().as_ref() == b"fgColor" {
                                    for a in fe.attributes().with_checks(false).flatten() {
                                        if a.key.as_ref() == b"rgb" {
                                            let mut v =
                                                String::from_utf8_lossy(&a.value).into_owned();
                                            v.make_ascii_uppercase();
                                            rgb = Some(v);
                                        }
                                    }
                                }
                                depth += 1;
                            }
                            Event::Empty(ref fe) => {
                                if fe.name().as_ref() == b"fgColor" {
                                    for a in fe.attributes().with_checks(false).flatten() {
                                        if a.key.as_ref() == b"rgb" {
                                            let mut v =
                                                String::from_utf8_lossy(&a.value).into_owned();
                                            v.make_ascii_uppercase();
                                            rgb = Some(v);
                                        }
                                    }
                                }
                                // depth не меняем
                            }
                            Event::End(_) => depth -= 1,
                            Event::Eof => break,
                            _ => {}
                        }
                    }

                    if let Some(k) = rgb {
                        ix.fill_by_rgb.entry(k).or_insert(fill_id);
                    }
                    fill_id += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        ix.fills_count = fill_id;

        // --- borders ---
        let mut rdr = Reader::from_reader(styles);
        rdr.config_mut().trim_text(true);
        let mut in_borders = false;
        let mut border_id = 0u32;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"borders" => in_borders = true,
                Event::End(ref e) if e.name().as_ref() == b"borders" => {
                    // in_borders = false;
                    break;
                }
                Event::Start(ref e) if in_borders && e.name().as_ref() == b"border" => {
                    let mut depth = 1;
                    let mut styles = [None, None, None, None]; // left,right,top,bottom

                    while depth > 0 {
                        match rdr.read_event()? {
                            Event::Start(ref be) => {
                                let side = match be.name().as_ref() {
                                    b"left" => Some(0),
                                    b"right" => Some(1),
                                    b"top" => Some(2),
                                    b"bottom" => Some(3),
                                    _ => None,
                                };
                                if let Some(i) = side {
                                    for a in be.attributes().with_checks(false).flatten() {
                                        if a.key.as_ref() == b"style" {
                                            styles[i] = Some(
                                                String::from_utf8_lossy(&a.value).into_owned(),
                                            );
                                        }
                                    }
                                }
                                depth += 1;
                            }
                            Event::Empty(ref be) => {
                                let side = match be.name().as_ref() {
                                    b"left" => Some(0),
                                    b"right" => Some(1),
                                    b"top" => Some(2),
                                    b"bottom" => Some(3),
                                    _ => None,
                                };
                                if let Some(i) = side {
                                    for a in be.attributes().with_checks(false).flatten() {
                                        if a.key.as_ref() == b"style" {
                                            styles[i] = Some(
                                                String::from_utf8_lossy(&a.value).into_owned(),
                                            );
                                        }
                                    }
                                }
                            }

                            Event::End(_) => depth -= 1,
                            Event::Eof => break,
                            _ => {}
                        }
                    }

                    // ключ только для «ровных» рамок, иначе пропускаем (мы такие и создаём)
                    if let (Some(l), Some(r), Some(t), Some(b)) =
                        (&styles[0], &styles[1], &styles[2], &styles[3])
                    {
                        if l == r && r == t && t == b {
                            ix.border_by_key.entry(l.clone()).or_insert(border_id);
                        }
                    }
                    border_id += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        ix.borders_count = border_id;

        // --- cellXfs ---
        let mut rdr = Reader::from_reader(styles);
        rdr.config_mut().trim_text(true);
        let mut in_xfs = false;
        let mut xf_id = 0u32;

        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"cellXfs" => in_xfs = true,
                Event::End(ref e) if e.name().as_ref() == b"cellXfs" => {
                    // in_xfs = false;
                    break;
                }

                Event::Start(ref e) | Event::Empty(ref e)
                    if in_xfs && e.name().as_ref() == b"xf" =>
                {
                    let mut num_fmt_id = 0u32;
                    let mut font_id: Option<u32> = None;
                    let mut fill_id: Option<u32> = None;
                    let mut border_id: Option<u32> = None;

                    for a in e.attributes().with_checks(false).flatten() {
                        match a.key.as_ref() {
                            b"numFmtId" => num_fmt_id = lexical_core::parse(&a.value).unwrap_or(0),
                            b"fontId" => font_id = Some(lexical_core::parse(&a.value).unwrap_or(0)),
                            b"fillId" => fill_id = Some(lexical_core::parse(&a.value).unwrap_or(0)),
                            b"borderId" => {
                                border_id = Some(lexical_core::parse(&a.value).unwrap_or(0))
                            }
                            _ => {}
                        }
                    }

                    // выцепим alignment (если есть)
                    let mut align: Option<AlignSpec> = None;
                    if matches!(ev, Event::Start(_)) {
                        let mut depth = 1;
                        while depth > 0 {
                            match rdr.read_event()? {
                                Event::Start(ref ae) => {
                                    depth += 1;
                                    if ae.name().as_ref() == b"alignment" {
                                        let mut spec = AlignSpec::default();
                                        for a in ae.attributes().with_checks(false).flatten() {
                                            let v = String::from_utf8_lossy(&a.value).into_owned();
                                            match a.key.as_ref() {
                                                b"horizontal" => spec.horiz = Some(v.parse()?),
                                                b"vertical" => spec.vert = Some(v.parse()?),
                                                b"wrapText" => {
                                                    if v == "1" {
                                                        spec.wrap = true
                                                    }
                                                }
                                                _ => {}
                                            }
                                        }
                                        align = Some(spec);
                                    }
                                }
                                Event::End(_) => depth -= 1,
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                    }

                    ix.xfs.push(XfParts {
                        num_fmt_id,
                        font_id,
                        fill_id,
                        border_id,
                        align: align.clone(),
                    });

                    let sk = StyleKey {
                        num_fmt_id,
                        font_id,
                        fill_id,
                        border_id,
                        align: align
                            .as_ref()
                            .map(|a| (a.horiz.clone(), a.vert.clone(), a.wrap)),
                    };
                    ix.xf_by_key.entry(sk).or_insert(xf_id);
                    xf_id += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }

        Ok(ix)
    }
}

impl XlsxEditor {
    fn style_ix_mut(&mut self) -> Result<&mut StyleIndex> {
        if self.styles_index.is_none() {
            let ix = StyleIndex::build(&self.styles_xml)?;
            self.styles_index = Some(ix);
        }
        Ok(self.styles_index.as_mut().unwrap())
    }

    fn invalidate_styles_ix(&mut self) {
        self.styles_index = None;
    }
}

/* ========================== PUBLIC API ==================================== */

impl XlsxEditor {
    pub fn set_border(&mut self, range: &str, border_style: &str) -> Result<&mut Self> {
        let border_id = self.ensure_border(border_style)?;
        self.apply_patch(
            range,
            StyleParts {
                border: Some(border_id),
                ..Default::default()
            },
        )?;
        Ok(self)
    }

    pub fn set_font(
        &mut self,
        range: &str,
        name: &str,
        size: f32,
        bold: bool,
        italic: bool,
    ) -> Result<&mut Self> {
        let font_id = self.ensure_font(name, size, bold, italic)?;
        self.apply_patch(
            range,
            StyleParts {
                font: Some(font_id),
                ..Default::default()
            },
        )?;
        Ok(self)
    }

    pub fn set_font_with_alignment(
        &mut self,
        range: &str,
        name: &str,
        size: f32,
        bold: bool,
        italic: bool,
        align: &AlignSpec,
    ) -> Result<&mut Self> {
        let font_id = self.ensure_font(name, size, bold, italic)?;
        self.apply_patch(
            range,
            StyleParts {
                font: Some(font_id),
                align: Some(align.clone()),
                ..Default::default()
            },
        )?;
        Ok(self)
    }

    pub fn set_fill(&mut self, range: &str, rgb: &str) -> Result<&mut Self> {
        let fill_id = self.ensure_fill(rgb)?;
        self.apply_patch(
            range,
            StyleParts {
                fill: Some(fill_id),
                ..Default::default()
            },
        )?;
        Ok(self)
    }

    pub fn set_alignment(&mut self, range: &str, align: &AlignSpec) -> Result<&mut Self> {
        if let Some((col0, row_start)) = parse_open_column_selector(range) {
            let patch = StyleParts {
                align: Some(align.clone()),
                ..Default::default()
            };
            self.apply_patch_col_one_pass(col0, row_start, &patch)?;
            if row_start == 1 {
                let default_sid = self.ensure_style_from_parts(&patch)?;
                self.set_column_properties(col0, None, Some(default_sid))?;
            }
            return Ok(self);
        }

        self.apply_patch(
            range,
            StyleParts {
                align: Some(align.clone()),
                ..Default::default()
            },
        )?;
        Ok(self)
    }
    pub fn remove_style(&mut self, range: &str) -> Result<&mut Self> {
        match parse_target(range)? {
            Target::Cell(cell) => self.remove_style_from_cell(&cell)?,
            Target::Rect { c0, r0, c1, r1 } => self.remove_style_rect(c0, r0, c1, r1)?,
            Target::Col(c) => self.remove_style_col(c)?,
            Target::Row(r) => self.remove_style_row(r)?,
        }
        Ok(self)
    }

    /// Публичный API для числового формата.
    pub fn set_number_format(&mut self, range: &str, fmt: &str) -> Result<()> {
        let style_id = self.ensure_style(Some(fmt), None, None, None, None)?;
        match parse_target(range)? {
            Target::Cell(c) => self.apply_style_to_cell(&c, style_id)?,
            Target::Rect { c0, r0, c1, r1 } => {
                for r in r0..=r1 {
                    for c in c0..=c1 {
                        let coord = format!("{}{}", col_letter(c), r);
                        self.apply_style_to_cell(&coord, style_id)?;
                    }
                }
            }
            Target::Col(c0) => self.force_column_number_format(c0, style_id)?,
            Target::Row(_row) => bail!("Row-level not implemented yet"),
        }
        Ok(())
    }

    pub fn set_column_width(&mut self, col_letter: &str, width: f64) -> Result<&mut Self> {
        let col0 = col_index(col_letter) as u32; // 0-based
        self.set_column_properties(col0, Some(width), None)?;
        Ok(self)
    }
}

/* ========================== CORE PATCH ENGINE ============================= */

impl XlsxEditor {
    #[inline]
    fn get_or_make_sid(
        &mut self,
        cache: &mut HashMap<Option<u32>, u32>,
        old_sid: Option<u32>,
        patch: &StyleParts,
    ) -> u32 {
        if let Some(&sid) = cache.get(&old_sid) {
            return sid;
        }
        let old_parts = self.read_style_parts(old_sid).unwrap();
        let merged = merge_style_parts(old_parts, patch);
        let sid = self.ensure_style_from_parts(&merged).unwrap();
        cache.insert(old_sid, sid);
        sid
    }

    /// Быстрый однопроходный патч диапазона: правит стиль только у существующих <c ...>.
    fn apply_patch_rect_one_pass(
        &mut self,
        c0: u32,
        r0: u32,
        c1: u32,
        r1: u32,
        patch: &StyleParts,
    ) -> Result<()> {
        let mut sid_cache: HashMap<Option<u32>, u32> = HashMap::new();
        let (range_col_start, range_col_end) = if c0 <= c1 { (c0, c1) } else { (c1, c0) };
        let (range_row_start, range_row_end) = if r0 <= r1 { (r0, r1) } else { (r1, r0) };
        let range_width = (range_col_end - range_col_start) as usize + 1;
        let mut coverage = vec![0u8; range_width];
        let mut next_row_needed = range_row_start;

        // забираем исходный буфер, чтобы свободно писать новый
        let src = std::mem::take(&mut self.sheet_xml);
        let mut dst = Vec::with_capacity(src.len() + 512);

        let find_row = memmem::Finder::new(b"<row ");
        let find_cell_open = memmem::Finder::new(b"<c ");
        let find_cell_selfclose = memmem::Finder::new(b"<c/");
        let find_gt = memmem::Finder::new(b">");

        let mut i = 0usize;

        while let Some(off) = find_row.find(&src[i..]) {
            let row_start = i + off;
            // всё до <row ...> — как есть
            dst.extend_from_slice(&src[i..row_start]);

            // конец открывающего тега <row ...>
            let row_tag_end =
                find_gt.find(&src[row_start..]).context("malformed <row>")? + row_start;

            // r="...":
            let mut row_r: Option<u32> = None;
            if let Some(pos) = find_bytes_from(&src, b" r=\"", row_start) {
                if pos < row_tag_end {
                    let v0 = pos + 4;
                    if let Some(v1) = find_bytes_from(&src, b"\"", v0) {
                        row_r = lexical_core::parse::<u32>(&src[v0..v1]).ok();
                    }
                }
            }

            // границы строки
            let row_end =
                find_bytes_from(&src, b"</row>", row_tag_end).context("</row> not found")?;
            let row_close_end = row_end + "</row>".len();

            let Some(cur_row) = row_r else {
                // нет номера — просто копируем как есть
                dst.extend_from_slice(&src[row_start..row_close_end]);
                i = row_close_end;
                continue;
            };

            if next_row_needed <= range_row_end && cur_row > range_row_end {
                while next_row_needed <= range_row_end {
                    self.append_full_row_for_range(
                        &mut dst,
                        next_row_needed,
                        range_col_start,
                        range_col_end,
                        &mut sid_cache,
                        patch,
                    );
                    next_row_needed += 1;
                }
            }

            if cur_row < range_row_start || cur_row > range_row_end {
                // вне диапазона — копируем как есть
                dst.extend_from_slice(&src[row_start..row_close_end]);
                i = row_close_end;
                continue;
            }

            if !coverage.is_empty() {
                coverage.fill(0);
            }
            while next_row_needed < cur_row {
                self.append_full_row_for_range(
                    &mut dst,
                    next_row_needed,
                    range_col_start,
                    range_col_end,
                    &mut sid_cache,
                    patch,
                );
                next_row_needed += 1;
            }

            // строка в диапазоне — копируем заголовок <row ...>
            dst.extend_from_slice(&src[row_start..=row_tag_end]);

            // обрабатываем содержимое до </row>
            let mut j = row_tag_end + 1;
            while j < row_end {
                // следующий <c ...> (включая самозакрывающийся)
                let next_open = find_cell_open.find(&src[j..]).map(|p| j + p);
                let next_sc = find_cell_selfclose.find(&src[j..]).map(|p| j + p);
                let next_cell = match (next_open, next_sc) {
                    (Some(a), Some(b)) => Some(a.min(b)),
                    (Some(a), None) => Some(a),
                    (None, Some(b)) => Some(b),
                    (None, None) => None,
                };

                match next_cell {
                    None => {
                        dst.extend_from_slice(&src[j..row_end]);
                        break;
                    }
                    Some(cpos) if cpos >= row_end => {
                        dst.extend_from_slice(&src[j..row_end]);
                        break;
                    }
                    Some(cpos) => {
                        // всё до <c...>
                        dst.extend_from_slice(&src[j..cpos]);

                        // граница тега
                        let tag_end = find_gt.find(&src[cpos..]).context("cell tag end")? + cpos;
                        let self_closing = tag_end >= 1 && src[tag_end - 1] == b'/';

                        // локальная копия тега для правки атрибутов
                        let mut cell_tag = src[cpos..=tag_end].to_vec();

                        // r="A12" → проверяем колонку
                        let mut col_idx = None;
                        if let Some(rpos) = find_bytes_from(&cell_tag, b" r=\"", 0) {
                            let v0 = rpos + 4;
                            if let Some(v1) = find_bytes_from(&cell_tag, b"\"", v0) {
                                let val = &cell_tag[v0..v1];
                                if let Some(p) = val.iter().position(|b| b.is_ascii_digit()) {
                                    let mut ci: u32 = 0;
                                    for &b in &val[..p] {
                                        let u = (b as char).to_ascii_uppercase() as u8;
                                        ci = ci * 26 + ((u - b'A') as u32 + 1);
                                    }
                                    col_idx = Some(ci - 1);
                                }
                            }
                        }

                        let mut col_in_range = false;
                        if let Some(ci0) = col_idx {
                            if ci0 >= range_col_start && ci0 <= range_col_end {
                                col_in_range = true;
                                if !coverage.is_empty() {
                                    coverage[(ci0 - range_col_start) as usize] = 1;
                                }
                            }
                        }

                        if col_in_range {
                            // старый s=".."
                            let old_sid = if let Some(sp) = find_bytes_from(&cell_tag, b" s=\"", 0)
                            {
                                let s0 = sp + 4;
                                let s1 = find_bytes_from(&cell_tag, b"\"", s0 + 1)
                                    .context("attr quote")?;
                                lexical_core::parse::<u32>(&cell_tag[s0..s1]).ok()
                            } else {
                                None
                            };

                            let new_sid = self.get_or_make_sid(&mut sid_cache, old_sid, patch);

                            // заменить/вставить s="..."
                            if let Some(sp) = find_bytes_from(&cell_tag, b" s=\"", 0) {
                                let s0 = sp + 4;
                                let s1 = find_bytes_from(&cell_tag, b"\"", s0 + 1)
                                    .context("attr quote")?;
                                cell_tag.splice(s0..s1, new_sid.to_string().bytes());
                            } else {
                                let ins = if self_closing {
                                    cell_tag.len() - 2
                                } else {
                                    cell_tag.len() - 1
                                };
                                cell_tag.splice(ins..ins, format!(r#" s="{}""#, new_sid).bytes());
                            }
                        }

                        // пишем тег
                        dst.extend_from_slice(&cell_tag);

                        if self_closing {
                            j = tag_end + 1;
                        } else {
                            // копируем содержимое ячейки до </c>
                            let c_close = find_bytes_from(&src, b"</c>", tag_end + 1)
                                .context("</c> missing")?;
                            dst.extend_from_slice(&src[tag_end + 1..=c_close + 3]);
                            j = c_close + 4;
                        }
                    }
                }
            }
            for (offset, flag) in coverage.iter().enumerate() {
                if *flag == 0 {
                    let col_idx = range_col_start + offset as u32;
                    let coord = format!("{}{}", col_letter(col_idx), cur_row);
                    let sid = self.get_or_make_sid(&mut sid_cache, None, patch);
                    let cell_tag = format!(r#"<c r="{}" s="{}"/>"#, coord, sid);
                    dst.extend_from_slice(cell_tag.as_bytes());
                }
            }

            // закрытие строки
            dst.extend_from_slice(&src[row_end..row_close_end]);
            i = row_close_end;

            next_row_needed = cur_row + 1;
        }

        while next_row_needed <= range_row_end {
            self.append_full_row_for_range(
                &mut dst,
                next_row_needed,
                range_col_start,
                range_col_end,
                &mut sid_cache,
                patch,
            );
            next_row_needed += 1;
        }

        // хвост документа
        dst.extend_from_slice(&src[i..]);
        self.sheet_xml = dst;
        Ok(())
    }

    fn append_full_row_for_range(
        &mut self,
        dst: &mut Vec<u8>,
        row: u32,
        col_start: u32,
        col_end: u32,
        sid_cache: &mut HashMap<Option<u32>, u32>,
        patch: &StyleParts,
    ) {
        let sid = self.get_or_make_sid(sid_cache, None, patch);
        dst.extend_from_slice(format!(r#"<row r="{}">"#, row).as_bytes());
        for col_idx in col_start..=col_end {
            let coord = format!("{}{}", col_letter(col_idx), row);
            let cell_tag = format!(r#"<c r="{}" s="{}"/>"#, coord, sid);
            dst.extend_from_slice(cell_tag.as_bytes());
        }
        dst.extend_from_slice(b"</row>");
    }

    fn apply_patch_col_one_pass(
        &mut self,
        col0: u32,
        row_start: u32,
        patch: &StyleParts,
    ) -> Result<()> {
        let mut sid_cache: HashMap<Option<u32>, u32> = HashMap::new();
        let col = col_letter(col0).to_ascii_uppercase();
        let col_bytes = col.as_bytes();

        let src = std::mem::take(&mut self.sheet_xml);
        let mut dst = Vec::with_capacity(src.len() + 256);

        let mut i = 0usize;
        let finder = memmem::Finder::new(b"<c");

        while let Some(off) = finder.find(&src[i..]) {
            let start = i + off;
            dst.extend_from_slice(&src[i..start]);

            let next = *src.get(start + 2).unwrap_or(&b'>');
            let is_cell = matches!(next, b' ' | b'>' | b'/' | b'r' | b's' | b't');
            let tag_end = find_bytes_from(&src, b">", start).context("cell tag end")? + 1;

            if !is_cell {
                dst.extend_from_slice(&src[start..tag_end]);
                i = tag_end;
                continue;
            }

            let mut cell = src[start..tag_end].to_vec();
            let mut in_target_col = false;

            if let Some(rpos) = find_bytes_from(&cell, b" r=\"", 0) {
                let v0 = rpos + 4;
                if let Some(v1) = find_bytes_from(&cell, b"\"", v0) {
                    let val = &cell[v0..v1];
                    let col_len = col_bytes.len();
                    if val.len() > col_len
                        && val[..col_len]
                            .iter()
                            .map(|b| b.to_ascii_uppercase())
                            .eq(col_bytes.iter().copied())
                        && val[col_len..].iter().all(|b| b.is_ascii_digit())
                    {
                        if let Ok(row_num) = lexical_core::parse::<u32>(&val[col_len..]) {
                            if row_num >= row_start {
                                in_target_col = true;
                            }
                        }
                    }
                }
            }

            if in_target_col {
                let old_sid = if let Some(sp) = find_bytes_from(&cell, b" s=\"", 0) {
                    let s0 = sp + 4;
                    let s1 = find_bytes_from(&cell, b"\"", s0 + 1).context("style quote")?;
                    lexical_core::parse::<u32>(&cell[s0..s1]).ok()
                } else {
                    None
                };
                let new_sid = self.get_or_make_sid(&mut sid_cache, old_sid, patch);

                if let Some(sp) = find_bytes_from(&cell, b" s=\"", 0) {
                    let s0 = sp + 4;
                    let s1 = find_bytes_from(&cell, b"\"", s0 + 1).context("style quote")?;
                    cell.splice(s0..s1, new_sid.to_string().bytes());
                } else {
                    let ins = if cell.len() >= 2 && cell[cell.len() - 2] == b'/' {
                        cell.len() - 2
                    } else {
                        cell.len() - 1
                    };
                    cell.splice(ins..ins, format!(r#" s="{}""#, new_sid).bytes());
                }
            }

            dst.extend_from_slice(&cell);
            i = tag_end;
        }

        dst.extend_from_slice(&src[i..]);
        self.sheet_xml = dst;
        Ok(())
    }
}

impl XlsxEditor {
    fn apply_patch(&mut self, range: &str, patch: StyleParts) -> Result<()> {
        let mut sid_cache: HashMap<Option<u32>, u32> = HashMap::new();

        match parse_target(range)? {
            Target::Cell(cell) => {
                let sid = self.cell_style_id(&cell)?;
                let new_sid = *sid_cache.entry(sid).or_insert_with(|| {
                    let old = self.read_style_parts(sid).unwrap();
                    let merged = merge_style_parts(old, &patch);
                    self.ensure_style_from_parts(&merged).unwrap()
                });
                self.apply_style_to_cell(&cell, new_sid)?;
            }
            Target::Rect { c0, r0, c1, r1 } => {
                // 1) сначала гарантируем наличие <c r=".."> в диапазоне
                // self.ensure_rect_cells_exist(c0, r0, c1, r1)?;
                // 2) теперь быстрый проход по существующим c-тегам с мерджем стиля
                self.apply_patch_rect_one_pass(c0, r0, c1, r1, &patch)?;
            }
            _ => bail!("Row/Col-level styling not implemented in this snippet"),
        }
        Ok(())
    }

    fn read_style_parts(&self, style_id: Option<u32>) -> Result<StyleParts> {
        if let Some(sid) = style_id {
            let (font, fill) = self.xf_components(sid)?;
            let border = self.xf_border(sid)?;
            let align = self.xf_alignment(sid)?;
            Ok(StyleParts {
                num_fmt_code: None,
                font,
                fill,
                border,
                align,
            })
        } else {
            Ok(StyleParts::default())
        }
    }
}

impl XlsxEditor {
    fn ensure_style_from_parts(&mut self, parts: &StyleParts) -> Result<u32> {
        // 1) numFmtId сначала (чтобы не держать &mut индекса)
        let num_fmt_id = if let Some(code) = parts.num_fmt_code.as_deref() {
            self.ensure_num_fmt(code)?
        } else {
            0
        };

        let font_id = parts.font;
        let fill_id = parts.fill;
        let border_id = parts.border;
        let align_key = parts
            .align
            .as_ref()
            .map(|a| (a.horiz.clone(), a.vert.clone(), a.wrap));
        let sk = StyleKey {
            num_fmt_id,
            font_id,
            fill_id,
            border_id,
            align: align_key.clone(),
        };

        // 2) короткий мут-заимствование: проверяем кэш
        {
            let ix = self.style_ix_mut()?;
            if let Some(&sid) = ix.xf_by_key.get(&sk) {
                return Ok(sid);
            }
        }

        // 3) пишем новый <xf> в XML
        let sid = self.add_new_xf_cached(
            num_fmt_id,
            font_id,
            fill_id,
            border_id,
            parts.align.as_ref(),
        )?;

        // 4) короткий мут-заимствование: обновляем индекс
        {
            let ix = self.style_ix_mut()?;
            ix.xfs.push(XfParts {
                num_fmt_id,
                font_id,
                fill_id,
                border_id,
                align: parts.align.clone(),
            });
            ix.xf_by_key.insert(sk, sid);
        }

        Ok(sid)
    }

    fn add_new_xf_cached(
        &mut self,
        fmt_id: u32,
        font_id: Option<u32>,
        fill_id: Option<u32>,
        border_id: Option<u32>,
        align: Option<&AlignSpec>,
    ) -> Result<u32> {
        let mut xf = String::from("<xf xfId=\"0\" ");

        if let Some(fid) = font_id {
            xf.push_str(&format!(r#"fontId="{fid}" applyFont="1" "#));
        }
        if let Some(fid) = fill_id {
            xf.push_str(&format!(r#"fillId="{fid}" applyFill="1" "#));
        }
        if let Some(bid) = border_id {
            xf.push_str(&format!(r#"borderId="{bid}" applyBorder="1" "#));
        }

        xf.push_str(&format!(
            r#"numFmtId="{}"{} "#,
            fmt_id,
            if fmt_id != 0 {
                r#" applyNumberFormat="1""#
            } else {
                ""
            }
        ));
        if align.is_some() {
            xf.push_str(r#"applyAlignment="1" "#);
        }
        xf.pop();
        xf.push('>');

        if let Some(al) = align {
            if al.horiz.is_some() || al.vert.is_some() || al.wrap {
                xf.push_str("<alignment");
                if let Some(h) = &al.horiz {
                    xf.push_str(&format!(r#" horizontal="{}""#, h));
                }
                if let Some(v) = &al.vert {
                    xf.push_str(&format!(r#" vertical="{}""#, v));
                }
                if al.wrap {
                    xf.push_str(r#" wrapText="1""#);
                }
                xf.push_str("/>");
            }
        }
        xf.push_str("</xf>");

        let pos = memmem::rfind(&self.styles_xml, b"</cellXfs>")
            .context("styles.xml: </cellXfs> not found")?;
        self.styles_xml.splice(pos..pos, xf.bytes());
        bump_count(&mut self.styles_xml, b"<cellXfs", b"count=\"")?;

        // индекс нового — это текущее количество <xf> до вставки
        let sid = {
            // быстрый подсчёт: не будем читать XML, просто возьмём длину из индекса
            let ix = self.styles_index.as_ref().unwrap();
            ix.xfs.len() as u32
        };
        Ok(sid)
    }
}
// impl XlsxEditor {
//     fn cell_exists(&self, coord: &str) -> bool {
//         let tag = format!(r#"<c r="{coord}""#);
//         memmem::find(&self.sheet_xml, tag.as_bytes()).is_some()
//     }

//     fn ensure_rect_cells_exist(&mut self, c0: u32, r0: u32, c1: u32, r1: u32) -> Result<()> {
//         for r in r0..=r1 {
//             for c in c0..=c1 {
//                 let coord = format!("{}{}", col_letter(c), r);
//                 if !self.cell_exists(&coord) {
//                     // создаём ячейку (apply_style_to_cell сам создаст row/cell при отсутствии)
//                     self.apply_style_to_cell(&coord, 0)?;
//                 }
//             }
//         }
//         Ok(())
//     }
// }

fn merge_align(base: Option<AlignSpec>, patch: Option<AlignSpec>) -> Option<AlignSpec> {
    match (base, patch) {
        (b, None) => b,
        (None, Some(p)) => Some(p),
        (Some(mut b), Some(p)) => {
            if p.horiz.is_some() {
                b.horiz = p.horiz;
            }
            if p.vert.is_some() {
                b.vert = p.vert;
            }
            b.wrap = b.wrap || p.wrap; // wrap только «наращиваем»
            Some(b)
        }
    }
}

fn merge_style_parts(mut base: StyleParts, patch: &StyleParts) -> StyleParts {
    if patch.align.is_some() {
        base.align = merge_align(base.align, patch.align.clone());
    }

    if patch.num_fmt_code.is_some() {
        base.num_fmt_code = patch.num_fmt_code.clone();
    }
    if patch.font.is_some() {
        base.font = patch.font;
    }
    if patch.fill.is_some() {
        base.fill = patch.fill;
    }
    if patch.border.is_some() {
        base.border = patch.border;
    }
    if patch.align.is_some() {
        base.align = patch.align.clone();
    }
    base
}

/* ========================== LOW-LEVEL HELPERS ============================= */

impl XlsxEditor {
    fn ensure_style(
        &mut self,
        num_fmt: Option<&str>,
        font_id: Option<u32>,
        fill_id: Option<u32>,
        border_id: Option<u32>,
        align: Option<&AlignSpec>,
    ) -> Result<u32> {
        let fmt_id: u32 = if let Some(code) = num_fmt {
            self.ensure_num_fmt(code)?
        } else {
            0
        };

        if align.is_none() {
            if let Some(id) = self.find_matching_xf(fmt_id, font_id, fill_id, border_id)? {
                return Ok(id);
            }
        }

        self.add_new_xf(fmt_id, font_id, fill_id, border_id, align)
    }

    fn ensure_num_fmt(&mut self, code: &str) -> Result<u32> {
        // A) есть в кеше?
        if let Some(id) = self
            .styles_index
            .as_ref()
            .and_then(|ix| ix.numfmt_by_code.get(code).copied())
        {
            return Ok(id);
        }

        // B) узнать next id (нужно инициализировать индекс!)
        let new_id = {
            let ix = self.style_ix_mut()?;
            // ещё раз проверим на гонку
            if let Some(&id) = ix.numfmt_by_code.get(code) {
                return Ok(id);
            }
            ix.next_custom_numfmt // читаем, но не меняем здесь!
        };

        // C) правим XML
        let tag = format!(r#"<numFmt numFmtId="{new_id}" formatCode="{code}"/>"#);
        if let Some(end) = memmem::rfind(&self.styles_xml, b"</numFmts>") {
            // блок уже есть → просто дописываем внутрь и бампим count
            self.styles_xml.splice(end..end, tag.bytes());
            bump_count(&mut self.styles_xml, b"<numFmts", b"count=\"")?;
        } else {
            // блока нет → создаём РОВНО один
            let root = memmem::find(&self.styles_xml, b"<styleSheet")
                .context("<styleSheet> root not found in styles.xml")?;
            let after_root = find_bytes_from(&self.styles_xml, b">", root)
                .context("<styleSheet> start tag '>' not found")?
                + 1;

            // старайся соблюдать порядок: numFmts должен стоять до <fonts>
            let before_fonts = memmem::find(&self.styles_xml, b"<fonts").unwrap_or(after_root);

            let block = format!(r#"<numFmts count="1">{tag}</numFmts>"#);
            self.styles_xml
                .splice(before_fonts..before_fonts, block.bytes());

            // ← ВАЖНО: НЕ делать вторую вставку по insert..insert
        }
        // D) обновляем индекс
        {
            let ix = self.style_ix_mut()?;
            ix.numfmt_by_code.insert(code.to_string(), new_id);
            ix.next_custom_numfmt = new_id + 1;
        }

        Ok(new_id)
    }

    fn find_matching_xf(
        &self,
        fmt_id: u32,
        font_id: Option<u32>,
        fill_id: Option<u32>,
        border_id: Option<u32>,
    ) -> Result<Option<u32>> {
        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);

        let mut in_xfs = false;
        let mut idx: u32 = 0;

        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"cellXfs" => in_xfs = true,
                Event::End(ref e) if e.name().as_ref() == b"cellXfs" => in_xfs = false,

                Event::Start(ref e) | Event::Empty(ref e)
                    if in_xfs && e.name().as_ref() == b"xf" =>
                {
                    // С xf с alignment мы не сравниваем — пропускаем
                    let mut has_alignment_child = false;
                    // Event::Start -> значит дальше внутри могут быть теги
                    if matches!(ev, Event::Start(_)) {
                        let mut depth = 1;
                        while depth > 0 {
                            match rdr.read_event()? {
                                Event::Start(ref ie) => {
                                    if ie.name().as_ref() == b"alignment" {
                                        has_alignment_child = true;
                                    }
                                    depth += 1;
                                }
                                Event::End(_) => depth -= 1,
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                    }
                    if has_alignment_child {
                        idx += 1;
                        continue;
                    }

                    let mut num = None::<u32>;
                    let mut fnt = None::<u32>;
                    let mut fil = None::<u32>;
                    let mut bdr = None::<u32>;
                    for a in e.attributes().with_checks(false).flatten() {
                        match a.key.as_ref() {
                            b"numFmtId" => num = Some(lexical_core::parse(&a.value)?),
                            b"fontId" => fnt = Some(lexical_core::parse(&a.value)?),
                            b"fillId" => fil = Some(lexical_core::parse(&a.value)?),
                            b"borderId" => bdr = Some(lexical_core::parse(&a.value)?),
                            _ => {}
                        }
                    }
                    let num_ok = num.unwrap_or(0) == fmt_id;
                    let font_ok = font_id.map_or(true, |v| Some(v) == fnt);
                    let fill_ok = fill_id.map_or(true, |v| Some(v) == fil);
                    let border_ok = border_id.map_or(true, |v| Some(v) == bdr);

                    if num_ok && font_ok && fill_ok && border_ok {
                        return Ok(Some(idx));
                    }
                    idx += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(None)
    }

    fn add_new_xf(
        &mut self,
        fmt_id: u32,
        font_id: Option<u32>,
        fill_id: Option<u32>,
        border_id: Option<u32>,
        align: Option<&AlignSpec>,
    ) -> Result<u32> {
        let mut xf = String::from("<xf xfId=\"0\" ");

        if let Some(fid) = font_id {
            xf.push_str(&format!(r#"fontId="{fid}" applyFont="1" "#));
        }
        if let Some(fid) = fill_id {
            xf.push_str(&format!(r#"fillId="{fid}" applyFill="1" "#));
        }
        if let Some(bid) = border_id {
            xf.push_str(&format!(r#"borderId="{bid}" applyBorder="1" "#));
        }
        xf.push_str(&format!(
            r#"numFmtId="{}"{} "#,
            fmt_id,
            if fmt_id != 0 {
                r#" applyNumberFormat="1""#
            } else {
                ""
            }
        ));
        if align.is_some() {
            xf.push_str(r#"applyAlignment="1" "#);
        }
        xf.pop();
        xf.push('>');

        if let Some(al) = align {
            if al.horiz.is_some() || al.vert.is_some() || al.wrap {
                xf.push_str("<alignment");
                if let Some(h) = &al.horiz {
                    xf.push_str(&format!(r#" horizontal="{}""#, h));
                }
                if let Some(v) = &al.vert {
                    xf.push_str(&format!(r#" vertical="{}""#, v));
                }
                if al.wrap {
                    xf.push_str(r#" wrapText="1""#);
                }
                xf.push_str("/>");
            }
        }
        xf.push_str("</xf>");

        let pos = memmem::rfind(&self.styles_xml, b"</cellXfs>")
            .context("styles.xml: </cellXfs> not found")?;
        self.styles_xml.splice(pos..pos, xf.bytes());
        bump_count(&mut self.styles_xml, b"<cellXfs", b"count=\"")?;

        // посчитать индекс нового
        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);
        let mut in_xfs = false;
        let mut cnt = 0u32;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"cellXfs" => in_xfs = true,
                Event::End(ref e) if e.name().as_ref() == b"cellXfs" => break,
                Event::Start(ref e) | Event::Empty(ref e)
                    if in_xfs && e.name().as_ref() == b"xf" =>
                {
                    cnt += 1
                }
                Event::Eof => break,
                _ => {}
            }
        }
        self.invalidate_styles_ix(); // индекс устаревает — пересоберём при следующем обращении

        Ok(cnt - 1)
    }

    fn ensure_font(&mut self, name: &str, size: f32, bold: bool, italic: bool) -> Result<u32> {
        let key = FontKey {
            name: name.to_string(),
            size_100: (size * 100.0).round() as u32,
            bold,
            italic,
        };

        // 0) индекс/поиск
        {
            let ix = self.style_ix_mut()?;
            if let Some(&id) = ix.font_by_key.get(&key) {
                return Ok(id);
            }
        }

        // 1) id до вставки
        let new_id = {
            let ix = self.style_ix_mut()?;
            if let Some(&id) = ix.font_by_key.get(&key) {
                return Ok(id);
            }
            ix.fonts_count
        };

        // 2) XML
        let insert = memmem::rfind(&self.styles_xml, b"</fonts>")
            .context("<fonts> block not found in styles.xml")?;
        let mut xml = String::from("<font>");
        if bold {
            xml.push_str("<b/>");
        }
        if italic {
            xml.push_str("<i/>");
        }
        xml.push_str(&format!(r#"<sz val="{}"/>"#, (key.size_100 as f32) / 100.0));
        xml.push_str(&format!(r#"<name val="{}"/>"#, name));
        xml.push_str("</font>");
        self.styles_xml.splice(insert..insert, xml.bytes());
        bump_count(&mut self.styles_xml, b"<fonts", b"count=\"")?;

        // 3) индекс
        {
            let ix = self.style_ix_mut()?;
            ix.font_by_key.insert(key, new_id);
            ix.fonts_count = new_id + 1;
        }

        Ok(new_id)
    }

    fn ensure_fill(&mut self, rgb: &str) -> Result<u32> {
        let mut key = rgb.to_string();
        key.make_ascii_uppercase();

        // 0) индекс/поиск
        {
            let ix = self.style_ix_mut()?;
            if let Some(&id) = ix.fill_by_rgb.get(&key) {
                return Ok(id);
            }
        }

        // 1) id до вставки
        let new_id = {
            let ix = self.style_ix_mut()?;
            if let Some(&id) = ix.fill_by_rgb.get(&key) {
                return Ok(id);
            }
            ix.fills_count
        };

        // 2) XML
        let insert = memmem::rfind(&self.styles_xml, b"</fills>")
            .context("<fills> block not found in styles.xml")?;
        let xml = format!(
            r#"<fill><patternFill patternType="solid"><fgColor rgb="{key}"/><bgColor indexed="64"/></patternFill></fill>"#
        );
        self.styles_xml.splice(insert..insert, xml.bytes());
        bump_count(&mut self.styles_xml, b"<fills", b"count=\"")?;

        // 3) индекс
        {
            let ix = self.style_ix_mut()?;
            ix.fill_by_rgb.insert(key, new_id);
            ix.fills_count = new_id + 1;
        }

        Ok(new_id)
    }

    fn ensure_border(&mut self, style: &str) -> Result<u32> {
        // 0) Убедимся, что индекс инициализирован и попробуем найти готовый
        {
            let ix = self.style_ix_mut()?;
            if let Some(&id) = ix.border_by_key.get(style) {
                return Ok(id);
            }
        }

        // 1) Снимем "текущий" id ДО модификации XML
        let new_id = {
            let ix = self.style_ix_mut()?;
            // повторная проверка на случай гонки
            if let Some(&id) = ix.border_by_key.get(style) {
                return Ok(id);
            }
            ix.borders_count
        };

        // 2) Вставляем XML
        let end_pos = memmem::rfind(&self.styles_xml, b"</borders>")
            .context("styles.xml: </borders> not found")?;
        let tag = format!(
            r#"<border><left style="{s}"/><right style="{s}"/><top style="{s}"/><bottom style="{s}"/><diagonal/></border>"#,
            s = style
        );
        self.styles_xml.splice(end_pos..end_pos, tag.bytes());
        bump_count(&mut self.styles_xml, b"<borders", b"count=\"")?;

        // 3) Обновляем индекс ПОСЛЕ вставки, используя pre‑id
        {
            let ix = self.style_ix_mut()?;
            ix.border_by_key.insert(style.to_string(), new_id);
            ix.borders_count = new_id + 1;
        }

        Ok(new_id)
    }

    fn xf_components(&self, style_id: u32) -> Result<(Option<u32>, Option<u32>)> {
        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);
        let mut in_xfs = false;
        let mut idx = 0u32;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"cellXfs" => in_xfs = true,
                Event::End(ref e) if e.name().as_ref() == b"cellXfs" => break,
                Event::Start(ref e) | Event::Empty(ref e)
                    if in_xfs && e.name().as_ref() == b"xf" =>
                {
                    if idx == style_id {
                        let mut font = None;
                        let mut fill = None;
                        for a in e.attributes().with_checks(false).flatten() {
                            match a.key.as_ref() {
                                b"fontId" => font = Some(lexical_core::parse(&a.value)?),
                                b"fillId" => fill = Some(lexical_core::parse(&a.value)?),
                                _ => {}
                            }
                        }
                        return Ok((font, fill));
                    }
                    idx += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok((None, None))
    }

    fn xf_border(&self, style_id: u32) -> Result<Option<u32>> {
        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);
        let mut in_xfs = false;
        let mut idx = 0u32;
        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"cellXfs" => in_xfs = true,
                Event::End(ref e) if e.name().as_ref() == b"cellXfs" => break,
                Event::Start(ref e) | Event::Empty(ref e)
                    if in_xfs && e.name().as_ref() == b"xf" =>
                {
                    if idx == style_id {
                        for a in e.attributes().with_checks(false).flatten() {
                            if a.key.as_ref() == b"borderId" {
                                let val: u32 = lexical_core::parse(&a.value)?;
                                return Ok(Some(val));
                            }
                        }
                        return Ok(None);
                    }
                    idx += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(None)
    }

    fn xf_alignment(&self, style_id: u32) -> Result<Option<AlignSpec>> {
        let mut rdr = Reader::from_reader(self.styles_xml.as_slice());
        rdr.config_mut().trim_text(true);
        let mut in_xfs = false;
        let mut xf_idx = 0u32;
        // let mut depth = 0;

        while let Ok(ev) = rdr.read_event() {
            match ev {
                Event::Start(ref e) if e.name().as_ref() == b"cellXfs" => in_xfs = true,
                Event::End(ref e) if e.name().as_ref() == b"cellXfs" => break,

                Event::Start(ref e) if in_xfs && e.name().as_ref() == b"xf" => {
                    if xf_idx == style_id {
                        let mut depth = 1;
                        while depth > 0 {
                            match rdr.read_event()? {
                                Event::Start(ref ie) => {
                                    depth += 1;
                                    if ie.name().as_ref() == b"alignment" {
                                        let mut spec = AlignSpec::default();
                                        for attr in ie.attributes().with_checks(false).flatten() {
                                            let val =
                                                String::from_utf8_lossy(&attr.value).into_owned();
                                            match attr.key.as_ref() {
                                                b"horizontal" => spec.horiz = Some(val.parse()?),
                                                b"vertical" => spec.vert = Some(val.parse()?),
                                                b"wrapText" => {
                                                    if val == "1" {
                                                        spec.wrap = true
                                                    }
                                                }
                                                _ => {}
                                            }
                                        }
                                        return Ok(Some(spec));
                                    }
                                }
                                Event::End(_) => depth -= 1,
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                        return Ok(None);
                    }
                    xf_idx += 1;
                }
                Event::Empty(ref _e) if in_xfs => {
                    if xf_idx == style_id {
                        return Ok(None);
                    }
                    xf_idx += 1;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        Ok(None)
    }

    fn cell_style_id(&self, coord: &str) -> Result<Option<u32>> {
        let tag = format!(r#"<c r="{coord}""#);
        if let Some(pos) = memmem::rfind(&self.sheet_xml, tag.as_bytes()) {
            if let Some(spos) = find_bytes_from(&self.sheet_xml, b" s=\"", pos) {
                let val_start = spos + 4;
                let val_end =
                    find_bytes_from(&self.sheet_xml, b"\"", val_start + 1).unwrap_or(val_start);
                let id = std::str::from_utf8(&self.sheet_xml[val_start..val_end])?
                    .parse::<u32>()
                    .unwrap_or(0);
                return Ok(Some(id));
            }
        }
        Ok(None)
    }

    fn apply_style_to_cell(&mut self, coord: &str, style: u32) -> Result<()> {
        let row_num = coord.trim_start_matches(|c: char| c.is_ascii_alphabetic());
        let row_tag = format!(r#"<row r="{row_num}""#);

        let row_pos = match memmem::rfind(&self.sheet_xml, row_tag.as_bytes()) {
            Some(p) => p,
            None => {
                self.set_cell(coord, "")?;
                return self.apply_style_to_cell(coord, style);
            }
        };

        let row_end =
            find_bytes_from(&self.sheet_xml, b"</row>", row_pos).context("</row> not found")?;

        let cell_tag = format!(r#"<c r="{coord}""#);
        let cpos = match find_bytes_from(&self.sheet_xml, cell_tag.as_bytes(), row_pos) {
            Some(p) => p,
            None => {
                let new_cell = format!(r#"<c r="{coord}" s="{style}"/>"#);
                self.sheet_xml.splice(row_end..row_end, new_cell.bytes());
                return Ok(());
            }
        };

        let ctag_end = find_bytes_from(&self.sheet_xml, b">", cpos).context("malformed <c> tag")?;

        if let Some(sattr) = find_bytes_from(&self.sheet_xml, b" s=\"", cpos) {
            if sattr < ctag_end {
                let val_start = sattr + 4;
                let val_end = find_bytes_from(&self.sheet_xml, b"\"", val_start + 1)
                    .context("attr closing '\"' not found")?;
                self.sheet_xml
                    .splice(val_start..val_end, style.to_string().bytes());
                return Ok(());
            }
        }
        self.sheet_xml
            .splice(ctag_end..ctag_end, format!(r#" s="{style}""#).bytes());
        Ok(())
    }

    fn remove_style_from_cell(&mut self, coord: &str) -> Result<()> {
        let row_start = coord
            .find(|c: char| c.is_ascii_digit())
            .context("invalid cell coordinate – no digits")?;
        let row_num = &coord[row_start..];
        let row_tag = format!(r#"<row r="{row_num}""#);

        let Some(row_pos) = memmem::rfind(&self.sheet_xml, row_tag.as_bytes()) else {
            return Ok(());
        };

        let row_end =
            find_bytes_from(&self.sheet_xml, b"</row>", row_pos).context("</row> not found")?;

        let cell_tag = format!(r#"<c r="{coord}""#);
        let Some(cpos) = find_bytes_from(&self.sheet_xml, cell_tag.as_bytes(), row_pos) else {
            return Ok(());
        };
        if cpos >= row_end {
            return Ok(());
        }

        let ctag_end = find_bytes_from(&self.sheet_xml, b">", cpos).context("malformed <c> tag")?;

        if let Some(sattr) = find_bytes_from(&self.sheet_xml, b" s=\"", cpos) {
            if sattr < ctag_end {
                let val_end = find_bytes_from(&self.sheet_xml, b"\"", sattr + 4)
                    .context("style attribute not closed")?;
                let remove_end = val_end + 1;
                self.sheet_xml.drain(sattr..remove_end);
            }
        }
        Ok(())
    }

    fn remove_style_rect(&mut self, c0: u32, r0: u32, c1: u32, r1: u32) -> Result<()> {
        let (col_start, col_end) = if c0 <= c1 { (c0, c1) } else { (c1, c0) };
        let (row_start, row_end) = if r0 <= r1 { (r0, r1) } else { (r1, r0) };

        for r in row_start..=row_end {
            for c in col_start..=col_end {
                let coord = format!("{}{}", col_letter(c), r);
                self.remove_style_from_cell(&coord)?;
            }
        }
        Ok(())
    }
}

impl XlsxEditor {
    /// Удалить стиль из всех ячеек указанной строки (row-level syntax: "12:").
    fn remove_style_row(&mut self, row: u32) -> Result<()> {
        let src = std::mem::take(&mut self.sheet_xml);
        let mut dst = Vec::with_capacity(src.len());

        let find_row = memmem::Finder::new(b"<row ");
        let find_gt = memmem::Finder::new(b">");
        let find_cell_open = memmem::Finder::new(b"<c ");
        let find_cell_selfclose = memmem::Finder::new(b"<c/");

        let mut i = 0usize;

        while let Some(off) = find_row.find(&src[i..]) {
            let row_start = i + off;
            // всё до <row ...> — как есть
            dst.extend_from_slice(&src[i..row_start]);

            // конец открывающего тега <row ...>
            let row_tag_end =
                find_gt.find(&src[row_start..]).context("malformed <row>")? + row_start;

            // r="..."
            let mut row_r: Option<u32> = None;
            if let Some(pos) = find_bytes_from(&src, b" r=\"", row_start) {
                if pos < row_tag_end {
                    let v0 = pos + 4;
                    if let Some(v1) = find_bytes_from(&src, b"\"", v0) {
                        row_r = lexical_core::parse::<u32>(&src[v0..v1]).ok();
                    }
                }
            }

            let row_end =
                find_bytes_from(&src, b"</row>", row_tag_end).context("</row> not found")?;
            let row_close_end = row_end + "</row>".len();

            let Some(cur_row) = row_r else {
                // без номера — просто копируем
                dst.extend_from_slice(&src[row_start..row_close_end]);
                i = row_close_end;
                continue;
            };

            if cur_row != row {
                // чужая строка — копируем как есть
                dst.extend_from_slice(&src[row_start..row_close_end]);
                i = row_close_end;
                continue;
            }

            // Это наша строка → копируем заголовок <row ...>
            dst.extend_from_slice(&src[row_start..=row_tag_end]);

            // Обрабатываем содержимое до </row>, вычищая s="..."
            let mut j = row_tag_end + 1;
            while j < row_end {
                let next_open = find_cell_open.find(&src[j..]).map(|p| j + p);
                let next_sc = find_cell_selfclose.find(&src[j..]).map(|p| j + p);
                let next_cell = match (next_open, next_sc) {
                    (Some(a), Some(b)) => Some(a.min(b)),
                    (Some(a), None) => Some(a),
                    (None, Some(b)) => Some(b),
                    (None, None) => None,
                };

                match next_cell {
                    None => {
                        dst.extend_from_slice(&src[j..row_end]);
                        break;
                    }
                    Some(cpos) if cpos >= row_end => {
                        dst.extend_from_slice(&src[j..row_end]);
                        break;
                    }
                    Some(cpos) => {
                        // всё до <c ...>
                        dst.extend_from_slice(&src[j..cpos]);

                        let tag_end = find_gt.find(&src[cpos..]).context("cell tag end")? + cpos;
                        let self_closing = tag_end >= 1 && src[tag_end - 1] == b'/';

                        let mut cell_tag = src[cpos..=tag_end].to_vec();

                        // внутри тега убираем атрибут s="..."
                        if let Some(sattr) = find_bytes_from(&cell_tag, b" s=\"", 0) {
                            // ищем закрывающую кавычку
                            if let Some(val_end) = find_bytes_from(&cell_tag, b"\"", sattr + 4) {
                                // удаляем s="...": до включая кавычку
                                let remove_end = val_end + 1;
                                cell_tag.drain(sattr..remove_end);
                            }
                        }

                        dst.extend_from_slice(&cell_tag);

                        if self_closing {
                            j = tag_end + 1;
                        } else {
                            // содержимое ячейки до </c> — просто копируем
                            let c_close = find_bytes_from(&src, b"</c>", tag_end + 1)
                                .context("</c> missing")?;
                            dst.extend_from_slice(&src[tag_end + 1..=c_close + 3]);
                            j = c_close + 4;
                        }
                    }
                }
            }

            // закрытие строки
            dst.extend_from_slice(&src[row_end..row_close_end]);
            i = row_close_end;
        }

        // хвост (после последней строки)
        dst.extend_from_slice(&src[i..]);

        self.sheet_xml = dst;
        Ok(())
    }
}

impl XlsxEditor {
    /// Удалить стиль из всех ячеек указанного столбца (col-level syntax: "A:", "C:").
    /// col0 — 0-based индекс столбца.
    fn remove_style_col(&mut self, col0: u32) -> Result<()> {
        // Имя столбца в Excel-нотации ("A", "G", "AA", ...)
        let col = col_letter(col0);
        // Ищем атрибуты r="G...", r="AA..." и т.п.
        let pattern = format!(r#" r="{}"#, col);
        let pat = pattern.as_bytes();

        let src = std::mem::take(&mut self.sheet_xml);
        let mut dst = Vec::with_capacity(src.len());

        let finder = memmem::Finder::new(pat);

        let mut i: usize = 0;

        while let Some(off) = finder.find(&src[i..]) {
            let p = i + off; // позиция начала ` r="COL`

            // После r="COL" сразу должна идти цифра → иначе это не наш столбец
            let row_start_idx = p + pat.len();
            if row_start_idx >= src.len() || !src[row_start_idx].is_ascii_digit() {
                // Ложное совпадение (например, r="GA12" при col="A")
                dst.extend_from_slice(&src[i..=p]);
                i = p + 1;
                continue;
            }

            // Ищем начало тега <c ...> — ближайший '<' слева
            let cpos = match memrchr(b'<', &src[..p]) {
                Some(pos) if pos + 1 < src.len() && src[pos + 1] == b'c' => pos,
                _ => {
                    // На всякий случай: что-то странное → просто копируем и идём дальше
                    dst.extend_from_slice(&src[i..=p]);
                    i = p + 1;
                    continue;
                }
            };

            // Ищем конец тега <c ...>
            let tag_end_rel = match memchr(b'>', &src[cpos..]) {
                Some(x) => x,
                None => {
                    // Поломанный XML — копируем хвост и выходим
                    dst.extend_from_slice(&src[i..]);
                    self.sheet_xml = dst;
                    return Ok(());
                }
            };
            let tag_end = cpos + tag_end_rel;

            // Ищем внутри тега атрибут s="..."
            let s_finder = memmem::Finder::new(b" s=\"");
            let sattr_opt = s_finder.find(&src[cpos..tag_end]).map(|rel| cpos + rel);

            if let Some(sattr) = sattr_opt {
                let val_start = sattr + 4; // после ` s="`
                let val_end = find_bytes_from(&src, b"\"", val_start)
                    .context("style attribute not closed")?;

                if val_end > tag_end {
                    // Кривая разметка — не трогаем тег
                    dst.extend_from_slice(&src[i..=tag_end]);
                    i = tag_end + 1;
                    continue;
                }

                // Всё до атрибута s="..." копируем как есть
                dst.extend_from_slice(&src[i..sattr]);
                // Сам атрибут s="...": выкидываем, начиная с пробела и до закрывающей кавычки
                // т.е. [sattr ..= val_end] пропускаем
                i = val_end + 1;
            } else {
                // У клетки нет явного стиля — ничего не меняем
                dst.extend_from_slice(&src[i..=tag_end]);
                i = tag_end + 1;
            }
        }

        // Хвост после последнего совпадения
        dst.extend_from_slice(&src[i..]);

        self.sheet_xml = dst;
        Ok(())
    }
}
