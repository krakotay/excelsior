//! style.rs – универсальный слой стилей + нормализация <cols>

use anyhow::{Context, Result, bail};
use std::collections::HashMap;
use std::{fmt, str::FromStr};

use crate::files_part::needs_xml_space_preserve;
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

#[derive(Debug, Clone, PartialEq)]
pub struct FontSpec {
    pub name: String,
    pub size: f32,
    pub bold: bool,
    pub italic: bool,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct StyleSpec {
    pub num_fmt_code: Option<String>,
    pub font: Option<FontSpec>,
    pub fill_rgb: Option<String>,
    pub border_style: Option<String>,
    pub align: Option<AlignSpec>,
}

/* ========================== CORE STYLE STRUCT ============================= */

#[derive(Debug, Clone, Default)]
pub(crate) struct StyleParts {
    pub num_fmt_id: Option<u32>,
    pub num_fmt_code: Option<String>,
    pub font: Option<u32>,
    pub fill: Option<u32>,
    pub border: Option<u32>,
    pub align: Option<AlignSpec>,
}

/* ========================== TARGET PARSER ================================= */

#[derive(Debug)]
enum Target {
    All,
    Cell(String),
    Rect { c0: u32, r0: u32, c1: u32, r1: u32 },
    Col(u32), // 0-based
    Row(u32),
}

impl Target {
    fn matches(&self, col: u32, row: u32) -> bool {
        match self {
            Self::All => true,
            Self::Cell(c) => crate::data::parse_coord(c).is_ok_and(|v| v == (col as usize, row)),
            Self::Rect { c0, r0, c1, r1 } => {
                col >= (*c0).min(*c1)
                    && col <= (*c0).max(*c1)
                    && row >= (*r0).min(*r1)
                    && row <= (*r0).max(*r1)
            }
            Self::Col(c) => col == *c,
            Self::Row(r) => row == *r,
        }
    }
}

fn parse_target(s: &str) -> Result<Target> {
    let s = s.to_ascii_uppercase();
    if let Some((a, b)) = s.split_once(':') {
        if b.is_empty() {
            if a.bytes().all(|b| b.is_ascii_alphabetic()) {
                return Ok(Target::Col(crate::data::parse_column(a)? as u32));
            }
            let row: u32 = a.parse().context("Invalid row selector")?;
            anyhow::ensure!(
                (1..=crate::data::MAX_ROWS).contains(&row),
                "XLSX row limit exceeded"
            );
            return Ok(Target::Row(row));
        }
        let (c0, r0) = crate::data::parse_coord(a)?;
        let (c1, r1) = crate::data::parse_coord(b)?;
        return Ok(Target::Rect {
            c0: c0 as u32,
            r0,
            c1: c1 as u32,
            r1,
        });
    }
    crate::data::parse_coord(&s)?;
    Ok(Target::Cell(s))
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
        return crate::data::parse_column(head).ok().map(|c| (c as u32, 1));
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
    if row_start > crate::data::MAX_ROWS {
        return None;
    }
    Some((
        crate::data::parse_column(&head[..p]).ok()? as u32,
        row_start,
    ))
}

impl StyleIndex {
    fn build(styles: &[u8]) -> Result<Self> {
        use crate::xml;
        fn number(element: &xml::Element<'_>, name: &str) -> Result<Option<u32>> {
            Ok(element.attr(name)?.map(|v| v.parse()).transpose()?)
        }
        fn flag(element: &xml::Element<'_>) -> Result<bool> {
            Ok(!matches!(
                element.attr("val")?.as_deref(),
                Some("0" | "false")
            ))
        }
        let root = xml::root(styles)?;
        let mut ix = Self {
            xfs: Vec::new(),
            numfmt_by_code: HashMap::new(),
            numfmt_code_by_id: HashMap::new(),
            next_custom_numfmt: 164,
            font_by_key: HashMap::new(),
            fill_by_rgb: HashMap::new(),
            border_by_key: HashMap::new(),
            xf_by_key: HashMap::new(),
            fonts_count: 0,
            fills_count: 0,
            borders_count: 0,
        };
        seed_builtin_num_fmts(&mut ix);
        if let Some(formats) = root.child("numFmts")? {
            for entry in formats.children() {
                let entry = entry?;
                if !entry.is("numFmt", xml::SHEET_NS) {
                    continue;
                }
                let id = number(&entry, "numFmtId")?.context("numFmtId missing")?;
                let code = entry
                    .attr("formatCode")?
                    .context("formatCode missing")?
                    .into_owned();
                ix.numfmt_by_code.insert(code.clone(), id);
                ix.numfmt_code_by_id.insert(id, code);
                ix.next_custom_numfmt = ix.next_custom_numfmt.max(id + 1);
            }
        }
        if let Some(fonts) = root.child("fonts")? {
            for font in fonts.children() {
                let font = font?;
                if !font.is("font", xml::SHEET_NS) {
                    continue;
                }
                let name = font
                    .child("name")?
                    .map(|n| n.attr("val").map(|s| s.map(|s| s.into_owned())))
                    .transpose()?
                    .flatten()
                    .unwrap_or_else(|| "Calibri".to_owned());
                let size = font
                    .child("sz")?
                    .map(|n| {
                        n.attr("val").and_then(|s| {
                            s.map(|s| s.parse::<f32>().map_err(Into::into)).transpose()
                        })
                    })
                    .transpose()?
                    .flatten()
                    .unwrap_or(11.0);
                anyhow::ensure!(
                    size.is_finite() && size > 0.0,
                    "Invalid font size in styles.xml"
                );
                let bold = font
                    .child("b")?
                    .map(|v| flag(&v))
                    .transpose()?
                    .unwrap_or(false);
                let italic = font
                    .child("i")?
                    .map(|v| flag(&v))
                    .transpose()?
                    .unwrap_or(false);
                ix.font_by_key.insert(
                    FontKey {
                        name,
                        size_100: (size * 100.0).round() as u32,
                        bold,
                        italic,
                    },
                    ix.fonts_count,
                );
                ix.fonts_count += 1;
            }
        }
        if let Some(fills) = root.child("fills")? {
            for fill in fills.children() {
                let fill = fill?;
                if !fill.is("fill", xml::SHEET_NS) {
                    continue;
                }
                if let Some(pattern) = fill.child("patternFill")?
                    && let Some(color) = pattern.child("fgColor")?
                    && let Some(rgb) = color.attr("rgb")?
                {
                    ix.fill_by_rgb
                        .insert(rgb.to_ascii_uppercase(), ix.fills_count);
                }
                ix.fills_count += 1;
            }
        }
        if let Some(borders) = root.child("borders")? {
            for border in borders.children() {
                let border = border?;
                if !border.is("border", xml::SHEET_NS) {
                    continue;
                }
                let mut sides = Vec::new();
                let mut plain = true;
                for name in ["left", "right", "top", "bottom"] {
                    if let Some(side) = border.child(name)? {
                        plain &= side.child("color")?.is_none();
                        sides.push(side.attr("style")?.map(|v| v.into_owned()));
                    } else {
                        sides.push(None);
                    }
                }
                if plain
                    && sides.iter().all(|s| s == &sides[0])
                    && let Some(style) = &sides[0]
                {
                    ix.border_by_key.insert(style.clone(), ix.borders_count);
                }
                ix.borders_count += 1;
            }
        }
        let xfs = root.require_child("cellXfs")?;
        for xf in xfs.children() {
            let xf = xf?;
            if !xf.is("xf", xml::SHEET_NS) {
                continue;
            }
            let num_fmt_id = number(&xf, "numFmtId")?.unwrap_or(0);
            let font_id = number(&xf, "fontId")?;
            let fill_id = number(&xf, "fillId")?;
            let border_id = number(&xf, "borderId")?;
            let align = if let Some(alignment) = xf.child("alignment")? {
                Some(AlignSpec {
                    horiz: alignment
                        .attr("horizontal")?
                        .map(|v| v.parse())
                        .transpose()?,
                    vert: alignment.attr("vertical")?.map(|v| v.parse()).transpose()?,
                    wrap: matches!(alignment.attr("wrapText")?.as_deref(), Some("1" | "true")),
                })
            } else {
                None
            };
            let key = StyleKey {
                num_fmt_id,
                font_id,
                fill_id,
                border_id,
                align: align
                    .as_ref()
                    .map(|a| (a.horiz.clone(), a.vert.clone(), a.wrap)),
            };
            let sid = ix.xfs.len() as u32;
            ix.xf_by_key.insert(key, sid);
            ix.xfs.push(XfParts {
                num_fmt_id,
                font_id,
                fill_id,
                border_id,
                align,
            });
        }
        Ok(ix)
    }
}
impl XlsxEditor {
    fn style_ix_mut(&mut self) -> Result<&mut StyleIndex> {
        if self.styles_index.is_none() {
            self.styles_index = Some(StyleIndex::build(&self.styles_xml)?.into());
        }
        Ok(&mut *self.styles_index.as_mut().unwrap())
    }
}

/* ========================== PUBLIC API ==================================== */

impl XlsxEditor {
    pub fn set_style(&mut self, range: &str, spec: &StyleSpec) -> Result<&mut Self> {
        let patch = self.build_style_parts(spec)?;
        if patch.num_fmt_code.is_none()
            && patch.font.is_none()
            && patch.fill.is_none()
            && patch.border.is_none()
            && patch.align.is_none()
        {
            return Ok(self);
        }

        if let Some((col0, row_start)) = parse_open_column_selector(range) {
            if patch.num_fmt_code.is_some()
                || patch.font.is_some()
                || patch.fill.is_some()
                || patch.border.is_some()
            {
                bail!(
                    "open column style patch currently supports alignment or number format via dedicated APIs only"
                );
            }
            self.apply_patch_col_one_pass(col0, row_start, &patch)?;
            if row_start == 1 && patch.align.is_some() {
                let default_sid = self.ensure_style_from_parts(&patch)?;
                self.set_column_properties(col0, None, Some(default_sid))?;
            }
            return Ok(self);
        }

        self.apply_patch(range, patch)?;
        Ok(self)
    }

    fn build_style_parts(&mut self, spec: &StyleSpec) -> Result<StyleParts> {
        let font = if let Some(font) = &spec.font {
            Some(self.ensure_font(&font.name, font.size, font.bold, font.italic)?)
        } else {
            None
        };
        let fill = if let Some(rgb) = spec.fill_rgb.as_deref() {
            Some(self.ensure_fill(rgb)?)
        } else {
            None
        };
        let border = if let Some(style) = spec.border_style.as_deref() {
            Some(self.ensure_border(style)?)
        } else {
            None
        };

        Ok(StyleParts {
            num_fmt_id: None,
            num_fmt_code: spec.num_fmt_code.clone(),
            font,
            fill,
            border,
            align: spec.align.clone(),
        })
    }

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
            Target::All => unreachable!(),
            Target::Cell(cell) => self.remove_style_from_cell(&cell)?,
            Target::Rect { c0, r0, c1, r1 } => self.remove_style_rect(c0, r0, c1, r1)?,
            Target::Col(c) => self.remove_style_col(c)?,
            Target::Row(r) => self.remove_style_row(r)?,
        }
        Ok(self)
    }

    /// Публичный API для числового формата.
    pub fn set_number_format(&mut self, range: &str, fmt: &str) -> Result<()> {
        let patch = StyleParts {
            num_fmt_code: Some(fmt.to_string()),
            ..Default::default()
        };
        let target = parse_target(range)?;
        match &target {
            Target::All => unreachable!(),
            Target::Cell(c) => {
                let sid = self.cell_style_id(c)?;
                let merged = merge_style_parts(self.read_style_parts(sid)?, &patch);
                let style_id = self.ensure_style_from_parts(&merged)?;
                self.apply_style_to_cell(c, style_id)?;
            }
            Target::Rect { c0, r0, c1, r1 } => {
                self.apply_patch_rect_one_pass(*c0, *r0, *c1, *r1, &patch)?;
            }
            Target::Col(c0) => {
                self.apply_patch_col_one_pass(*c0, 1, &patch)?;
                let default_sid = self.ensure_style_from_parts(&patch)?;
                self.set_column_properties(*c0, None, Some(default_sid))?;
            }
            Target::Row(row) => self.apply_patch_row_one_pass(*row, &patch)?,
        }
        if fmt == "@" {
            self.convert_target_cells_to_inline_text(&target)?;
        }
        Ok(())
    }

    pub fn set_column_width(&mut self, col_letter: &str, width: f64) -> Result<&mut Self> {
        let col0 = crate::data::parse_column(col_letter)? as u32;
        anyhow::ensure!(
            width.is_finite() && width > 0.0 && width <= 255.0,
            "Column width must be in range (0, 255]"
        );
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
    ) -> Result<u32> {
        if let Some(&sid) = cache.get(&old_sid) {
            return Ok(sid);
        }
        let old_parts = self.read_style_parts(old_sid)?;
        let merged = merge_style_parts(old_parts, patch);
        let sid = self.ensure_style_from_parts(&merged)?;
        cache.insert(old_sid, sid);
        Ok(sid)
    }

    fn apply_patch_rect_one_pass(
        &mut self,
        c0: u32,
        r0: u32,
        c1: u32,
        r1: u32,
        patch: &StyleParts,
    ) -> Result<()> {
        self.edit_style_target(&Target::Rect { c0, r0, c1, r1 }, Some(patch), None, true)
    }
    fn apply_patch_col_one_pass(&mut self, col: u32, row: u32, patch: &StyleParts) -> Result<()> {
        self.edit_style_target(
            &Target::Rect {
                c0: col,
                r0: row,
                c1: col,
                r1: crate::data::MAX_ROWS,
            },
            Some(patch),
            None,
            false,
        )
    }
    fn apply_patch_row_one_pass(&mut self, row: u32, patch: &StyleParts) -> Result<()> {
        self.edit_style_target(&Target::Row(row), Some(patch), None, false)
    }
    fn edit_style_target(
        &mut self,
        target: &Target,
        patch: Option<&StyleParts>,
        fixed: Option<u32>,
        materialize: bool,
    ) -> Result<()> {
        let mut cache = HashMap::new();
        self.transform_cells(target, materialize, |editor, cell, _col, _row| {
            let old = cell.attr("s")?.map(|v| v.parse::<u32>()).transpose()?;
            let sid = if let Some(patch) = patch {
                Some(editor.get_or_make_sid(&mut cache, old, patch)?)
            } else {
                fixed
            };
            let value = sid.map(|v| v.to_string());
            if old == sid {
                return Ok(None);
            }
            let opening = cell.set_attr("s", value.as_deref())?;
            Ok(Some(crate::xml::patch_opening(cell, &opening)))
        })
    }
    fn transform_cells<F>(
        &mut self,
        target: &Target,
        materialize: bool,
        mut transform: F,
    ) -> Result<()>
    where
        F: FnMut(&mut Self, &crate::xml::Element<'_>, usize, u32) -> Result<Option<Vec<u8>>>,
    {
        let src = self.sheet_xml.clone();
        let data = crate::xml::require(&src, "sheetData")?;
        let rect = if materialize {
            match target {
                Target::Cell(c) => {
                    let (col, row) = crate::data::parse_coord(c)?;
                    Some((col as u32, row, col as u32, row))
                }
                Target::Rect { c0, r0, c1, r1 } => Some((
                    (*c0).min(*c1),
                    (*r0).min(*r1),
                    (*c0).max(*c1),
                    (*r0).max(*r1),
                )),
                _ => None,
            }
        } else {
            None
        };
        let mut content = crate::xml::begin_content(&src, &data)?;
        let mut next = rect.map(|(_, row, _, _)| row);
        let mut changed = false;
        for row in data.children() {
            let row = row?;
            anyhow::ensure!(
                row.is("row", crate::xml::SHEET_NS),
                "Unexpected element in sheetData"
            );
            let number: u32 = row.attr("r")?.context("Row number missing")?.parse()?;
            if let Some((_, _, _, last)) = rect {
                while next.is_some_and(|n| n < number && n <= last) {
                    let n = next.unwrap();
                    let empty = format!(
                        r#"<row xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="{n}"></row>"#
                    );
                    let element = crate::xml::root_in_context(empty.as_bytes(), &data)?;
                    content.extend_from_slice(
                        &self
                            .transform_row(&element, n, target, rect, &mut transform)?
                            .unwrap_or_else(|| empty.as_bytes().to_vec()),
                    );
                    next = Some(n + 1);
                    changed = true;
                }
            }
            if let Some(updated) = self.transform_row(&row, number, target, rect, &mut transform)? {
                content.extend_from_slice(&updated);
                changed = true;
            } else {
                content.extend_from_slice(row.bytes());
            }
            if next.is_some_and(|n| n <= number) {
                next = Some(number + 1);
            }
        }
        if let Some((_, _, _, last)) = rect {
            while next.is_some_and(|n| n <= last) {
                let n = next.unwrap();
                let empty = format!(
                    r#"<row xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="{n}"></row>"#
                );
                let element = crate::xml::root_in_context(empty.as_bytes(), &data)?;
                content.extend_from_slice(
                    &self
                        .transform_row(&element, n, target, rect, &mut transform)?
                        .unwrap_or_else(|| empty.as_bytes().to_vec()),
                );
                next = Some(n + 1);
                changed = true;
            }
        }
        if changed {
            crate::xml::end_content(&src, &data, &mut content)?;
            self.sheet_xml = content.into();
            self.sheet_data_end = None;
            self.row_index = None;
            if let Some((_, _, column, last)) = rect {
                self.last_row = self.last_row.max(last);
                if let Some((columns, rows)) = self.dimensions.get(&self.sheet_path).copied() {
                    self.dimensions.insert(
                        self.sheet_path.clone(),
                        (columns.max(column as usize + 1), rows.max(last)),
                    );
                }
            }
        }
        Ok(())
    }
    fn transform_row<F>(
        &mut self,
        row: &crate::xml::Element<'_>,
        number: u32,
        target: &Target,
        rect: Option<(u32, u32, u32, u32)>,
        transform: &mut F,
    ) -> Result<Option<Vec<u8>>>
    where
        F: FnMut(&mut Self, &crate::xml::Element<'_>, usize, u32) -> Result<Option<Vec<u8>>>,
    {
        let row_matches = match target {
            Target::Cell(coord) => crate::data::parse_coord(coord)?.1 == number,
            Target::Rect { r0, r1, .. } => number >= (*r0).min(*r1) && number <= (*r0).max(*r1),
            Target::Row(row) => *row == number,
            _ => true,
        };
        if !row_matches {
            return Ok(None);
        }
        if rect.is_none() {
            let mut result: Option<Vec<u8>> = None;
            let mut cursor = row.content.start;
            for cell in row.children() {
                let cell = cell?;
                if !cell.is("c", crate::xml::SHEET_NS) {
                    continue;
                }
                let coord = cell.attr("r")?.context("Cell coordinate missing")?;
                let (col, cell_row) = crate::data::parse_coord(&coord)?;
                anyhow::ensure!(cell_row == number, "Cell/row coordinate mismatch");
                if target.matches(col as u32, number)
                    && let Some(replacement) = transform(self, &cell, col, number)?
                {
                    let result = result.get_or_insert_with(|| Vec::with_capacity(row.full.len()));
                    result.extend_from_slice(row.segment(cursor..cell.full.start));
                    result.extend_from_slice(&replacement);
                    cursor = cell.full.end;
                }
            }
            return match result {
                Some(mut content) => {
                    content.extend_from_slice(row.segment(cursor..row.content.end));
                    Ok(Some(row.with_content(&content)?))
                }
                None => Ok(None),
            };
        }
        let mut cells = std::collections::BTreeMap::new();
        let mut others = Vec::new();
        for cell in row.children() {
            let cell = cell?;
            if cell.is("c", crate::xml::SHEET_NS) {
                let coord = cell.attr("r")?.context("Cell coordinate missing")?;
                let (col, cell_row) = crate::data::parse_coord(&coord)?;
                anyhow::ensure!(cell_row == number, "Cell/row coordinate mismatch");
                anyhow::ensure!(
                    cells.insert(col as u32, cell).is_none(),
                    "Duplicate cell coordinate"
                );
            } else {
                others.push(cell);
            }
        }
        let fill = rect
            .filter(|(_, r0, _, r1)| number >= *r0 && number <= *r1)
            .map(|(c0, _, c1, _)| c0..=c1);
        let mut missing = Vec::new();
        if let Some(fill) = fill {
            for col in fill {
                if !cells.contains_key(&col) {
                    missing.push(col);
                }
            }
        }
        let mut generated = Vec::with_capacity(missing.len());
        for col in missing {
            let tag = format!(r#"<c r="{}{number}"/>"#, col_letter(col));
            generated.push((
                col,
                crate::xml::child_fragment(row, tag.as_bytes())?.into_owned(),
            ));
        }
        for (col, bytes) in &generated {
            cells.insert(*col, crate::xml::root_in_context(bytes, row)?);
        }
        let mut content = Vec::new();
        let mut changed = !generated.is_empty();
        for (col, cell) in cells {
            if target.matches(col, number)
                && let Some(replacement) = transform(self, &cell, col as usize, number)?
            {
                content.extend_from_slice(&replacement);
                changed = true;
                continue;
            }
            content.extend_from_slice(cell.bytes());
        }
        if !changed {
            return Ok(None);
        }
        for other in others {
            content.extend_from_slice(other.bytes());
        }
        Ok(Some(row.with_content(&content)?))
    }
    fn convert_target_cells_to_inline_text(&mut self, target: &Target) -> Result<()> {
        self.transform_cells(target, false, |_, cell, _col, _row| {
            if !matches!(cell.attr("t")?.as_deref(), None | Some("n")) || cell.child("f")?.is_some()
            {
                return Ok(None);
            }
            let Some(value) = cell.child("v")? else {
                return Ok(None);
            };
            let text = crate::xml::text(&value)?;
            let mut writer = quick_xml::Writer::new(Vec::new());
            writer.create_element("is").write_inner_content(|w| {
                let mut t = w.create_element("t");
                if needs_xml_space_preserve(&text) {
                    t = t.with_attribute(("xml:space", "preserve"));
                }
                t.write_text_content(quick_xml::events::BytesText::new(&text))?;
                Ok(())
            })?;
            let replacement = cell.with_content(&writer.into_inner())?;
            let element = crate::xml::root(&replacement)?;
            let tag = element.set_attr("t", Some("inlineStr"))?;
            Ok(Some(crate::xml::patch_opening(&element, &tag)))
        })
    }
}

impl XlsxEditor {
    fn apply_patch(&mut self, range: &str, patch: StyleParts) -> Result<()> {
        let mut sid_cache: HashMap<Option<u32>, u32> = HashMap::new();

        match parse_target(range)? {
            Target::All => unreachable!(),
            Target::Cell(cell) => {
                let sid = self.cell_style_id(&cell)?;
                let new_sid = self.get_or_make_sid(&mut sid_cache, sid, &patch)?;
                self.apply_style_to_cell(&cell, new_sid)?;
            }
            Target::Rect { c0, r0, c1, r1 } => {
                // 1) сначала гарантируем наличие <c r=".."> в диапазоне
                // self.ensure_rect_cells_exist(c0, r0, c1, r1)?;
                // 2) теперь быстрый проход по существующим c-тегам с мерджем стиля
                self.apply_patch_rect_one_pass(c0, r0, c1, r1, &patch)?;
            }
            Target::Col(c0) => self.apply_patch_col_one_pass(c0, 1, &patch)?,
            Target::Row(row) => self.apply_patch_row_one_pass(row, &patch)?,
        }
        Ok(())
    }

    pub(crate) fn read_style_parts(&mut self, style_id: Option<u32>) -> Result<StyleParts> {
        let Some(sid) = style_id else {
            return Ok(StyleParts::default());
        };

        let (num_fmt_id, font_id, fill_id, border_id, align) = {
            let xf = self
                .style_ix_mut()?
                .xfs
                .get(sid as usize)
                .with_context(|| format!("style id {sid} not found in style index"))?;
            (
                xf.num_fmt_id,
                xf.font_id,
                xf.fill_id,
                xf.border_id,
                xf.align.clone(),
            )
        };

        let num_fmt_code = if num_fmt_id == 0 {
            None
        } else {
            self.style_ix_mut()?
                .numfmt_code_by_id
                .get(&num_fmt_id)
                .cloned()
        };

        Ok(StyleParts {
            num_fmt_id: Some(num_fmt_id),
            num_fmt_code,
            font: font_id,
            fill: fill_id,
            border: border_id,
            align,
        })
    }
}

impl XlsxEditor {
    fn ensure_style_from_parts(&mut self, parts: &StyleParts) -> Result<u32> {
        // 1) numFmtId сначала (чтобы не держать &mut индекса)
        let num_fmt_id = if let Some(code) = parts.num_fmt_code.as_deref() {
            self.ensure_num_fmt(code)?
        } else {
            parts.num_fmt_id.unwrap_or(0)
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
        let mut xf = format!(
            r#"<xf numFmtId="{fmt_id}" fontId="{}" fillId="{}" borderId="{}" xfId="0""#,
            font_id.unwrap_or(0),
            fill_id.unwrap_or(0),
            border_id.unwrap_or(0)
        );

        if fmt_id != 0 {
            xf.push_str(r#" applyNumberFormat="1""#);
        }
        if font_id.is_some() {
            xf.push_str(r#" applyFont="1""#);
        }
        if fill_id.is_some() {
            xf.push_str(r#" applyFill="1""#);
        }
        if border_id.is_some() {
            xf.push_str(r#" applyBorder="1""#);
        }
        if align.is_some() {
            xf.push_str(r#" applyAlignment="1""#);
        }

        if let Some(al) = align {
            if al.horiz.is_some() || al.vert.is_some() || al.wrap {
                xf.push('>');
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
                xf.push_str("</xf>");
            } else {
                xf.push_str("/>");
            }
        } else {
            xf.push_str("/>");
        }

        self.styles_xml =
            crate::xml::append_child(&self.styles_xml, "cellXfs", xf.as_bytes())?.into();
        crate::xml::bump_count(&mut self.styles_xml, "cellXfs")?;

        // индекс нового — это текущее количество <xf> до вставки
        let sid = {
            // быстрый подсчёт: не будем читать XML, просто возьмём длину из индекса
            let ix = self.styles_index.as_ref().unwrap();
            ix.xfs.len() as u32
        };
        Ok(sid)
    }
}

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
        base.num_fmt_id = None;
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

pub(crate) fn builtin_num_fmt_id(code: &str) -> Option<u32> {
    match code {
        "General" => Some(0),
        "0" => Some(1),
        "0.00" => Some(2),
        "#,##0" => Some(3),
        "#,##0.00" => Some(4),
        "0%" => Some(9),
        "0.00%" => Some(10),
        "0.00E+00" => Some(11),
        "@" => Some(49),
        _ => None,
    }
}

fn seed_builtin_num_fmts(ix: &mut StyleIndex) {
    for (code, id) in [
        ("General", 0),
        ("0", 1),
        ("0.00", 2),
        ("#,##0", 3),
        ("#,##0.00", 4),
        ("0%", 9),
        ("0.00%", 10),
        ("0.00E+00", 11),
        ("@", 49),
    ] {
        ix.numfmt_by_code.entry(code.to_string()).or_insert(id);
        ix.numfmt_code_by_id
            .entry(id)
            .or_insert_with(|| code.to_string());
    }
}

impl XlsxEditor {
    fn ensure_num_fmt(&mut self, code: &str) -> Result<u32> {
        if let Some(id) = builtin_num_fmt_id(code) {
            let ix = self.style_ix_mut()?;
            ix.numfmt_by_code.entry(code.to_string()).or_insert(id);
            ix.numfmt_code_by_id
                .entry(id)
                .or_insert_with(|| code.to_string());
            return Ok(id);
        }

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
        let escaped_code = quick_xml::escape::escape(code);
        let tag = format!(r#"<numFmt numFmtId="{new_id}" formatCode="{escaped_code}"/>"#);
        let root = crate::xml::root(&self.styles_xml)?;
        if root.child("numFmts")?.is_some() {
            self.styles_xml =
                crate::xml::append_child(&self.styles_xml, "numFmts", tag.as_bytes())?.into();
            crate::xml::bump_count(&mut self.styles_xml, "numFmts")?;
        } else {
            let before = root
                .child("fonts")?
                .map(|e| e.full.start)
                .unwrap_or(root.content.start);
            let block = format!(r#"<numFmts count="1">{tag}</numFmts>"#);
            let block = crate::xml::child_fragment(&root, block.as_bytes())?.into_owned();
            self.styles_xml.splice(before..before, block);
        }
        // D) обновляем индекс
        {
            let ix = self.style_ix_mut()?;
            ix.numfmt_by_code.insert(code.to_string(), new_id);
            ix.numfmt_code_by_id.insert(new_id, code.to_string());
            ix.next_custom_numfmt = new_id + 1;
        }

        Ok(new_id)
    }

    fn ensure_font(&mut self, name: &str, size: f32, bold: bool, italic: bool) -> Result<u32> {
        anyhow::ensure!(
            size.is_finite() && size > 0.0,
            "Font size must be finite and positive"
        );
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
        let mut xml = String::from("<font>");
        if bold {
            xml.push_str("<b/>");
        }
        if italic {
            xml.push_str("<i/>");
        }
        xml.push_str(&format!(r#"<sz val="{}"/>"#, (key.size_100 as f32) / 100.0));
        xml.push_str(&format!(
            r#"<name val="{}"/>"#,
            quick_xml::escape::escape(name)
        ));
        xml.push_str("</font>");
        self.styles_xml =
            crate::xml::append_child(&self.styles_xml, "fonts", xml.as_bytes())?.into();
        crate::xml::bump_count(&mut self.styles_xml, "fonts")?;

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
        let xml = format!(
            r#"<fill><patternFill patternType="solid"><fgColor rgb="{key}"/><bgColor indexed="64"/></patternFill></fill>"#
        );
        self.styles_xml =
            crate::xml::append_child(&self.styles_xml, "fills", xml.as_bytes())?.into();
        crate::xml::bump_count(&mut self.styles_xml, "fills")?;

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
        let tag = format!(
            r#"<border><left style="{s}"/><right style="{s}"/><top style="{s}"/><bottom style="{s}"/><diagonal/></border>"#,
            s = style
        );
        self.styles_xml =
            crate::xml::append_child(&self.styles_xml, "borders", tag.as_bytes())?.into();
        crate::xml::bump_count(&mut self.styles_xml, "borders")?;

        // 3) Обновляем индекс ПОСЛЕ вставки, используя pre‑id
        {
            let ix = self.style_ix_mut()?;
            ix.border_by_key.insert(style.to_string(), new_id);
            ix.borders_count = new_id + 1;
        }

        Ok(new_id)
    }

    pub(crate) fn cell_style_id(&self, coord: &str) -> Result<Option<u32>> {
        let (col, number) = crate::data::parse_coord(coord)?;
        let data = crate::xml::require(&self.sheet_xml, "sheetData")?;
        for row in data.children() {
            let row = row?;
            if row
                .attr("r")?
                .as_deref()
                .and_then(|s| s.parse::<u32>().ok())
                != Some(number)
            {
                continue;
            }
            for cell in row.children() {
                let cell = cell?;
                if !cell.is("c", crate::xml::SHEET_NS) {
                    continue;
                }
                if let Some(address) = cell.attr("r")?
                    && crate::data::parse_coord(&address)? == (col, number)
                {
                    return Ok(cell.attr("s")?.map(|s| s.parse()).transpose()?);
                }
            }
        }
        Ok(None)
    }
    fn apply_style_to_cell(&mut self, coord: &str, style: u32) -> Result<()> {
        self.edit_style_target(&Target::Cell(coord.to_owned()), None, Some(style), true)
    }
    fn remove_style_from_cell(&mut self, coord: &str) -> Result<()> {
        self.edit_style_target(&Target::Cell(coord.to_owned()), None, None, false)
    }
    fn remove_style_rect(&mut self, c0: u32, r0: u32, c1: u32, r1: u32) -> Result<()> {
        self.edit_style_target(&Target::Rect { c0, r0, c1, r1 }, None, None, false)
    }
    fn remove_style_row(&mut self, row: u32) -> Result<()> {
        self.edit_style_target(&Target::Row(row), None, None, false)
    }
    fn remove_style_col(&mut self, col: u32) -> Result<()> {
        self.edit_style_target(&Target::Col(col), None, None, false)
    }

    pub fn force_column_formats(&mut self, formats: &[(usize, String)]) -> Result<()> {
        if formats.is_empty() {
            return Ok(());
        }
        let mut patches = HashMap::new();
        let mut defaults = Vec::new();
        for (col, fmt) in formats {
            anyhow::ensure!(
                *col < crate::data::MAX_COLUMNS,
                "XLSX column limit exceeded"
            );
            let patch = StyleParts {
                num_fmt_code: Some(fmt.clone()),
                ..Default::default()
            };
            defaults.push((*col as u32, self.ensure_style_from_parts(&patch)?));
            patches.insert(*col as u32, patch);
        }
        self.set_column_style_ids(&defaults)?;
        let mut cache = HashMap::new();
        self.transform_cells(&Target::All, false, |editor, cell, col, _row| {
            let Some(patch) = patches.get(&(col as u32)) else {
                return Ok(None);
            };
            let old = cell.attr("s")?.map(|v| v.parse::<u32>()).transpose()?;
            let sid = if let Some(sid) = cache.get(&(col, old)) {
                *sid
            } else {
                let parts = editor.read_style_parts(old)?;
                let sid = editor.ensure_style_from_parts(&merge_style_parts(parts, patch))?;
                cache.insert((col, old), sid);
                sid
            };
            if old == Some(sid) {
                return Ok(None);
            }
            let tag = cell.set_attr("s", Some(&sid.to_string()))?;
            Ok(Some(crate::xml::patch_opening(cell, &tag)))
        })
    }
}
