use anyhow::{Context, Result};
use memchr::memmem;

pub fn col_letter(mut n: u32) -> String {
    let mut s = String::new();
    loop {
        s.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    s
}

pub fn col_index(s: &str) -> usize {
    crate::data::parse_column(s).expect("Invalid Excel column")
}
pub fn split_coord(coord: &str) -> (u32, u32) {
    let (col, row) = crate::data::parse_coord(coord).expect("Invalid Excel coordinate");
    (col as u32, row)
}

#[inline]
pub fn find_bytes_from(hay: &[u8], needle: &[u8], start: usize) -> Option<usize> {
    if start >= hay.len() {
        return None;
    }
    // поищем в срезе с нужного оффсета и поправим индекс
    memmem::find(&hay[start..], needle).map(|i| i + start)
}

pub fn bump_count(xml: &mut Vec<u8>, tag: &[u8], attr: &[u8]) -> Result<()> {
    let name = std::str::from_utf8(tag)?
        .trim_start_matches('<')
        .trim_end_matches('>');
    let attribute = std::str::from_utf8(attr)?.trim_end_matches(['=', '\"']);
    let element = crate::xml::require(xml, name)?;
    let value: u32 = element
        .attr(attribute)?
        .context("XML count attribute missing")?
        .parse()?;
    let next = value.checked_add(1).context("XML count overflow")?;
    let opening = element.set_attr(attribute, Some(&next.to_string()))?;
    let range = element.open;
    xml.splice(range, opening);
    Ok(())
}
