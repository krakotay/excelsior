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
    s.bytes().fold(0, |acc, b| {
        acc * 26 + (b.to_ascii_uppercase() - b'A' + 1) as usize
    }) - 1
}

pub fn split_coord(coord: &str) -> (u32, u32) {
    let p = coord.find(|c: char| c.is_ascii_digit()).unwrap();
    (
        col_index(&coord[..p]) as u32,
        coord[p..].parse::<u32>().unwrap(),
    )
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
    if let Some(pos) = memmem::rfind(xml, tag) {
        if let Some(a) = find_bytes_from(xml, attr, pos) {
            let start = a + attr.len();
            let end = find_bytes_from(xml, b"\"", start).context("closing quote not found")?;
            let mut num: u32 = std::str::from_utf8(&xml[start..end])?.parse()?;
            num += 1;
            xml.splice(start..end, num.to_string().bytes());
            return Ok(());
        }
    }
    Err(anyhow::anyhow!("attribute count not found"))
}
