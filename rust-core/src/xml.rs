//! One XML boundary for OOXML: validated element spans and decoded attributes.
//! Edits keep unknown XML verbatim; no DOM or quote/attribute-order assumptions.
use anyhow::{Context, Result, ensure};
use quick_xml::{
    Reader, Writer, XmlVersion,
    events::{BytesStart, Event},
};
use std::{borrow::Cow, collections::HashMap, io::Write, ops::Range, sync::Arc};

pub(crate) const SHEET_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
pub(crate) const DOC_REL_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const CT_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

#[derive(Clone, Debug)]
pub(crate) struct Element<'a> {
    source: &'a [u8],
    pub name: &'a str,
    pub open: Range<usize>,
    pub content: Range<usize>,
    pub full: Range<usize>,
    pub empty: bool,
    start: BytesStart<'a>,
    scope: Arc<HashMap<String, String>>,
}
impl<'a> Element<'a> {
    pub fn local_name(&self) -> &str {
        self.name.rsplit(':').next().unwrap_or(self.name)
    }
    pub fn namespace(&self) -> Option<&str> {
        let prefix = self
            .name
            .split_once(':')
            .map(|(prefix, _)| prefix)
            .unwrap_or("");
        self.scope.get(prefix).map(String::as_str)
    }
    pub fn is(&self, name: &str, namespace: &str) -> bool {
        self.local_name() == name && self.namespace() == Some(namespace)
    }
    pub fn same_namespace(&self, other: &Element<'_>) -> bool {
        self.namespace() == other.namespace()
    }
    pub fn attr_ns(&self, name: &str, namespace: &str) -> Result<Option<Cow<'_, str>>> {
        for attribute in self.start.attributes() {
            let attribute = attribute?;
            if attribute.key.local_name().into_inner() == name
                && attribute
                    .key
                    .prefix()
                    .and_then(|prefix| self.scope.get(prefix.into_inner()))
                    .map(String::as_str)
                    == Some(namespace)
            {
                return Ok(Some(attribute.normalized_value(XmlVersion::Implicit1_0)?));
            }
        }
        Ok(None)
    }
    pub fn attr(&self, name: &str) -> Result<Option<Cow<'_, str>>> {
        for attribute in self.start.attributes() {
            let attribute = attribute?;
            if attribute.key.into_inner() == name {
                return Ok(Some(attribute.normalized_value(XmlVersion::Implicit1_0)?));
            }
        }
        Ok(None)
    }
    pub fn children(&self) -> Elements<'a> {
        Elements::range(self.source, self.content.clone(), self.scope.clone())
    }
    pub fn child(&self, name: &str) -> Result<Option<Element<'a>>> {
        for child in self.children() {
            let child = child?;
            if child.local_name() == name && self.same_namespace(&child) {
                return Ok(Some(child));
            }
        }
        Ok(None)
    }
    pub fn require_child(&self, name: &str) -> Result<Element<'a>> {
        self.child(name)?
            .with_context(|| format!("<{name}> not found in <{}>", self.name))
    }
    pub fn segment(&self, range: Range<usize>) -> &'a [u8] {
        &self.source[range]
    }
    pub fn bytes(&self) -> &'a [u8] {
        &self.source[self.full.clone()]
    }
    pub fn set_attr(&self, name: &str, value: Option<&str>) -> Result<Vec<u8>> {
        let mut start = BytesStart::new(self.name);
        for attribute in self.start.attributes() {
            let attribute = attribute?;
            if attribute.key.into_inner() != name {
                let value = attribute.normalized_value(XmlVersion::Implicit1_0)?;
                start.push_attribute((attribute.key.into_inner(), value.as_ref()));
            }
        }
        if let Some(value) = value {
            start.push_attribute((name, value));
        }
        let mut writer = Writer::new(Vec::new());
        writer.write_event(if self.empty {
            Event::Empty(start)
        } else {
            Event::Start(start)
        })?;
        Ok(writer.into_inner())
    }
    pub fn with_content(&self, content: &[u8]) -> Result<Vec<u8>> {
        let mut out = Vec::with_capacity(self.open.len() + content.len() + self.name.len() + 3);
        let mut writer = Writer::new(&mut out);
        writer.write_event(Event::Start(self.start.clone()))?;
        writer.get_mut().write_all(content)?;
        writer.write_event(Event::End(self.start.to_end()))?;
        Ok(out)
    }
}

pub(crate) struct Elements<'a> {
    source: &'a [u8],
    reader: Reader<&'a [u8]>,
    base: usize,
    finished: bool,
    scope: Arc<HashMap<String, String>>,
    stop_before: Option<&'static str>,
    closing_name: Option<String>,
}
impl<'a> Elements<'a> {
    pub fn range(
        source: &'a [u8],
        range: Range<usize>,
        scope: Arc<HashMap<String, String>>,
    ) -> Self {
        // quick-xml 0.42 strips a BOM without counting it in buffer_position.
        // Strip it explicitly so every span stays relative to the original bytes.
        let bom = usize::from(source[range.clone()].starts_with(b"\xef\xbb\xbf")) * 3;
        let base = range.start + bom;
        Self {
            source,
            reader: Reader::from_reader(&source[base..range.end]),
            base,
            finished: false,
            scope,
            stop_before: None,
            closing_name: None,
        }
    }
    fn read_next(&mut self) -> Result<Option<Element<'a>>> {
        loop {
            match self.reader.read_event()? {
                Event::Start(start) => {
                    if self.stop_before == Some(start.local_name().as_ref()) {
                        return Ok(None);
                    }
                    let open_end = self.reader.buffer_position() as usize + self.base;
                    let begin = open_end - start.as_ref().len() - 2;
                    let name = std::str::from_utf8(
                        &self.source[begin + 1..begin + 1 + start.name().as_ref().len()],
                    )?;
                    let scope = namespace_scope(&start, &self.scope)?;
                    let span = self.reader.read_to_end(start.name())?;
                    let end = self.reader.buffer_position() as usize + self.base;
                    return Ok(Some(Element {
                        source: self.source,
                        name,
                        open: begin..open_end,
                        content: open_end..span.end as usize + self.base,
                        full: begin..end,
                        empty: false,
                        start,
                        scope,
                    }));
                }
                Event::Empty(start) => {
                    let end = self.reader.buffer_position() as usize + self.base;
                    let begin = end - start.as_ref().len() - 3;
                    let name = std::str::from_utf8(
                        &self.source[begin + 1..begin + 1 + start.name().as_ref().len()],
                    )?;
                    let scope = namespace_scope(&start, &self.scope)?;
                    return Ok(Some(Element {
                        source: self.source,
                        name,
                        open: begin..end,
                        content: end..end,
                        full: begin..end,
                        empty: true,
                        start,
                        scope,
                    }));
                }
                Event::Eof => return Ok(None),
                Event::End(end) => {
                    ensure!(
                        self.closing_name.as_deref() == Some(end.name().as_ref()),
                        "Unexpected XML closing element"
                    );
                    return Ok(None);
                }
                _ => {}
            }
        }
    }
}
impl<'a> Iterator for Elements<'a> {
    type Item = Result<Element<'a>>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.finished {
            return None;
        }
        match self.read_next() {
            Ok(Some(element)) => Some(Ok(element)),
            Ok(None) => {
                self.finished = true;
                None
            }
            Err(error) => {
                self.finished = true;
                Some(Err(error))
            }
        }
    }
}

pub(crate) fn root(xml: &[u8]) -> Result<Element<'_>> {
    root_scoped(xml, Arc::new(HashMap::new()))
}
fn root_scoped(xml: &[u8], scope: Arc<HashMap<String, String>>) -> Result<Element<'_>> {
    let mut elements = Elements::range(xml, 0..xml.len(), scope);
    let root = elements.next().context("XML has no root element")??;
    ensure!(
        elements.next().transpose()?.is_none(),
        "XML has multiple root elements"
    );
    Ok(root)
}
pub(crate) fn root_in_context<'a>(xml: &'a [u8], parent: &Element<'_>) -> Result<Element<'a>> {
    root_scoped(xml, parent.scope.clone())
}
fn namespace_scope(
    start: &BytesStart<'_>,
    parent: &Arc<HashMap<String, String>>,
) -> Result<Arc<HashMap<String, String>>> {
    let mut scope = parent.clone();
    for attribute in start.attributes() {
        let attribute = attribute?;
        let key = attribute.key.into_inner();
        if key == "xmlns" || key.starts_with("xmlns:") {
            let prefix = key.strip_prefix("xmlns:").unwrap_or("");
            Arc::make_mut(&mut scope).insert(
                prefix.to_owned(),
                attribute
                    .normalized_value(XmlVersion::Implicit1_0)?
                    .into_owned(),
            );
        }
    }
    Ok(scope)
}
pub(crate) fn child_fragment<'a>(parent: &Element<'_>, bytes: &'a [u8]) -> Result<Cow<'a, [u8]>> {
    let namespace = parent.namespace();
    if namespace.is_none() || parent.scope.get("").map(String::as_str) == namespace {
        return Ok(Cow::Borrowed(bytes));
    }
    let child = root(bytes)?;
    let tag = child.set_attr("xmlns", namespace)?;
    Ok(Cow::Owned(patch_opening(&child, &tag)))
}
pub(crate) fn optional<'a>(source: &'a [u8], name: &str) -> Result<Option<Element<'a>>> {
    let context = root_context(source)?;
    if context.name.rsplit(':').next() == Some(name) {
        return Ok(Some(root(source)?));
    }
    find_child(source, name, &context, false)
}
pub(crate) fn header_child<'a>(source: &'a [u8], name: &str) -> Result<Option<Element<'a>>> {
    let context = root_context(source)?;
    find_child(source, name, &context, true)
}
fn find_child<'a>(
    source: &'a [u8],
    name: &str,
    context: &RootContext,
    header_only: bool,
) -> Result<Option<Element<'a>>> {
    if context.empty {
        return Ok(None);
    }
    let mut elements = Elements::range(
        source,
        context.content_start..source.len(),
        context.scope.clone(),
    );
    elements.closing_name = Some(context.name.clone());
    if header_only {
        elements.stop_before = Some("sheetData");
    }
    for element in elements {
        let element = element?;
        if element.local_name() == name
            && element.namespace() == context.scope.get(&context.prefix).map(String::as_str)
        {
            return Ok(Some(element));
        }
    }
    Ok(None)
}
pub(crate) fn require<'a>(source: &'a [u8], name: &str) -> Result<Element<'a>> {
    optional(source, name)?.with_context(|| format!("<{name}> not found"))
}
pub(crate) fn append_child(xml: &[u8], parent: &str, child: &[u8]) -> Result<Vec<u8>> {
    let element = require(xml, parent)?;
    let mut content = xml[element.content.clone()].to_vec();
    content.extend_from_slice(&child_fragment(&element, child)?);
    let replacement = element.with_content(&content)?;
    let mut result = xml.to_vec();
    result.splice(element.full, replacement);
    Ok(result)
}
pub(crate) fn ensure_default_namespace(xml: Vec<u8>, namespace: &str) -> Result<Vec<u8>> {
    let root = root(&xml)?;
    if root.namespace() == Some(namespace) {
        return Ok(xml);
    }
    ensure!(
        root.namespace().is_none() && !root.name.contains(':'),
        "Unexpected OOXML namespace for {}",
        root.name
    );
    let tag = root.set_attr("xmlns", Some(namespace))?;
    let range = root.open.clone();
    let mut out = xml;
    out.splice(range, tag);
    Ok(out)
}
pub(crate) fn replace_element(xml: &[u8], element: &Element<'_>, replacement: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(xml.len() - element.full.len() + replacement.len());
    out.extend_from_slice(&xml[..element.full.start]);
    out.extend_from_slice(replacement);
    out.extend_from_slice(&xml[element.full.end..]);
    out
}
pub(crate) fn bump_count(xml: &mut Vec<u8>, name: &str) -> Result<()> {
    let element = require(xml, name)?;
    let mut count = 0;
    for child in element.children() {
        let child = child?;
        if child.same_namespace(&element) {
            count += 1;
        }
    }
    let tag = element.set_attr("count", Some(&count.to_string()))?;
    let range = element.open;
    xml.splice(range, tag);
    Ok(())
}
pub(crate) fn escape_attribute(value: &str) -> Cow<'_, str> {
    quick_xml::escape::escape(value)
}

pub(crate) fn patch_opening(element: &Element<'_>, tag: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(element.full.len() + tag.len());
    out.extend_from_slice(tag);
    out.extend_from_slice(&element.source[element.open.end..element.full.end]);
    out
}
pub(crate) fn text(element: &Element<'_>) -> Result<String> {
    let mut reader = Reader::from_reader(&element.source[element.content.clone()]);
    let mut result = String::new();
    loop {
        match reader.read_event()? {
            Event::Text(value) => result.push_str(&quick_xml::escape::unescape(value.as_ref())?),
            Event::CData(value) => result.push_str(value.as_ref()),
            Event::GeneralRef(value) => {
                let encoded = format!("&{};", value.as_ref());
                result.push_str(&quick_xml::escape::unescape(&encoded)?);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(result)
}

/// Build a part replacement directly in its final buffer, avoiding a second
/// worksheet-sized content/replacement allocation.
pub(crate) fn begin_content(source: &[u8], element: &Element<'_>) -> Result<Vec<u8>> {
    let mut out = Vec::with_capacity(source.len());
    if element.empty {
        out.extend_from_slice(&source[..element.full.start]);
        Writer::new(&mut out).write_event(Event::Start(element.start.clone()))?;
    } else {
        out.extend_from_slice(&source[..element.open.end]);
    }
    Ok(out)
}
pub(crate) fn end_content(source: &[u8], element: &Element<'_>, out: &mut Vec<u8>) -> Result<()> {
    if element.empty {
        Writer::new(&mut *out).write_event(Event::End(element.start.to_end()))?;
        out.extend_from_slice(&source[element.full.end..]);
    } else {
        out.extend_from_slice(&source[element.content.end..]);
    }
    Ok(())
}

/// Namespace context from the root opening tag only. This is constant work for
/// cached tail appends; it deliberately does not rescan the worksheet body.
#[derive(Clone)]
pub(crate) struct RootContext {
    scope: Arc<HashMap<String, String>>,
    prefix: String,
    name: String,
    pub content_start: usize,
    empty: bool,
}
pub(crate) fn root_context(source: &[u8]) -> Result<RootContext> {
    let bom = usize::from(source.starts_with(b"\xef\xbb\xbf")) * 3;
    let mut reader = Reader::from_reader(&source[bom..]);
    loop {
        let (start, empty) = match reader.read_event()? {
            Event::Start(start) => (start, false),
            Event::Empty(start) => (start, true),
            Event::Eof => anyhow::bail!("XML has no root element"),
            _ => continue,
        };
        let scope = namespace_scope(&start, &Arc::new(HashMap::new()))?;
        let name = start.name().as_ref().to_owned();
        let prefix = start
            .name()
            .prefix()
            .map(|p| p.into_inner().to_owned())
            .unwrap_or_default();
        return Ok(RootContext {
            scope,
            prefix,
            name,
            content_start: reader.buffer_position() as usize + bom,
            empty,
        });
    }
}
pub(crate) fn root_in_root_context<'a>(
    bytes: &'a [u8],
    context: &RootContext,
) -> Result<Element<'a>> {
    root_scoped(bytes, context.scope.clone())
}

impl RootContext {
    pub fn fragment<'a>(&self, bytes: &'a [u8]) -> Result<Cow<'a, [u8]>> {
        let target = self.scope.get(&self.prefix);
        if target.is_none() || self.scope.get("") == target {
            return Ok(Cow::Borrowed(bytes));
        }
        let child = root(bytes)?;
        let tag = child.set_attr("xmlns", target.map(String::as_str))?;
        Ok(Cow::Owned(patch_opening(&child, &tag)))
    }
}

impl Element<'_> {
    pub(crate) fn context(&self) -> RootContext {
        RootContext {
            scope: self.scope.clone(),
            prefix: self
                .name
                .split_once(':')
                .map(|(p, _)| p)
                .unwrap_or("")
                .to_owned(),
            name: self.name.to_owned(),
            content_start: self.open.end,
            empty: self.empty,
        }
    }
    pub(crate) fn attributes(&self) -> Result<Vec<(String, String)>> {
        self.start
            .attributes()
            .map(|attribute| {
                let attribute = attribute?;
                Ok((
                    attribute.key.into_inner().to_owned(),
                    attribute
                        .normalized_value(XmlVersion::Implicit1_0)?
                        .into_owned(),
                ))
            })
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn spans_handle_quotes_prefixes_cdata_and_empty_elements() -> Result<()> {
        let xml = br#"<?xml version='1.0'?><x:root xmlns:x='urn:test'><x:item
 data = 'a > &amp; b'><![CDATA[<item>fake</item>]]><x:empty /></x:item><!-- > --></x:root>"#;
        let root = root(xml)?;
        let item = root.require_child("item")?;
        assert_eq!(item.attr("data")?.as_deref(), Some("a > & b"));
        assert!(item.require_child("empty")?.empty);
        assert!(std::str::from_utf8(item.bytes())?.starts_with("<x:item"));
        assert!(std::str::from_utf8(item.bytes())?.ends_with("</x:item>"));
        Ok(())
    }
    #[test]
    fn bom_offsets_refer_to_original_bytes() -> Result<()> {
        for bytes in [
            b"\xef\xbb\xbf<root><child/></root>".as_slice(),
            b"\xef\xbb\xbf<?xml version='1.0'?><root><child/></root>".as_slice(),
        ] {
            let element = root(bytes)?;
            assert_eq!(element.bytes(), b"<root><child/></root>");
            assert_eq!(element.require_child("child")?.bytes(), b"<child/>");
            assert_eq!(require(bytes, "child")?.bytes(), b"<child/>");
            assert_eq!(
                &bytes[root_context(bytes)?.content_start..][..8],
                b"<child/>"
            );
        }
        Ok(())
    }
    #[test]
    fn malformed_input_returns_an_error() {
        assert!(root(b"<root><row></root>").is_err());
    }
}
