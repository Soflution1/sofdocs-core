use std::io::{Cursor, Read};

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use tracing::warn;
use zip::ZipArchive;

use super::model::*;
use crate::error::{Result, SofDocsError};

/// Parse a .docx file from raw bytes into a Document model.
pub fn parse_docx(bytes: &[u8]) -> Result<Document> {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor)?;

    let styles = parse_styles(&mut archive)?;

    let document_xml = read_entry(&mut archive, "word/document.xml")?;
    let body = parse_document_body(&document_xml, &styles)?;

    Ok(Document {
        metadata: DocumentMetadata::default(),
        body,
        styles,
    })
}

/// Parse a .docx from a file path (convenience for native builds).
#[cfg(feature = "native")]
pub fn parse_docx_file(path: &std::path::Path) -> Result<Document> {
    let bytes = std::fs::read(path)?;
    parse_docx(&bytes)
}

fn read_entry(archive: &mut ZipArchive<Cursor<&[u8]>>, name: &str) -> Result<String> {
    let mut file = archive
        .by_name(name)
        .map_err(|_| SofDocsError::MissingEntry(name.to_string()))?;
    let mut contents = String::new();
    file.read_to_string(&mut contents)?;
    Ok(contents)
}

/// Extract local name from a namespace-prefixed QName bytes, returning an owned String.
fn tag_local_name(name_bytes: &[u8]) -> String {
    let s = std::str::from_utf8(name_bytes).unwrap_or("");
    s.rsplit(':').next().unwrap_or(s).to_string()
}

fn parse_styles(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<Vec<StyleDefinition>> {
    let xml = match read_entry(archive, "word/styles.xml") {
        Ok(xml) => xml,
        Err(_) => return Ok(Vec::new()),
    };

    let mut reader = Reader::from_str(&xml);
    let mut styles = Vec::new();
    let mut current_style: Option<StyleDefinition> = None;
    let mut in_rpr = false;
    let mut in_ppr = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let local = tag_local_name(e.name().as_ref());
                match local.as_str() {
                    "style" => {
                        let mut sd = StyleDefinition::default();
                        for attr in e.attributes().flatten() {
                            let key = tag_local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "styleId" => sd.style_id = val,
                                "type" => {
                                    sd.style_type = match val.as_str() {
                                        "paragraph" => StyleType::Paragraph,
                                        "character" => StyleType::Character,
                                        "table" => StyleType::Table,
                                        "numbering" => StyleType::Numbering,
                                        _ => StyleType::Paragraph,
                                    };
                                }
                                _ => {}
                            }
                        }
                        current_style = Some(sd);
                    }
                    "name" if current_style.is_some() => {
                        if let Some(ref mut sd) = current_style {
                            for attr in e.attributes().flatten() {
                                if tag_local_name(attr.key.as_ref()) == "val" {
                                    sd.name =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                    }
                    "basedOn" if current_style.is_some() => {
                        if let Some(ref mut sd) = current_style {
                            for attr in e.attributes().flatten() {
                                if tag_local_name(attr.key.as_ref()) == "val" {
                                    sd.based_on =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                    }
                    "rPr" if current_style.is_some() => in_rpr = true,
                    "pPr" if current_style.is_some() => in_ppr = true,
                    _ if in_rpr && current_style.is_some() => {
                        if let Some(ref mut sd) = current_style {
                            apply_run_property(&mut sd.run_style, &local, e);
                        }
                    }
                    _ if in_ppr && current_style.is_some() => {
                        if let Some(ref mut sd) = current_style {
                            apply_paragraph_property(&mut sd.paragraph_properties, &local, e);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let local = tag_local_name(e.name().as_ref());
                match local.as_str() {
                    "style" => {
                        if let Some(sd) = current_style.take() {
                            styles.push(sd);
                        }
                    }
                    "rPr" => in_rpr = false,
                    "pPr" => in_ppr = false,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                warn!("Error parsing styles.xml: {e}");
                break;
            }
            _ => {}
        }
    }

    Ok(styles)
}

fn parse_document_body(xml: &str, styles: &[StyleDefinition]) -> Result<DocumentBody> {
    let mut reader = Reader::from_str(xml);
    let mut body = DocumentBody::default();

    let mut in_body = false;
    let mut current_paragraph: Option<Paragraph> = None;
    let mut current_run: Option<Run> = None;
    let mut in_rpr = false;
    let mut in_ppr = false;
    let mut in_text = false;

    let mut current_table: Option<Table> = None;
    let mut current_row: Option<TableRow> = None;
    let mut current_cell: Option<TableCell> = None;
    let mut in_table = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let local = tag_local_name(e.name().as_ref());
                match local.as_str() {
                    "body" => in_body = true,
                    "tbl" if in_body => {
                        in_table = true;
                        current_table = Some(Table::default());
                    }
                    "tr" if in_table => {
                        current_row = Some(TableRow::default());
                    }
                    "tc" if in_table => {
                        current_cell = Some(TableCell::default());
                    }
                    "p" if in_body => {
                        current_paragraph = Some(Paragraph::default());
                        in_ppr = false;
                    }
                    "pPr" if current_paragraph.is_some() => {
                        in_ppr = true;
                    }
                    "pStyle" if in_ppr && current_paragraph.is_some() => {
                        if let Some(ref mut para) = current_paragraph {
                            for attr in e.attributes().flatten() {
                                if tag_local_name(attr.key.as_ref()) == "val" {
                                    let style_id =
                                        String::from_utf8_lossy(&attr.value).to_string();
                                    para.properties.heading_level =
                                        resolve_heading_level(&style_id, styles);
                                    para.properties.style_id = Some(style_id);
                                }
                            }
                        }
                    }
                    "jc" if in_ppr && current_paragraph.is_some() => {
                        if let Some(ref mut para) = current_paragraph {
                            apply_paragraph_property(&mut para.properties, &local, e);
                        }
                    }
                    "numPr" if in_ppr => {}
                    "numId" if in_ppr && current_paragraph.is_some() => {
                        if let Some(ref mut para) = current_paragraph {
                            let num =
                                para.properties.numbering.get_or_insert_with(Default::default);
                            for attr in e.attributes().flatten() {
                                if tag_local_name(attr.key.as_ref()) == "val" {
                                    if let Ok(v) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        num.num_id = v;
                                    }
                                }
                            }
                        }
                    }
                    "ilvl" if in_ppr && current_paragraph.is_some() => {
                        if let Some(ref mut para) = current_paragraph {
                            let num =
                                para.properties.numbering.get_or_insert_with(Default::default);
                            for attr in e.attributes().flatten() {
                                if tag_local_name(attr.key.as_ref()) == "val" {
                                    if let Ok(v) =
                                        String::from_utf8_lossy(&attr.value).parse::<u32>()
                                    {
                                        num.level = v;
                                    }
                                }
                            }
                        }
                    }
                    "r" if current_paragraph.is_some() => {
                        current_run = Some(Run::default());
                        in_rpr = false;
                    }
                    "rPr" if current_run.is_some() => {
                        in_rpr = true;
                    }
                    "t" if current_run.is_some() => {
                        in_text = true;
                    }
                    _ if in_rpr && current_run.is_some() => {
                        if let Some(ref mut run) = current_run {
                            apply_run_property(&mut run.style, &local, e);
                        }
                    }
                    _ if in_ppr && current_paragraph.is_some() => {
                        if let Some(ref mut para) = current_paragraph {
                            apply_paragraph_property(&mut para.properties, &local, e);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if in_text {
                    if let Some(ref mut run) = current_run {
                        let text = e.unescape().unwrap_or_default();
                        run.text.push_str(&text);
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local = tag_local_name(e.name().as_ref());
                match local.as_str() {
                    "body" => in_body = false,
                    "t" => in_text = false,
                    "r" => {
                        in_text = false;
                        if let (Some(run), Some(ref mut para)) =
                            (current_run.take(), &mut current_paragraph)
                        {
                            if !run.text.is_empty() {
                                para.runs.push(run);
                            }
                        }
                        in_rpr = false;
                    }
                    "rPr" => in_rpr = false,
                    "pPr" => in_ppr = false,
                    "p" => {
                        if let Some(para) = current_paragraph.take() {
                            if in_table {
                                if let Some(ref mut cell) = current_cell {
                                    cell.paragraphs.push(para);
                                }
                            } else {
                                body.paragraphs.push(para);
                            }
                        }
                    }
                    "tc" => {
                        if let (Some(cell), Some(ref mut row)) =
                            (current_cell.take(), &mut current_row)
                        {
                            row.cells.push(cell);
                        }
                    }
                    "tr" => {
                        if let (Some(row), Some(ref mut table)) =
                            (current_row.take(), &mut current_table)
                        {
                            table.rows.push(row);
                        }
                    }
                    "tbl" => {
                        if let Some(table) = current_table.take() {
                            body.tables.push(table);
                        }
                        in_table = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                warn!("Error parsing document.xml: {e}");
                break;
            }
            _ => {}
        }
    }

    Ok(body)
}

fn apply_run_property(style: &mut RunStyle, local: &str, e: &quick_xml::events::BytesStart) {
    match local {
        "b" => style.bold = !has_val_false(e),
        "i" => style.italic = !has_val_false(e),
        "u" => style.underline = !has_val_false(e),
        "strike" => style.strikethrough = !has_val_false(e),
        "rFonts" => {
            for attr in e.attributes().flatten() {
                let key = tag_local_name(attr.key.as_ref());
                if key == "ascii" || key == "hAnsi" || key == "cs" {
                    style.font_family =
                        Some(String::from_utf8_lossy(&attr.value).to_string());
                    break;
                }
            }
        }
        "sz" => {
            for attr in e.attributes().flatten() {
                if tag_local_name(attr.key.as_ref()) == "val" {
                    if let Ok(half_pts) = String::from_utf8_lossy(&attr.value).parse::<f32>() {
                        style.font_size_pt = Some(half_pts / 2.0);
                    }
                }
            }
        }
        "color" => {
            for attr in e.attributes().flatten() {
                if tag_local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    if val != "auto" {
                        style.color = Some(val);
                    }
                }
            }
        }
        "highlight" => {
            for attr in e.attributes().flatten() {
                if tag_local_name(attr.key.as_ref()) == "val" {
                    style.highlight = Some(String::from_utf8_lossy(&attr.value).to_string());
                }
            }
        }
        "vertAlign" => {
            for attr in e.attributes().flatten() {
                if tag_local_name(attr.key.as_ref()) == "val" {
                    let val = String::from_utf8_lossy(&attr.value);
                    style.vertical_align = match val.as_ref() {
                        "superscript" => Some(VerticalAlign::Superscript),
                        "subscript" => Some(VerticalAlign::Subscript),
                        _ => None,
                    };
                }
            }
        }
        _ => {}
    }
}

fn apply_paragraph_property(
    props: &mut ParagraphProperties,
    local: &str,
    e: &quick_xml::events::BytesStart,
) {
    match local {
        "jc" => {
            for attr in e.attributes().flatten() {
                if tag_local_name(attr.key.as_ref()) == "val" {
                    props.alignment = Some(String::from_utf8_lossy(&attr.value).to_string());
                }
            }
        }
        "ind" => {
            for attr in e.attributes().flatten() {
                let key = tag_local_name(attr.key.as_ref());
                let val = String::from_utf8_lossy(&attr.value);
                match key.as_str() {
                    "left" | "start" => {
                        props.indent_left_twips = val.parse().ok();
                    }
                    "right" | "end" => {
                        props.indent_right_twips = val.parse().ok();
                    }
                    "firstLine" => {
                        props.indent_first_line_twips = val.parse().ok();
                    }
                    _ => {}
                }
            }
        }
        "spacing" => {
            for attr in e.attributes().flatten() {
                let key = tag_local_name(attr.key.as_ref());
                let val = String::from_utf8_lossy(&attr.value);
                match key.as_str() {
                    "before" => props.spacing_before_twips = val.parse().ok(),
                    "after" => props.spacing_after_twips = val.parse().ok(),
                    "line" => props.line_spacing_twips = val.parse().ok(),
                    _ => {}
                }
            }
        }
        _ => {}
    }
}

fn resolve_heading_level(style_id: &str, styles: &[StyleDefinition]) -> u8 {
    let lower = style_id.to_lowercase();

    if lower.starts_with("heading") || lower.starts_with("titre") {
        if let Some(digit) = lower.chars().last().and_then(|c| c.to_digit(10)) {
            return digit.min(6) as u8;
        }
    }

    for sd in styles {
        if sd.style_id == style_id {
            if let Some(ref name) = sd.name {
                let name_lower = name.to_lowercase();
                if name_lower.starts_with("heading") || name_lower.starts_with("titre") {
                    if let Some(digit) = name_lower.chars().last().and_then(|c| c.to_digit(10)) {
                        return digit.min(6) as u8;
                    }
                }
            }
        }
    }

    0
}

fn has_val_false(e: &quick_xml::events::BytesStart) -> bool {
    for attr in e.attributes().flatten() {
        if tag_local_name(attr.key.as_ref()) == "val" {
            let val = String::from_utf8_lossy(&attr.value);
            return val == "false" || val == "0";
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_tag_local_name() {
        assert_eq!(tag_local_name(b"w:body"), "body");
        assert_eq!(tag_local_name(b"body"), "body");
        assert_eq!(tag_local_name(b"w:p"), "p");
        assert_eq!(tag_local_name(b"mc:AlternateContent"), "AlternateContent");
    }

    #[test]
    fn test_resolve_heading_level() {
        assert_eq!(resolve_heading_level("Heading1", &[]), 1);
        assert_eq!(resolve_heading_level("Heading3", &[]), 3);
        assert_eq!(resolve_heading_level("Titre2", &[]), 2);
        assert_eq!(resolve_heading_level("Normal", &[]), 0);
    }
}
