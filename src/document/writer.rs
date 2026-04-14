use std::io::{Cursor, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use super::model::*;
use crate::error::Result;

/// Write a Document model back to .docx bytes (ZIP with OOXML).
pub fn write_docx(doc: &Document) -> Result<Vec<u8>> {
    let buf = Vec::new();
    let cursor = Cursor::new(buf);
    let mut zip = ZipWriter::new(cursor);
    let opts = SimpleFileOptions::default();

    zip.start_file("[Content_Types].xml", opts)?;
    zip.write_all(content_types_xml().as_bytes())?;

    zip.start_file("_rels/.rels", opts)?;
    zip.write_all(rels_xml().as_bytes())?;

    zip.start_file("word/_rels/document.xml.rels", opts)?;
    zip.write_all(document_rels_xml().as_bytes())?;

    zip.start_file("word/styles.xml", opts)?;
    zip.write_all(styles_xml(&doc.styles).as_bytes())?;

    zip.start_file("word/document.xml", opts)?;
    zip.write_all(document_xml(&doc.body).as_bytes())?;

    let cursor = zip.finish()?;
    Ok(cursor.into_inner())
}

fn content_types_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#.to_string()
}

fn rels_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#.to_string()
}

fn document_rels_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#.to_string()
}

fn document_xml(body: &DocumentBody) -> String {
    let mut xml = String::with_capacity(4096);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>"#);

    for para in &body.paragraphs {
        write_paragraph(&mut xml, para);
    }

    for table in &body.tables {
        write_table(&mut xml, table);
    }

    xml.push_str("\n  </w:body>\n</w:document>");
    xml
}

fn write_paragraph(xml: &mut String, para: &Paragraph) {
    xml.push_str("\n    <w:p>");

    if has_paragraph_properties(&para.properties) {
        xml.push_str("<w:pPr>");
        if let Some(ref style_id) = para.properties.style_id {
            xml.push_str(&format!("<w:pStyle w:val=\"{}\"/>", xml_escape(style_id)));
        }
        if let Some(ref alignment) = para.properties.alignment {
            xml.push_str(&format!("<w:jc w:val=\"{}\"/>", xml_escape(alignment)));
        }
        if para.properties.indent_left_twips.is_some()
            || para.properties.indent_right_twips.is_some()
            || para.properties.indent_first_line_twips.is_some()
        {
            xml.push_str("<w:ind");
            if let Some(v) = para.properties.indent_left_twips {
                xml.push_str(&format!(" w:left=\"{v}\""));
            }
            if let Some(v) = para.properties.indent_right_twips {
                xml.push_str(&format!(" w:right=\"{v}\""));
            }
            if let Some(v) = para.properties.indent_first_line_twips {
                xml.push_str(&format!(" w:firstLine=\"{v}\""));
            }
            xml.push_str("/>");
        }
        if para.properties.spacing_before_twips.is_some()
            || para.properties.spacing_after_twips.is_some()
            || para.properties.line_spacing_twips.is_some()
        {
            xml.push_str("<w:spacing");
            if let Some(v) = para.properties.spacing_before_twips {
                xml.push_str(&format!(" w:before=\"{v}\""));
            }
            if let Some(v) = para.properties.spacing_after_twips {
                xml.push_str(&format!(" w:after=\"{v}\""));
            }
            if let Some(v) = para.properties.line_spacing_twips {
                xml.push_str(&format!(" w:line=\"{v}\""));
            }
            xml.push_str("/>");
        }
        if let Some(ref num) = para.properties.numbering {
            xml.push_str(&format!(
                "<w:numPr><w:ilvl w:val=\"{}\"/><w:numId w:val=\"{}\"/></w:numPr>",
                num.level, num.num_id
            ));
        }
        xml.push_str("</w:pPr>");
    }

    for run in &para.runs {
        write_run(xml, run);
    }

    xml.push_str("</w:p>");
}

fn write_run(xml: &mut String, run: &Run) {
    xml.push_str("<w:r>");

    if has_run_style(&run.style) {
        xml.push_str("<w:rPr>");
        if run.style.bold {
            xml.push_str("<w:b/>");
        }
        if run.style.italic {
            xml.push_str("<w:i/>");
        }
        if run.style.underline {
            xml.push_str("<w:u w:val=\"single\"/>");
        }
        if run.style.strikethrough {
            xml.push_str("<w:strike/>");
        }
        if let Some(ref font) = run.style.font_family {
            xml.push_str(&format!(
                "<w:rFonts w:ascii=\"{}\" w:hAnsi=\"{}\"/>",
                xml_escape(font),
                xml_escape(font)
            ));
        }
        if let Some(size) = run.style.font_size_pt {
            let half_pts = (size * 2.0) as u32;
            xml.push_str(&format!("<w:sz w:val=\"{half_pts}\"/>"));
        }
        if let Some(ref color) = run.style.color {
            xml.push_str(&format!("<w:color w:val=\"{}\"/>", xml_escape(color)));
        }
        if let Some(ref highlight) = run.style.highlight {
            xml.push_str(&format!(
                "<w:highlight w:val=\"{}\"/>",
                xml_escape(highlight)
            ));
        }
        if let Some(ref va) = run.style.vertical_align {
            let val = match va {
                VerticalAlign::Superscript => "superscript",
                VerticalAlign::Subscript => "subscript",
            };
            xml.push_str(&format!("<w:vertAlign w:val=\"{val}\"/>"));
        }
        xml.push_str("</w:rPr>");
    }

    xml.push_str(&format!(
        "<w:t xml:space=\"preserve\">{}</w:t>",
        xml_escape(&run.text)
    ));
    xml.push_str("</w:r>");
}

fn write_table(xml: &mut String, table: &Table) {
    xml.push_str("\n    <w:tbl>");
    if table.properties.width_twips.is_some() || table.properties.alignment.is_some() {
        xml.push_str("<w:tblPr>");
        if let Some(w) = table.properties.width_twips {
            xml.push_str(&format!(
                "<w:tblW w:w=\"{w}\" w:type=\"dxa\"/>"
            ));
        }
        if let Some(ref a) = table.properties.alignment {
            xml.push_str(&format!("<w:jc w:val=\"{}\"/>", xml_escape(a)));
        }
        xml.push_str("</w:tblPr>");
    }
    for row in &table.rows {
        xml.push_str("<w:tr>");
        if let Some(h) = row.height_twips {
            xml.push_str(&format!(
                "<w:trPr><w:trHeight w:val=\"{h}\"/></w:trPr>"
            ));
        }
        for cell in &row.cells {
            xml.push_str("<w:tc>");
            if cell.properties.width_twips.is_some() || cell.properties.shading_color.is_some() {
                xml.push_str("<w:tcPr>");
                if let Some(w) = cell.properties.width_twips {
                    xml.push_str(&format!(
                        "<w:tcW w:w=\"{w}\" w:type=\"dxa\"/>"
                    ));
                }
                if let Some(ref sc) = cell.properties.shading_color {
                    xml.push_str(&format!(
                        "<w:shd w:fill=\"{}\"/>",
                        xml_escape(sc)
                    ));
                }
                xml.push_str("</w:tcPr>");
            }
            for p in &cell.paragraphs {
                write_paragraph(xml, p);
            }
            xml.push_str("</w:tc>");
        }
        xml.push_str("</w:tr>");
    }
    xml.push_str("</w:tbl>");
}

fn styles_xml(styles: &[StyleDefinition]) -> String {
    let mut xml = String::with_capacity(2048);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#);

    for sd in styles {
        let stype = match sd.style_type {
            StyleType::Paragraph => "paragraph",
            StyleType::Character => "character",
            StyleType::Table => "table",
            StyleType::Numbering => "numbering",
        };
        xml.push_str(&format!(
            "\n  <w:style w:type=\"{}\" w:styleId=\"{}\">",
            stype,
            xml_escape(&sd.style_id)
        ));
        if let Some(ref name) = sd.name {
            xml.push_str(&format!("<w:name w:val=\"{}\"/>", xml_escape(name)));
        }
        if let Some(ref based_on) = sd.based_on {
            xml.push_str(&format!(
                "<w:basedOn w:val=\"{}\"/>",
                xml_escape(based_on)
            ));
        }
        if has_run_style(&sd.run_style) {
            xml.push_str("<w:rPr>");
            write_run_style_xml(&mut xml, &sd.run_style);
            xml.push_str("</w:rPr>");
        }
        xml.push_str("</w:style>");
    }

    xml.push_str("\n</w:styles>");
    xml
}

fn write_run_style_xml(xml: &mut String, style: &RunStyle) {
    if style.bold {
        xml.push_str("<w:b/>");
    }
    if style.italic {
        xml.push_str("<w:i/>");
    }
    if style.underline {
        xml.push_str("<w:u w:val=\"single\"/>");
    }
    if let Some(ref font) = style.font_family {
        xml.push_str(&format!(
            "<w:rFonts w:ascii=\"{}\" w:hAnsi=\"{}\"/>",
            xml_escape(font),
            xml_escape(font)
        ));
    }
    if let Some(size) = style.font_size_pt {
        let half_pts = (size * 2.0) as u32;
        xml.push_str(&format!("<w:sz w:val=\"{half_pts}\"/>"));
    }
    if let Some(ref color) = style.color {
        xml.push_str(&format!("<w:color w:val=\"{}\"/>", xml_escape(color)));
    }
}

fn has_paragraph_properties(props: &ParagraphProperties) -> bool {
    props.style_id.is_some()
        || props.alignment.is_some()
        || props.indent_left_twips.is_some()
        || props.indent_right_twips.is_some()
        || props.indent_first_line_twips.is_some()
        || props.spacing_before_twips.is_some()
        || props.spacing_after_twips.is_some()
        || props.line_spacing_twips.is_some()
        || props.numbering.is_some()
}

fn has_run_style(style: &RunStyle) -> bool {
    style.bold
        || style.italic
        || style.underline
        || style.strikethrough
        || style.font_family.is_some()
        || style.font_size_pt.is_some()
        || style.color.is_some()
        || style.highlight.is_some()
        || style.vertical_align.is_some()
}

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
