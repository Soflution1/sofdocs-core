use std::io::{Cursor, Write};

use sofdocs_core::parse_docx;
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

/// Build a minimal valid .docx in memory (ZIP with word/document.xml and [Content_Types].xml).
fn make_test_docx(document_xml: &str) -> Vec<u8> {
    let buf = Vec::new();
    let cursor = Cursor::new(buf);
    let mut zip = ZipWriter::new(cursor);
    let options = SimpleFileOptions::default();

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(document_xml.as_bytes()).unwrap();

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

#[test]
fn parse_simple_paragraph() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello, SofDocs!</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let bytes = make_test_docx(xml);
    let doc = parse_docx(&bytes).expect("should parse");

    assert_eq!(doc.body.paragraphs.len(), 1);
    assert_eq!(doc.body.paragraphs[0].runs.len(), 1);
    assert_eq!(doc.body.paragraphs[0].runs[0].text, "Hello, SofDocs!");
}

#[test]
fn parse_bold_italic_run() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:i/>
          <w:sz w:val="28"/>
          <w:color w:val="FF0000"/>
        </w:rPr>
        <w:t>Bold Italic Red</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let bytes = make_test_docx(xml);
    let doc = parse_docx(&bytes).expect("should parse");

    let run = &doc.body.paragraphs[0].runs[0];
    assert!(run.style.bold);
    assert!(run.style.italic);
    assert!(!run.style.underline);
    assert_eq!(run.style.font_size_pt, Some(14.0));
    assert_eq!(run.style.color, Some("FF0000".to_string()));
    assert_eq!(run.text, "Bold Italic Red");
}

#[test]
fn parse_multiple_paragraphs() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Title</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>First paragraph text.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Second paragraph.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let bytes = make_test_docx(xml);
    let doc = parse_docx(&bytes).expect("should parse");

    assert_eq!(doc.body.paragraphs.len(), 3);
    assert_eq!(doc.body.paragraphs[0].properties.heading_level, 1);
    assert_eq!(doc.body.paragraphs[0].runs[0].text, "Title");
    assert_eq!(doc.body.paragraphs[1].properties.heading_level, 0);
    assert_eq!(doc.word_count(), 6);
}

#[test]
fn parse_table() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>A1</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>B1</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>A2</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>B2</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#;

    let bytes = make_test_docx(xml);
    let doc = parse_docx(&bytes).expect("should parse");

    assert_eq!(doc.body.tables.len(), 1);
    assert_eq!(doc.body.tables[0].rows.len(), 2);
    assert_eq!(doc.body.tables[0].rows[0].cells.len(), 2);
    assert_eq!(doc.body.tables[0].rows[0].cells[0].paragraphs[0].runs[0].text, "A1");
    assert_eq!(doc.body.tables[0].rows[1].cells[1].paragraphs[0].runs[0].text, "B2");
}

#[test]
fn parse_alignment_and_font() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Arial"/>
          <w:u/>
        </w:rPr>
        <w:t>Centered underlined Arial</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let bytes = make_test_docx(xml);
    let doc = parse_docx(&bytes).expect("should parse");

    let para = &doc.body.paragraphs[0];
    assert_eq!(para.properties.alignment, Some("center".to_string()));

    let run = &para.runs[0];
    assert!(run.style.underline);
    assert_eq!(run.style.font_family, Some("Arial".to_string()));
}

#[test]
fn render_to_html_basic() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>Hello</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let bytes = make_test_docx(xml);
    let doc = parse_docx(&bytes).expect("should parse");
    let html = sofdocs_core::render_to_html(&doc);

    assert!(html.contains("font-weight:bold"));
    assert!(html.contains("Hello"));
    assert!(html.contains("<p data-para="));
}

#[test]
fn to_plain_text() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Line one</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Line two</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let bytes = make_test_docx(xml);
    let doc = parse_docx(&bytes).expect("should parse");

    let text = doc.to_plain_text();
    assert!(text.contains("Line one"));
    assert!(text.contains("Line two"));
    assert_eq!(doc.word_count(), 4);
    assert_eq!(doc.paragraph_count(), 2);
}
