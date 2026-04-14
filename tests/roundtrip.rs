use std::io::{Cursor, Write};

use sofdocs_core::document::editor::{self, DocPosition, DocSelection, StyleChange, UndoStack};
use sofdocs_core::document::writer::write_docx;
use sofdocs_core::{parse_docx, render_to_html};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

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
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#;

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#
            .as_bytes(),
    )
    .unwrap();

    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    zip.write_all(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#
            .as_bytes(),
    )
    .unwrap();

    zip.start_file("word/styles.xml", options).unwrap();
    zip.write_all(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:styles>"#
            .as_bytes(),
    )
    .unwrap();

    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(document_xml.as_bytes()).unwrap();

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

#[test]
fn roundtrip_simple_paragraph() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello roundtrip!</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let original = parse_docx(&make_test_docx(xml)).unwrap();
    let saved_bytes = write_docx(&original).unwrap();
    let reparsed = parse_docx(&saved_bytes).unwrap();

    assert_eq!(original.body.paragraphs.len(), reparsed.body.paragraphs.len());
    assert_eq!(
        original.body.paragraphs[0].runs[0].text,
        reparsed.body.paragraphs[0].runs[0].text
    );
    assert_eq!(original.to_plain_text(), reparsed.to_plain_text());
}

#[test]
fn roundtrip_styled_text() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/><w:i/><w:sz w:val="28"/><w:color w:val="0000FF"/></w:rPr>
        <w:t>Styled text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let original = parse_docx(&make_test_docx(xml)).unwrap();
    let saved_bytes = write_docx(&original).unwrap();
    let reparsed = parse_docx(&saved_bytes).unwrap();

    let orig_run = &original.body.paragraphs[0].runs[0];
    let repr_run = &reparsed.body.paragraphs[0].runs[0];

    assert_eq!(orig_run.text, repr_run.text);
    assert_eq!(orig_run.style.bold, repr_run.style.bold);
    assert_eq!(orig_run.style.italic, repr_run.style.italic);
    assert_eq!(orig_run.style.font_size_pt, repr_run.style.font_size_pt);
    assert_eq!(orig_run.style.color, repr_run.style.color);
}

#[test]
fn roundtrip_edit_then_save() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Hello world</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let mut doc = parse_docx(&make_test_docx(xml)).unwrap();

    // Insert text
    editor::insert_text(&mut doc, DocPosition { paragraph: 0, offset: 5 }, " beautiful");
    assert_eq!(doc.body.paragraphs[0].runs[0].text, "Hello beautiful world");

    // Save and reparse
    let saved_bytes = write_docx(&doc).unwrap();
    let reparsed = parse_docx(&saved_bytes).unwrap();
    assert_eq!(reparsed.to_plain_text().trim(), "Hello beautiful world");
}

#[test]
fn roundtrip_apply_style() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Bold this word</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let mut doc = parse_docx(&make_test_docx(xml)).unwrap();

    // Make "this" bold (chars 5..9)
    let sel = DocSelection {
        start: DocPosition { paragraph: 0, offset: 5 },
        end: DocPosition { paragraph: 0, offset: 9 },
    };
    editor::apply_style(&mut doc, sel, &StyleChange::ToggleBold);

    let saved_bytes = write_docx(&doc).unwrap();
    let reparsed = parse_docx(&saved_bytes).unwrap();

    let plain = reparsed.to_plain_text();
    assert!(plain.contains("Bold this word"));

    // Verify there's a bold run in the reparsed doc
    let has_bold = reparsed.body.paragraphs[0]
        .runs
        .iter()
        .any(|r| r.style.bold);
    assert!(has_bold);
}

#[test]
fn roundtrip_undo_redo() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Original</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let mut doc = parse_docx(&make_test_docx(xml)).unwrap();
    let mut undo_stack = UndoStack::new();

    // Insert text
    let pos = DocPosition { paragraph: 0, offset: 8 };
    editor::insert_text(&mut doc, pos, " text");
    undo_stack.push(sofdocs_core::document::editor::EditOp::InsertText {
        position: pos,
        text: " text".to_string(),
    });

    assert_eq!(doc.to_plain_text().trim(), "Original text");

    // Undo
    editor::undo(&mut doc, &mut undo_stack);
    assert_eq!(doc.to_plain_text().trim(), "Original");

    // Redo
    editor::redo(&mut doc, &mut undo_stack);
    assert_eq!(doc.to_plain_text().trim(), "Original text");
}

#[test]
fn roundtrip_multiple_paragraphs() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r><w:t>Centered Title</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Normal paragraph.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let original = parse_docx(&make_test_docx(xml)).unwrap();
    let saved_bytes = write_docx(&original).unwrap();
    let reparsed = parse_docx(&saved_bytes).unwrap();

    assert_eq!(reparsed.body.paragraphs.len(), 2);
    assert_eq!(
        reparsed.body.paragraphs[0].properties.alignment,
        Some("center".to_string())
    );
    assert_eq!(reparsed.body.paragraphs[1].runs[0].text, "Normal paragraph.");
}

#[test]
fn roundtrip_html_consistency() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Test HTML</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let original = parse_docx(&make_test_docx(xml)).unwrap();
    let html1 = render_to_html(&original);

    let saved_bytes = write_docx(&original).unwrap();
    let reparsed = parse_docx(&saved_bytes).unwrap();
    let html2 = render_to_html(&reparsed);

    assert_eq!(html1, html2);
}

#[test]
fn roundtrip_split_paragraph() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>First line second line</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let mut doc = parse_docx(&make_test_docx(xml)).unwrap();

    // Split at "First line " (offset 11)
    editor::split_paragraph(&mut doc, DocPosition { paragraph: 0, offset: 11 });
    assert_eq!(doc.body.paragraphs.len(), 2);

    let saved_bytes = write_docx(&doc).unwrap();
    let reparsed = parse_docx(&saved_bytes).unwrap();
    assert_eq!(reparsed.body.paragraphs.len(), 2);
    assert_eq!(reparsed.body.paragraphs[0].runs[0].text, "First line ");
    assert_eq!(reparsed.body.paragraphs[1].runs[0].text, "second line");
}

#[test]
fn roundtrip_image_preservation() {
    use sofdocs_core::document::model::{ImageEntry, InlineImage};

    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Before</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

    let mut doc = parse_docx(&make_test_docx(xml)).unwrap();

    // Add a fake image entry
    let fake_png = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0xFF];
    doc.images.push(ImageEntry {
        r_id: "rId10".to_string(),
        path: "media/image1.png".to_string(),
        content_type: "image/png".to_string(),
        data: fake_png.clone(),
    });

    // Add an image run
    doc.body.paragraphs.push(sofdocs_core::document::model::Paragraph {
        runs: vec![sofdocs_core::document::model::Run {
            image: Some(InlineImage {
                r_id: "rId10".to_string(),
                width_emu: 914400,
                height_emu: 457200,
                description: Some("Test roundtrip img".to_string()),
                ..Default::default()
            }),
            ..Default::default()
        }],
        ..Default::default()
    });

    let saved_bytes = write_docx(&doc).unwrap();
    let reparsed = parse_docx(&saved_bytes).unwrap();

    // Image data preserved in ZIP
    assert_eq!(reparsed.images.len(), 1);
    assert_eq!(reparsed.images[0].data, fake_png);
    assert_eq!(reparsed.images[0].content_type, "image/png");

    // Image run preserved
    assert_eq!(reparsed.body.paragraphs.len(), 2);
    let img_run = &reparsed.body.paragraphs[1].runs[0];
    assert!(img_run.image.is_some());
    let img = img_run.image.as_ref().unwrap();
    assert_eq!(img.r_id, "rId10");
    assert_eq!(img.width_emu, 914400);
    assert_eq!(img.height_emu, 457200);
    assert_eq!(img.description, Some("Test roundtrip img".to_string()));
}
