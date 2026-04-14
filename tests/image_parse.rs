use std::io::{Cursor, Write};

use sofdocs_core::{parse_docx, render_to_html};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

fn make_docx_with_image() -> Vec<u8> {
    let buf = Vec::new();
    let cursor = Cursor::new(buf);
    let mut zip = ZipWriter::new(cursor);
    let options = SimpleFileOptions::default();

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
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

    // Fake 1x1 PNG (minimal valid PNG)
    let fake_png: Vec<u8> = vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, // RGB, no interlace
        0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
        0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01,
        0xE2, 0x21, 0xBC, 0x33,
        0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND
        0xAE, 0x42, 0x60, 0x82,
    ];
    zip.start_file("word/media/image1.png", options).unwrap();
    zip.write_all(&fake_png).unwrap();

    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Text before image</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="914400"/>
            <wp:docPr id="1" name="Picture 1" descr="Test image"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:blipFill>
                    <a:blip r:embed="rId2"/>
                  </pic:blipFill>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Text after image</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(document_xml.as_bytes()).unwrap();

    let cursor = zip.finish().unwrap();
    cursor.into_inner()
}

#[test]
fn parse_inline_image() {
    let bytes = make_docx_with_image();
    let doc = parse_docx(&bytes).unwrap();

    // Should have 3 paragraphs
    assert_eq!(doc.body.paragraphs.len(), 3);

    // First paragraph: text
    assert_eq!(doc.body.paragraphs[0].runs[0].text, "Text before image");
    assert!(doc.body.paragraphs[0].runs[0].image.is_none());

    // Second paragraph: image run
    assert_eq!(doc.body.paragraphs[1].runs.len(), 1);
    let img_run = &doc.body.paragraphs[1].runs[0];
    assert!(img_run.image.is_some());
    let img = img_run.image.as_ref().unwrap();
    assert_eq!(img.r_id, "rId2");
    assert_eq!(img.width_emu, 914400); // 1 inch = 914400 EMU
    assert_eq!(img.height_emu, 914400);
    assert_eq!(img.description, Some("Test image".to_string()));

    // Third paragraph: text
    assert_eq!(doc.body.paragraphs[2].runs[0].text, "Text after image");

    // Images extracted from ZIP
    assert_eq!(doc.images.len(), 1);
    assert_eq!(doc.images[0].r_id, "rId2");
    assert_eq!(doc.images[0].content_type, "image/png");
    assert!(!doc.images[0].data.is_empty());
}

#[test]
fn render_image_as_data_uri() {
    let bytes = make_docx_with_image();
    let doc = parse_docx(&bytes).unwrap();
    let html = render_to_html(&doc);

    assert!(html.contains("data:image/png;base64,"));
    assert!(html.contains("width:96px")); // 914400 / 9525 = 96
    assert!(html.contains("height:96px"));
    assert!(html.contains("Text before image"));
    assert!(html.contains("Text after image"));
}
