use serde::{Deserialize, Serialize};

/// Root document structure parsed from a .docx file.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Document {
    pub metadata: DocumentMetadata,
    pub body: DocumentBody,
    pub styles: Vec<StyleDefinition>,
    pub numbering_definitions: Vec<NumberingDefinition>,
    pub images: Vec<ImageEntry>,
}

/// A numbering (list) definition from numbering.xml.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct NumberingDefinition {
    pub abstract_num_id: u32,
    /// The concrete numIds that reference this abstractNum (from `<w:num>` entries).
    pub num_ids: Vec<u32>,
    pub levels: Vec<NumberingLevel>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct NumberingLevel {
    pub level: u32,
    pub num_fmt: String,
    pub lvl_text: String,
    pub start: u32,
}

/// An image stored inside the docx archive.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ImageEntry {
    pub r_id: String,
    pub path: String,
    pub content_type: String,
    #[serde(skip)]
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocumentMetadata {
    pub title: Option<String>,
    pub author: Option<String>,
    pub description: Option<String>,
    pub created: Option<String>,
    pub modified: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocumentBody {
    pub paragraphs: Vec<Paragraph>,
    pub tables: Vec<Table>,
    pub headers: Vec<HeaderFooter>,
    pub footers: Vec<HeaderFooter>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct HeaderFooter {
    pub r_id: String,
    pub hf_type: String,
    pub paragraphs: Vec<Paragraph>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Paragraph {
    pub properties: ParagraphProperties,
    pub runs: Vec<Run>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ParagraphProperties {
    /// Style ID referencing a named style (e.g., "Heading1")
    pub style_id: Option<String>,
    /// Resolved heading level: 0 = body text, 1-6 = heading
    pub heading_level: u8,
    pub alignment: Option<String>,
    pub indent_left_twips: Option<i32>,
    pub indent_right_twips: Option<i32>,
    pub indent_first_line_twips: Option<i32>,
    pub spacing_before_twips: Option<u32>,
    pub spacing_after_twips: Option<u32>,
    pub line_spacing_twips: Option<u32>,
    /// Numbered or bulleted list info
    pub numbering: Option<NumberingInfo>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct NumberingInfo {
    pub num_id: u32,
    pub level: u32,
}

/// Inline image embedded in a run.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct InlineImage {
    pub r_id: String,
    pub content_type: String,
    #[serde(skip)]
    pub data: Vec<u8>,
    pub width_emu: u64,
    pub height_emu: u64,
    pub description: Option<String>,
}

/// A "run" is a contiguous span of text sharing the same formatting.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Run {
    pub text: String,
    pub style: RunStyle,
    pub image: Option<InlineImage>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct RunStyle {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_family: Option<String>,
    /// Font size in points (half-points in OOXML are converted on parse)
    pub font_size_pt: Option<f32>,
    /// Hex color without `#` prefix, e.g. "FF0000"
    pub color: Option<String>,
    pub highlight: Option<String>,
    pub vertical_align: Option<VerticalAlign>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum VerticalAlign {
    Superscript,
    Subscript,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Table {
    pub rows: Vec<TableRow>,
    pub properties: TableProperties,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableProperties {
    pub width_twips: Option<u32>,
    pub alignment: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
    pub height_twips: Option<u32>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableCell {
    pub paragraphs: Vec<Paragraph>,
    pub properties: TableCellProperties,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableCellProperties {
    pub width_twips: Option<u32>,
    pub vertical_merge: Option<String>,
    pub shading_color: Option<String>,
}

/// Named style definition from styles.xml.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct StyleDefinition {
    pub style_id: String,
    pub name: Option<String>,
    pub style_type: StyleType,
    pub based_on: Option<String>,
    pub run_style: RunStyle,
    pub paragraph_properties: ParagraphProperties,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub enum StyleType {
    #[default]
    Paragraph,
    Character,
    Table,
    Numbering,
}

impl Document {
    /// Returns the full plain text content of the document.
    pub fn to_plain_text(&self) -> String {
        let mut text = String::new();
        for paragraph in &self.body.paragraphs {
            for run in &paragraph.runs {
                text.push_str(&run.text);
            }
            text.push('\n');
        }
        for table in &self.body.tables {
            for row in &table.rows {
                for (i, cell) in row.cells.iter().enumerate() {
                    if i > 0 {
                        text.push('\t');
                    }
                    for paragraph in &cell.paragraphs {
                        for run in &paragraph.runs {
                            text.push_str(&run.text);
                        }
                    }
                }
                text.push('\n');
            }
        }
        text
    }

    /// Returns a word count.
    pub fn word_count(&self) -> usize {
        self.to_plain_text().split_whitespace().count()
    }

    /// Returns the number of paragraphs.
    pub fn paragraph_count(&self) -> usize {
        self.body.paragraphs.len()
    }
}
