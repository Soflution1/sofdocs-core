use super::model::*;

/// Represents a position in the document: paragraph index + character offset within that paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct DocPosition {
    pub paragraph: usize,
    pub offset: usize,
}

/// A selection range within the document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct DocSelection {
    pub start: DocPosition,
    pub end: DocPosition,
}

impl DocSelection {
    pub fn is_collapsed(&self) -> bool {
        self.start == self.end
    }

    /// Ensure start <= end.
    pub fn normalized(&self) -> Self {
        if self.start.paragraph > self.end.paragraph
            || (self.start.paragraph == self.end.paragraph && self.start.offset > self.end.offset)
        {
            Self {
                start: self.end,
                end: self.start,
            }
        } else {
            *self
        }
    }
}

/// An undoable edit operation.
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub enum EditOp {
    InsertText {
        position: DocPosition,
        text: String,
    },
    DeleteRange {
        selection: DocSelection,
        deleted_content: Vec<Paragraph>,
    },
    ApplyStyle {
        selection: DocSelection,
        style_change: StyleChange,
        previous_runs: Vec<(usize, Vec<Run>)>,
    },
    SetAlignment {
        paragraph: usize,
        new_alignment: Option<String>,
        old_alignment: Option<String>,
    },
    SplitParagraph {
        position: DocPosition,
    },
}

#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub enum StyleChange {
    ToggleBold,
    ToggleItalic,
    ToggleUnderline,
    ToggleStrikethrough,
    ToggleSuperscript,
    ToggleSubscript,
    SetFontFamily(String),
    SetFontSize(f32),
    SetColor(String),
    SetHighlight(String),
    ClearFormatting,
}

/// Undo/redo stack holding edit operations.
pub struct UndoStack {
    undo: Vec<EditOp>,
    redo: Vec<EditOp>,
}

impl UndoStack {
    pub fn new() -> Self {
        Self {
            undo: Vec::new(),
            redo: Vec::new(),
        }
    }

    pub fn push(&mut self, op: EditOp) {
        self.undo.push(op);
        self.redo.clear();
    }

    pub fn can_undo(&self) -> bool {
        !self.undo.is_empty()
    }

    pub fn can_redo(&self) -> bool {
        !self.redo.is_empty()
    }

    pub fn pop_undo(&mut self) -> Option<EditOp> {
        self.undo.pop()
    }

    pub fn push_redo(&mut self, op: EditOp) {
        self.redo.push(op);
    }

    pub fn pop_redo(&mut self) -> Option<EditOp> {
        self.redo.pop()
    }
}

impl Default for UndoStack {
    fn default() -> Self {
        Self::new()
    }
}

/// Get the total character count in a paragraph (sum of all run texts).
fn paragraph_char_count(para: &Paragraph) -> usize {
    para.runs.iter().map(|r| r.text.len()).sum()
}

/// Find which run contains the given character offset, and the offset within that run.
fn find_run_at_offset(para: &Paragraph, offset: usize) -> (usize, usize) {
    let mut remaining = offset;
    for (i, run) in para.runs.iter().enumerate() {
        if remaining <= run.text.len() {
            return (i, remaining);
        }
        remaining -= run.text.len();
    }
    let last = para.runs.len().saturating_sub(1);
    let last_len = para.runs.last().map(|r| r.text.len()).unwrap_or(0);
    (last, last_len)
}

/// Insert text at a position. Returns the end position after insertion.
pub fn insert_text(doc: &mut Document, pos: DocPosition, text: &str) -> DocPosition {
    if pos.paragraph >= doc.body.paragraphs.len() {
        doc.body.paragraphs.push(Paragraph::default());
    }
    let para = &mut doc.body.paragraphs[pos.paragraph];

    if para.runs.is_empty() {
        para.runs.push(Run {
            text: text.to_string(),
            ..Default::default()
        });
        return DocPosition {
            paragraph: pos.paragraph,
            offset: text.len(),
        };
    }

    let (run_idx, run_offset) = find_run_at_offset(para, pos.offset);
    para.runs[run_idx].text.insert_str(run_offset, text);

    DocPosition {
        paragraph: pos.paragraph,
        offset: pos.offset + text.len(),
    }
}

/// Delete the content in a selection. Returns the collapsed position.
pub fn delete_range(doc: &mut Document, sel: DocSelection) -> (DocPosition, Vec<Paragraph>) {
    let sel = sel.normalized();

    if sel.is_collapsed() {
        return (sel.start, Vec::new());
    }

    if sel.start.paragraph == sel.end.paragraph {
        let para = &mut doc.body.paragraphs[sel.start.paragraph];
        let deleted = extract_paragraph_range(para, sel.start.offset, sel.end.offset);
        delete_chars_in_paragraph(para, sel.start.offset, sel.end.offset);
        return (sel.start, vec![deleted]);
    }

    let mut deleted_paras = Vec::new();

    // Save the deleted content from the first paragraph (from start.offset to end)
    let first_para = &doc.body.paragraphs[sel.start.paragraph];
    let first_len = paragraph_char_count(first_para);
    deleted_paras.push(extract_paragraph_range(
        first_para,
        sel.start.offset,
        first_len,
    ));

    // Save full middle paragraphs
    for i in (sel.start.paragraph + 1)..sel.end.paragraph {
        deleted_paras.push(doc.body.paragraphs[i].clone());
    }

    // Save deleted content from last paragraph
    let last_para = &doc.body.paragraphs[sel.end.paragraph];
    deleted_paras.push(extract_paragraph_range(last_para, 0, sel.end.offset));

    // Now perform the deletion: keep the tail of the last paragraph
    let last_para = &doc.body.paragraphs[sel.end.paragraph];
    let last_len = paragraph_char_count(last_para);
    let tail_runs = extract_runs_from_offset(last_para, sel.end.offset, last_len);

    // Truncate the first paragraph at start.offset
    delete_chars_in_paragraph(
        &mut doc.body.paragraphs[sel.start.paragraph],
        sel.start.offset,
        first_len,
    );

    // Append tail of last paragraph to first
    for run in tail_runs {
        if !run.text.is_empty() {
            doc.body.paragraphs[sel.start.paragraph].runs.push(run);
        }
    }

    // Remove the middle + last paragraphs
    let remove_count = sel.end.paragraph - sel.start.paragraph;
    for _ in 0..remove_count {
        doc.body.paragraphs.remove(sel.start.paragraph + 1);
    }

    (sel.start, deleted_paras)
}

/// Split a paragraph at a position (Enter key).
pub fn split_paragraph(doc: &mut Document, pos: DocPosition) {
    if pos.paragraph >= doc.body.paragraphs.len() {
        doc.body.paragraphs.push(Paragraph::default());
        return;
    }

    let para = &doc.body.paragraphs[pos.paragraph];
    let total_len = paragraph_char_count(para);
    let tail_runs = extract_runs_from_offset(para, pos.offset, total_len);

    // Truncate current paragraph
    delete_chars_in_paragraph(
        &mut doc.body.paragraphs[pos.paragraph],
        pos.offset,
        total_len,
    );

    let new_para = Paragraph {
        properties: doc.body.paragraphs[pos.paragraph].properties.clone(),
        runs: tail_runs,
        bookmarks: Vec::new(),
    };

    doc.body.paragraphs.insert(pos.paragraph + 1, new_para);
}

/// Apply a style change to a selection.
pub fn apply_style(
    doc: &mut Document,
    sel: DocSelection,
    change: &StyleChange,
) -> Vec<(usize, Vec<Run>)> {
    let sel = sel.normalized();
    let mut previous_runs = Vec::new();

    let start_para = sel.start.paragraph;
    let end_para = sel.end.paragraph.min(doc.body.paragraphs.len().saturating_sub(1));

    for pi in start_para..=end_para {
        let para = &doc.body.paragraphs[pi];
        previous_runs.push((pi, para.runs.clone()));

        let para_len = paragraph_char_count(para);
        let start_offset = if pi == start_para { sel.start.offset } else { 0 };
        let end_offset = if pi == end_para {
            sel.end.offset
        } else {
            para_len
        };

        split_runs_at_boundaries(&mut doc.body.paragraphs[pi], start_offset, end_offset);

        let para = &mut doc.body.paragraphs[pi];
        let mut char_pos = 0;
        for run in &mut para.runs {
            let run_end = char_pos + run.text.len();
            if char_pos >= start_offset && run_end <= end_offset && !run.text.is_empty() {
                apply_style_change(&mut run.style, change);
            }
            char_pos = run_end;
        }
    }

    previous_runs
}

/// Set paragraph alignment.
pub fn set_alignment(
    doc: &mut Document,
    paragraph: usize,
    alignment: Option<String>,
) -> Option<String> {
    if paragraph >= doc.body.paragraphs.len() {
        return None;
    }
    let old = doc.body.paragraphs[paragraph].properties.alignment.clone();
    doc.body.paragraphs[paragraph].properties.alignment = alignment;
    old
}

fn apply_style_change(style: &mut RunStyle, change: &StyleChange) {
    match change {
        StyleChange::ToggleBold => style.bold = !style.bold,
        StyleChange::ToggleItalic => style.italic = !style.italic,
        StyleChange::ToggleUnderline => style.underline = !style.underline,
        StyleChange::ToggleStrikethrough => style.strikethrough = !style.strikethrough,
        StyleChange::ToggleSuperscript => {
            style.vertical_align = match style.vertical_align {
                Some(VerticalAlign::Superscript) => None,
                _ => Some(VerticalAlign::Superscript),
            };
        }
        StyleChange::ToggleSubscript => {
            style.vertical_align = match style.vertical_align {
                Some(VerticalAlign::Subscript) => None,
                _ => Some(VerticalAlign::Subscript),
            };
        }
        StyleChange::SetFontFamily(f) => style.font_family = Some(f.clone()),
        StyleChange::SetFontSize(s) => style.font_size_pt = Some(*s),
        StyleChange::SetColor(c) => style.color = Some(c.clone()),
        StyleChange::SetHighlight(h) => style.highlight = Some(h.clone()),
        StyleChange::ClearFormatting => {
            *style = RunStyle::default();
        }
    }
}

/// Split runs at character boundaries so that the range [start..end] aligns with run boundaries.
fn split_runs_at_boundaries(para: &mut Paragraph, start: usize, end: usize) {
    split_run_at(para, end);
    split_run_at(para, start);
}

fn split_run_at(para: &mut Paragraph, offset: usize) {
    let mut char_pos = 0;
    for i in 0..para.runs.len() {
        let run_len = para.runs[i].text.len();
        if char_pos == offset || char_pos + run_len <= offset {
            char_pos += run_len;
            continue;
        }
        if offset > char_pos && offset < char_pos + run_len {
            let split_at = offset - char_pos;
            let tail_text = para.runs[i].text[split_at..].to_string();
            para.runs[i].text.truncate(split_at);
            let new_run = Run {
                text: tail_text,
                style: para.runs[i].style.clone(),
                image: para.runs[i].image.clone(),
                hyperlink: para.runs[i].hyperlink.clone(),
            };
            para.runs.insert(i + 1, new_run);
            return;
        }
        char_pos += run_len;
    }
}

fn delete_chars_in_paragraph(para: &mut Paragraph, start: usize, end: usize) {
    split_runs_at_boundaries(para, start, end);

    let mut char_pos = 0;
    para.runs.retain(|run| {
        let run_end = char_pos + run.text.len();
        let keep = run_end <= start || char_pos >= end;
        char_pos = run_end;
        keep
    });
}

fn extract_paragraph_range(para: &Paragraph, start: usize, end: usize) -> Paragraph {
    let runs = extract_runs_from_offset(para, start, end);
    Paragraph {
        properties: para.properties.clone(),
        runs,
        bookmarks: Vec::new(),
    }
}

fn extract_runs_from_offset(para: &Paragraph, start: usize, end: usize) -> Vec<Run> {
    let mut result = Vec::new();
    let mut char_pos = 0;

    for run in &para.runs {
        let run_end = char_pos + run.text.len();

        if run_end <= start || char_pos >= end {
            char_pos = run_end;
            continue;
        }

        let slice_start = if char_pos < start {
            start - char_pos
        } else {
            0
        };
        let slice_end = if run_end > end {
            end - char_pos
        } else {
            run.text.len()
        };

        let text = run.text[slice_start..slice_end].to_string();
        if !text.is_empty() {
            result.push(Run {
                text,
                style: run.style.clone(),
                image: run.image.clone(),
                hyperlink: run.hyperlink.clone(),
            });
        }

        char_pos = run_end;
    }

    result
}

/// Perform an undo operation. Returns the inverse op for redo.
pub fn undo(doc: &mut Document, undo_stack: &mut UndoStack) -> bool {
    let op = match undo_stack.pop_undo() {
        Some(op) => op,
        None => return false,
    };

    let redo_op = execute_inverse(doc, &op);
    undo_stack.push_redo(redo_op);
    true
}

/// Perform a redo operation.
pub fn redo(doc: &mut Document, undo_stack: &mut UndoStack) -> bool {
    let op = match undo_stack.pop_redo() {
        Some(op) => op,
        None => return false,
    };

    let inverse = execute_inverse(doc, &op);
    undo_stack.undo.push(inverse);
    true
}

/// Set paragraph indentation (values in twips).
pub fn set_indent(doc: &mut Document, paragraph: usize, left: i32, right: i32, first_line: i32) {
    if paragraph >= doc.body.paragraphs.len() { return; }
    let props = &mut doc.body.paragraphs[paragraph].properties;
    props.indent_left_twips = if left != 0 { Some(left) } else { None };
    props.indent_right_twips = if right != 0 { Some(right) } else { None };
    props.indent_first_line_twips = if first_line != 0 { Some(first_line) } else { None };
}

/// Set paragraph spacing (values in twips, line in 240ths of a line).
pub fn set_spacing(doc: &mut Document, paragraph: usize, before: u32, after: u32, line: u32) {
    if paragraph >= doc.body.paragraphs.len() { return; }
    let props = &mut doc.body.paragraphs[paragraph].properties;
    props.spacing_before_twips = if before != 0 { Some(before) } else { None };
    props.spacing_after_twips = if after != 0 { Some(after) } else { None };
    props.line_spacing_twips = if line != 0 { Some(line) } else { None };
}

/// Set heading level (0 = normal, 1-6 = heading).
pub fn set_heading_level(doc: &mut Document, paragraph: usize, level: u8) {
    if paragraph >= doc.body.paragraphs.len() { return; }
    doc.body.paragraphs[paragraph].properties.heading_level = level.min(6);
    if level > 0 {
        doc.body.paragraphs[paragraph].properties.style_id = None;
    }
}

/// Toggle list on a paragraph. list_type: "bullet" or "decimal".
pub fn toggle_list(doc: &mut Document, paragraph: usize, list_type: &str) {
    if paragraph >= doc.body.paragraphs.len() { return; }
    let props = &mut doc.body.paragraphs[paragraph].properties;

    if props.numbering.is_some() {
        props.numbering = None;
        return;
    }

    let num_fmt = list_type.to_string();
    let abs_id = doc.numbering_definitions.len() as u32;
    let num_id = abs_id + 1;

    let existing = doc.numbering_definitions.iter().find(|d| {
        d.levels.first().map(|l| l.num_fmt.as_str()) == Some(list_type)
    });

    let actual_num_id = if let Some(def) = existing {
        def.num_ids.first().copied().unwrap_or(def.abstract_num_id + 1)
    } else {
        doc.numbering_definitions.push(NumberingDefinition {
            abstract_num_id: abs_id,
            num_ids: vec![num_id],
            levels: vec![NumberingLevel {
                level: 0,
                num_fmt,
                lvl_text: if list_type == "bullet" { "•".to_string() } else { "%1.".to_string() },
                start: 1,
            }],
        });
        num_id
    };

    props.numbering = Some(NumberingInfo { num_id: actual_num_id, level: 0 });
}

/// Find all occurrences of a text query. Returns Vec of (para_idx, char_offset, length).
pub fn find_text(doc: &Document, query: &str) -> Vec<(usize, usize, usize)> {
    let mut results = Vec::new();
    if query.is_empty() { return results; }

    let query_lower = query.to_lowercase();
    for (pi, para) in doc.body.paragraphs.iter().enumerate() {
        let full_text: String = para.runs.iter().map(|r| r.text.as_str()).collect();
        let full_lower = full_text.to_lowercase();
        let mut search_from = 0;
        while let Some(pos) = full_lower[search_from..].find(&query_lower) {
            let abs_pos = search_from + pos;
            results.push((pi, abs_pos, query.len()));
            search_from = abs_pos + 1;
        }
    }
    results
}

/// Replace text at a specific location.
pub fn replace_text_at(doc: &mut Document, para: usize, offset: usize, len: usize, replacement: &str) {
    if para >= doc.body.paragraphs.len() { return; }
    let sel = DocSelection {
        start: DocPosition { paragraph: para, offset },
        end: DocPosition { paragraph: para, offset: offset + len },
    };
    delete_range(doc, sel);
    insert_text(doc, DocPosition { paragraph: para, offset }, replacement);
}

/// Replace all occurrences. Returns count of replacements.
pub fn replace_all(doc: &mut Document, query: &str, replacement: &str) -> usize {
    let mut count = 0;
    loop {
        let matches = find_text(doc, query);
        if matches.is_empty() { break; }
        let (para, offset, len) = matches[0];
        replace_text_at(doc, para, offset, len, replacement);
        count += 1;
    }
    count
}

/// Insert a table after a given paragraph.
pub fn insert_table(doc: &mut Document, _after_paragraph: usize, rows: usize, cols: usize) {
    let mut table_rows = Vec::with_capacity(rows);
    for _ in 0..rows {
        let mut cells = Vec::with_capacity(cols);
        for _ in 0..cols {
            cells.push(TableCell {
                paragraphs: vec![Paragraph::default()],
                ..Default::default()
            });
        }
        table_rows.push(TableRow { cells, ..Default::default() });
    }
    doc.body.tables.push(Table {
        rows: table_rows,
        ..Default::default()
    });
}

/// Insert a page break before a paragraph.
pub fn insert_page_break(doc: &mut Document, paragraph: usize) {
    if paragraph >= doc.body.paragraphs.len() {
        doc.body.paragraphs.push(Paragraph {
            properties: ParagraphProperties { page_break_before: true, ..Default::default() },
            ..Default::default()
        });
        return;
    }
    let pos = DocPosition { paragraph, offset: 0 };
    split_paragraph(doc, pos);
    doc.body.paragraphs[paragraph + 1].properties.page_break_before = true;
}

/// Insert a hyperlink on a run range.
pub fn insert_hyperlink(doc: &mut Document, para: usize, start_offset: usize, end_offset: usize, url: &str) {
    if para >= doc.body.paragraphs.len() { return; }
    split_runs_at_boundaries(&mut doc.body.paragraphs[para], start_offset, end_offset);
    let paragraph = &mut doc.body.paragraphs[para];
    let mut char_pos = 0;
    for run in &mut paragraph.runs {
        let run_end = char_pos + run.text.len();
        if char_pos >= start_offset && run_end <= end_offset && !run.text.is_empty() {
            run.hyperlink = Some(HyperlinkInfo { url: url.to_string(), tooltip: None });
        }
        char_pos = run_end;
    }
}

/// Insert a bookmark at a position.
pub fn insert_bookmark(doc: &mut Document, para: usize, _offset: usize, name: &str) {
    if para >= doc.body.paragraphs.len() { return; }
    let max_id = doc.body.paragraphs.iter()
        .flat_map(|p| p.bookmarks.iter())
        .map(|b| b.id)
        .max()
        .unwrap_or(0);
    doc.body.paragraphs[para].bookmarks.push(Bookmark {
        id: max_id + 1,
        name: name.to_string(),
    });
}

fn execute_inverse(doc: &mut Document, op: &EditOp) -> EditOp {
    match op {
        EditOp::InsertText { position, text } => {
            let end = DocPosition {
                paragraph: position.paragraph,
                offset: position.offset + text.len(),
            };
            let sel = DocSelection {
                start: *position,
                end,
            };
            let (_, deleted) = delete_range(doc, sel);
            EditOp::DeleteRange {
                selection: sel,
                deleted_content: deleted,
            }
        }
        EditOp::DeleteRange {
            selection,
            deleted_content,
        } => {
            let sel = selection.normalized();
            let mut restored_text = String::new();
            for (i, para) in deleted_content.iter().enumerate() {
                if i > 0 {
                    restored_text.push('\n');
                }
                for run in &para.runs {
                    restored_text.push_str(&run.text);
                }
            }
            insert_text(doc, sel.start, &restored_text);
            EditOp::InsertText {
                position: sel.start,
                text: restored_text,
            }
        }
        EditOp::ApplyStyle {
            selection,
            style_change,
            previous_runs,
        } => {
            let current_runs: Vec<(usize, Vec<Run>)> = previous_runs
                .iter()
                .map(|(pi, _)| {
                    let runs = doc.body.paragraphs[*pi].runs.clone();
                    (*pi, runs)
                })
                .collect();

            for (pi, runs) in previous_runs {
                if *pi < doc.body.paragraphs.len() {
                    doc.body.paragraphs[*pi].runs = runs.clone();
                }
            }

            EditOp::ApplyStyle {
                selection: *selection,
                style_change: style_change.clone(),
                previous_runs: current_runs,
            }
        }
        EditOp::SetAlignment {
            paragraph,
            new_alignment,
            old_alignment,
        } => {
            set_alignment(doc, *paragraph, old_alignment.clone());
            EditOp::SetAlignment {
                paragraph: *paragraph,
                new_alignment: old_alignment.clone(),
                old_alignment: new_alignment.clone(),
            }
        }
        EditOp::SplitParagraph { position } => {
            // Merge the two paragraphs back
            if position.paragraph + 1 < doc.body.paragraphs.len() {
                let next_runs = doc.body.paragraphs[position.paragraph + 1].runs.clone();
                doc.body.paragraphs.remove(position.paragraph + 1);
                doc.body.paragraphs[position.paragraph]
                    .runs
                    .extend(next_runs);
            }
            EditOp::SplitParagraph {
                position: *position,
            }
        }
    }
}
